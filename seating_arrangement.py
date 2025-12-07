#!/usr/bin/env python3
"""
Exam seating arrangement + attendance PDF generator (single workbook version).

This script is tailored for an input Excel workbook (default: input_data_tt.xlsx)
that contains ALL the required sheets:

    1. in_timetable
       Columns:
           Date       : exam date
           Day        : weekday (unused for allocation)
           Morning    : semicolon-separated course codes,
                        e.g. "CS249; CH426; MM304; ..."
           Evening    : same as Morning OR "NO EXAM"

    2. in_course_roll_mapping
       Columns:
           rollno         : student roll number
           register_sem   : (unused)
           schedule_sem   : (unused)
           course_code    : course code (e.g., "CS249")

    3. in_roll_name_mapping
       Columns:
           Roll           : student roll number
           Name           : student name

    4. in_room_capacity
       Columns:
           Room No.       : room number (e.g., 6101)
           Exam Capacity  : integer capacity of the room
           Block          : building/block code (e.g., B1, B2)
           other columns  : ignored

High-level behaviour (aligned with the problem statement):

1. From the timetable (in_timetable) and course-roll mapping
   (in_course_roll_mapping), we construct a "registrations" table with columns:
        date, slot, coursecode, rollno

2. For each (date, slot):
   - Find the largest course (by number of students) and allocate it first.
   - Use a greedy allocation to minimise number of rooms used per course.
   - Try to avoid placing a single course in multiple buildings when possible.
   - When multiple buildings are needed, keep rooms within a building as
     adjacent/close as possible (sorted by room number).

3. User inputs:
   - buffer: integer; for a room with capacity C, effective capacity is C - buffer
             (not below 0).
   - mode: "sparse" or "dense"
       * sparse: per-subject capacity = floor(effective / 2)
       * dense : per-subject capacity = effective

4. Clash checking:
   - For each (date, slot), if any roll number appears in more than one course,
     we report a clash to the terminal and the error log
     (but allocation still proceeds).

5. Logging:
   - Uses Python's logging module.
   - logs/execution.log : INFO and above.
   - logs/errors.txt    : ERROR and above.
   - Also logs to console.
   - try/except around main() so script does not crash abruptly.

6. Excel Output:
   - op_overall_seating_arrangement.xlsx
        Columns:
          Date, Slot, Building, Room, CourseCode,
          RollNumbers (semicolon-separated),
          Names (semicolon-separated)
   - op_seats_left.xlsx
        Columns:
          Date, Slot, Building, Room,
          EffectiveCapacityPerSubject, UsedSeats, SeatsLeft
   - Additionally, per-slot files:
        output/YYYY-MM-DD/morning/seating_arrangement.xlsx
        output/YYYY-MM-DD/evening/seating_arrangement.xlsx

7. Attendance PDFs:
   - For each (date, slot, room, course) group, generate a PDF with filename:

         YYYY_MM_DD_<SESSION>_R<ROOM>_<SUBCODE>.pdf

     Example:
         2016_05_04_Morning_R6102_PH703.pdf

   - PDFs are written under:

         <attendance-dir>/<date>/<slot>/<filename>.pdf

     where <attendance-dir> is a command-line argument (default: attendance_pdfs).

   - If a folder "photos/" exists, the script assumes photos are named:

         photos/<ROLL>.jpg

     and tries to add a small photo thumbnail next to each student row
     (if the file is missing, it simply leaves the photo cell blank).

8. Roll-name mapping:
   - If a roll number is missing in in_roll_name_mapping, its name is set to
     "Unknown Name".

Command line usage:

    python3 seating_arrangement.py \
        --input input_data_tt.xlsx \
        --buffer 5 \
        --mode sparse \
        --output-dir output \
        --attendance-dir attendance_pdfs \
        --photos-dir photos

If you run simply:

    python3 seating_arrangement.py

it will look for "input_data_tt.xlsx" in the current directory with the
above sheet structure.

IMPORTANT:
    - This script uses the "reportlab" library to generate PDFs.
      Install it once using:

          pip install reportlab

"""

import argparse
import logging
import os
import sys
import traceback
from collections import defaultdict

import pandas as pd

# PDF generation (reportlab)
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.lib.utils import ImageReader
except ImportError:  # handled at runtime
    # We will log a clear error later if user tries to generate PDFs without reportlab.
    A4 = None
    mm = 1
    canvas = None
    ImageReader = None


LOG_DIR = "logs"


def setup_logging():
    os.makedirs(LOG_DIR, exist_ok=True)

    logger = logging.getLogger("seating")
    logger.setLevel(logging.INFO)

    # Avoid duplicate handlers if main() is called more than once
    if logger.handlers:
        return logger

    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    # Execution log
    exec_handler = logging.FileHandler(os.path.join(LOG_DIR, "execution.log"))
    exec_handler.setLevel(logging.INFO)
    exec_handler.setFormatter(fmt)

    # Error log
    error_handler = logging.FileHandler(os.path.join(LOG_DIR, "errors.txt"))
    error_handler.setLevel(logging.ERROR)
    error_handler.setFormatter(fmt)

    # Console
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(fmt)

    logger.addHandler(exec_handler)
    logger.addHandler(error_handler)
    logger.addHandler(console_handler)

    return logger


def read_excel_stripped(path: str, sheet_name=None, logger: logging.Logger = None) -> pd.DataFrame:
    """
    Helper to read Excel and strip spaces from string columns.
    """
    if logger:
        logger.info("Reading Excel file: %s (sheet=%s)", path, sheet_name)
    df = pd.read_excel(path, sheet_name=sheet_name)
    for col in df.columns:
        if pd.api.types.is_string_dtype(df[col]) or df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
    return df


def find_sheet(xls: pd.ExcelFile, target: str):
    """
    Find a sheet by name, case-insensitive.
    """
    for name in xls.sheet_names:
        if name.strip().lower() == target.strip().lower():
            return name
    return None


def load_inputs_from_workbook(wb_path: str, logger: logging.Logger):
    """
    Load data from a single Excel workbook with multiple sheets:

        - in_timetable
        - in_course_roll_mapping
        - in_roll_name_mapping
        - in_room_capacity

    Returns:
        reg_df   : DataFrame with columns [date, slot, coursecode, rollno]
        class_df : DataFrame with columns [building, room, capacity]
        roll_to_name : dict mapping rollno -> name
    """
    if not os.path.exists(wb_path):
        raise FileNotFoundError(f"Input file not found: {wb_path}")

    logger.info("Opening workbook: %s", wb_path)
    xls = pd.ExcelFile(wb_path)

    # Resolve sheet names
    timetable_sheet = find_sheet(xls, "in_timetable")
    course_map_sheet = find_sheet(xls, "in_course_roll_mapping")
    roll_name_sheet = find_sheet(xls, "in_roll_name_mapping")
    room_sheet = find_sheet(xls, "in_room_capacity")

    missing_sheets = [
        name for name, found in [
            ("in_timetable", timetable_sheet),
            ("in_course_roll_mapping", course_map_sheet),
            ("in_roll_name_mapping", roll_name_sheet),
            ("in_room_capacity", room_sheet),
        ] if found is None
    ]
    if missing_sheets:
        raise ValueError(f"Missing required sheets in workbook: {missing_sheets}")

    # Read sheets
    tt_df = read_excel_stripped(wb_path, sheet_name=timetable_sheet, logger=logger)
    cr_df = read_excel_stripped(wb_path, sheet_name=course_map_sheet, logger=logger)
    rn_df = read_excel_stripped(wb_path, sheet_name=roll_name_sheet, logger=logger)
    room_df = read_excel_stripped(wb_path, sheet_name=room_sheet, logger=logger)

    # --------------------------
    # Build registrations table
    # --------------------------
    # Normalise column names for timetable
    tt_df.columns = [c.strip().lower() for c in tt_df.columns]
    # Expect at least: date, morning, evening
    if "date" not in tt_df.columns:
        raise ValueError("Timetable sheet must contain a 'Date' column.")
    if "morning" not in tt_df.columns or "evening" not in tt_df.columns:
        raise ValueError("Timetable sheet must contain 'Morning' and 'Evening' columns.")

    # Normalise course-roll mapping
    cr_df.columns = [c.strip().lower() for c in cr_df.columns]
    # Expect: rollno, course_code
    if "rollno" not in cr_df.columns and "roll" in cr_df.columns:
        cr_df = cr_df.rename(columns={"roll": "rollno"})
    if "course_code" not in cr_df.columns and "coursecode" in cr_df.columns:
        cr_df = cr_df.rename(columns={"coursecode": "course_code"})
    if "rollno" not in cr_df.columns or "course_code" not in cr_df.columns:
        raise ValueError(
            "Course-roll mapping sheet must contain 'rollno' and 'course_code' columns."
        )

    cr_df["rollno"] = cr_df["rollno"].astype(str).str.strip()
    cr_df["course_code"] = cr_df["course_code"].astype(str).str.strip()

    # Build reg_df = rows of (date, slot, coursecode, rollno)
    registrations_rows = []

    for _, row in tt_df.iterrows():
        # Date as YYYY-MM-DD string
        date_val = row["date"]
        if isinstance(date_val, pd.Timestamp):
            date_str = date_val.date().isoformat()
        else:
            date_str = str(date_val).strip()

        for slot_col, slot_label in [("morning", "morning"), ("evening", "evening")]:
            cell = row.get(slot_col, None)
            if pd.isna(cell):
                continue
            cell_str = str(cell).strip()
            if not cell_str:
                continue
            if cell_str.upper().startswith("NO EXAM"):
                # No exam in this slot
                continue

            # Split "CS249; CH426; ..." into individual course codes
            course_codes = [c.strip() for c in cell_str.split(";") if c.strip()]
            for course in course_codes:
                # Find all rolls registered for this course
                rolls = cr_df.loc[cr_df["course_code"] == course, "rollno"]
                rolls = [str(r).strip() for r in rolls]
                for r in rolls:
                    registrations_rows.append({
                        "date": date_str,
                        "slot": slot_label,
                        "coursecode": course,
                        "rollno": r,
                    })

    if not registrations_rows:
        raise ValueError("No registrations could be built from timetable and course-roll mapping.")

    reg_df = pd.DataFrame(registrations_rows)
    logger.info("Constructed registrations table with %d rows.", len(reg_df))

    # --------------------------
    # Build classrooms table
    # --------------------------
    room_df.columns = [c.strip().lower() for c in room_df.columns]
    # Map known columns
    rename_room = {}
    for col in room_df.columns:
        if col in {"room no.", "roomno", "room_no", "room"}:
            rename_room[col] = "room"
        elif col in {"exam capacity", "capacity", "cap"}:
            rename_room[col] = "capacity"
        elif col in {"block", "building"}:
            rename_room[col] = "building"
    room_df = room_df.rename(columns=rename_room)

    for col in ["building", "room", "capacity"]:
        if col not in room_df.columns:
            raise ValueError(f"Room capacity sheet missing required column: {col}")

    class_df = room_df[["building", "room", "capacity"]].copy()

    def safe_int(x):
        try:
            return int(str(x).strip())
        except Exception:
            return 0

    class_df["building"] = class_df["building"].astype(str).str.strip()
    class_df["room"] = class_df["room"].astype(str).str.strip()
    class_df["capacity"] = class_df["capacity"].apply(safe_int)

    # --------------------------
    # Build roll-name mapping
    # --------------------------
    rn_df.columns = [c.strip().lower() for c in rn_df.columns]
    rename_rollmap = {}
    for col in rn_df.columns:
        if col in {"rollno", "roll", "roll_number"}:
            rename_rollmap[col] = "rollno"
        elif col in {"name", "studentname", "student_name"}:
            rename_rollmap[col] = "name"
    rn_df = rn_df.rename(columns=rename_rollmap)

    if "rollno" not in rn_df.columns or "name" not in rn_df.columns:
        raise ValueError(
            "Roll-name mapping sheet must contain columns for roll and name."
        )

    rn_df["rollno"] = rn_df["rollno"].astype(str).str.strip()
    rn_df["name"] = rn_df["name"].astype(str).str.strip()

    roll_to_name = dict(zip(rn_df["rollno"], rn_df["name"]))

    return reg_df, class_df, roll_to_name


def compute_effective_capacities(class_df: pd.DataFrame, buffer: int,
                                 mode: str, logger: logging.Logger):
    """
    Returns:
        rooms_info: list of dicts with keys:
            building, room, capacity, effective_capacity, per_subject_capacity
    """
    rooms_info = []
    for _, row in class_df.iterrows():
        building = row["building"]
        room = row["room"]
        cap = int(row["capacity"])
        effective = max(cap - buffer, 0)
        if mode == "sparse":
            per_subject = effective // 2
        else:
            per_subject = effective
        logger.info(
            "Room %s-%s: capacity=%d, effective=%d, per_subject=%d (mode=%s)",
            building, room, cap, effective, per_subject, mode
        )
        rooms_info.append({
            "building": building,
            "room": room,
            "capacity": cap,
            "effective_capacity": effective,
            "per_subject_capacity": per_subject,
        })
    return rooms_info


def check_clashes_for_slot(slot_df: pd.DataFrame, logger: logging.Logger):
    """
    For a given (date, slot), check if any roll appears in more than one course.
    Prints clashes and logs them.
    """
    courses = sorted(slot_df["coursecode"].unique())
    course_to_rolls = {}
    for c in courses:
        rs = sorted(set(slot_df[slot_df["coursecode"] == c]["rollno"]))
        course_to_rolls[c] = set(rs)

    clashes_found = False
    for i in range(len(courses)):
        for j in range(i + 1, len(courses)):
            ci, cj = courses[i], courses[j]
            inter = course_to_rolls[ci].intersection(course_to_rolls[cj])
            if inter:
                clashes_found = True
                for roll in sorted(inter):
                    msg = f"CLASH: roll {roll} in both {ci} and {cj}"
                    print(msg)
                    logger.error(msg)
    if not clashes_found:
        logger.info("No clashes detected for this slot.")


def allocate_for_slot(date: str, slot: str, slot_df: pd.DataFrame,
                      rooms_info, roll_to_name, logger: logging.Logger):
    """
    Allocate students to rooms for a single (date, slot).

    Returns:
        allocations: list of dicts with keys:
            date, slot, building, room, coursecode, rollno, name
        seats_left: dict[(building, room)] -> remaining_seats
    """
    logger.info("Allocating for %s %s", date, slot)

    # course -> list(rolls) (deduplicated, order preserved)
    course_to_rolls = defaultdict(list)
    for _, row in slot_df.iterrows():
        c = row["coursecode"]
        r = row["rollno"]
        if r not in course_to_rolls[c]:
            course_to_rolls[c].append(r)

    # Log course sizes
    for c, rolls in course_to_rolls.items():
        logger.info("Course %s has %d students", c, len(rolls))

    # Per-room remaining capacities
    room_caps = {}
    building_rooms = defaultdict(list)  # building -> list of (building, room)
    total_capacity = 0
    for rinfo in rooms_info:
        b = rinfo["building"]
        rn = rinfo["room"]
        cap = int(rinfo["per_subject_capacity"])
        key = (b, rn)
        room_caps[key] = cap
        building_rooms[b].append(key)
        total_capacity += cap

    # Sort rooms within each building to keep them "adjacent"
    for b in building_rooms:
        building_rooms[b] = sorted(building_rooms[b], key=lambda x: x[1])

    # Total students
    total_students = sum(len(v) for v in course_to_rolls.values())
    logger.info("Total students in this slot: %d", total_students)
    logger.info("Total available effective seats: %d", total_capacity)

    if total_students > total_capacity:
        msg = (f"Cannot allocate due to excess students for {date} {slot}. "
               f"Students: {total_students}, Seats: {total_capacity}")
        print(msg)
        logger.error(msg)
        return [], room_caps

    # Sort courses by descending size (largest course first)
    sorted_courses = sorted(course_to_rolls.items(),
                            key=lambda kv: len(kv[1]),
                            reverse=True)

    allocations = []

    for course, rolls in sorted_courses:
        remaining = list(rolls)  # copy
        logger.info("Allocating course %s (%d students)", course, len(remaining))

        # First try single-building allocation
        single_building_found = False
        building_cap = {}
        for b, rkeys in building_rooms.items():
            cap_b = sum(room_caps[k] for k in rkeys)
            building_cap[b] = cap_b
        # Buildings that can hold the entire course
        can_hold = [b for b, cap_b in building_cap.items() if cap_b >= len(remaining)]

        if can_hold:
            # Choose building with minimal sufficient capacity
            chosen_building = sorted(can_hold, key=lambda b: building_cap[b])[0]
            single_building_found = True
            logger.info(
                "Course %s allocated within single building %s (capacity=%d)",
                course, chosen_building, building_cap[chosen_building]
            )

            for key in building_rooms[chosen_building]:
                if not remaining:
                    break
                cap = room_caps[key]
                if cap <= 0:
                    continue
                take = min(cap, len(remaining))
                assigned = remaining[:take]
                remaining = remaining[take:]
                room_caps[key] -= take
                b, rn = key
                for roll in assigned:
                    name = roll_to_name.get(roll, "Unknown Name")
                    allocations.append({
                        "date": date,
                        "slot": slot,
                        "building": b,
                        "room": rn,
                        "coursecode": course,
                        "rollno": roll,
                        "name": name,
                    })
        # If single-building allocation not possible, spread across buildings
        if not single_building_found:
            logger.info(
                "Course %s cannot fit in a single building, spreading across multiple.",
                course
            )
            # Sort all rooms globally by (building, room)
            all_rooms_sorted = sorted(room_caps.keys(), key=lambda x: (x[0], x[1]))
            for key in all_rooms_sorted:
                if not remaining:
                    break
                cap = room_caps[key]
                if cap <= 0:
                    continue
                take = min(cap, len(remaining))
                assigned = remaining[:take]
                remaining = remaining[take:]
                room_caps[key] -= take
                b, rn = key
                for roll in assigned:
                    name = roll_to_name.get(roll, "Unknown Name")
                    allocations.append({
                        "date": date,
                        "slot": slot,
                        "building": b,
                        "room": rn,
                        "coursecode": course,
                        "rollno": roll,
                        "name": name,
                    })

        if remaining:
            # This should not happen if total capacity check passed, but log anyway
            msg = (f"WARNING: After allocation, course {course} still has "
                   f"{len(remaining)} unallocated students for {date} {slot}.")
            print(msg)
            logger.error(msg)

    return allocations, room_caps


def build_overall_and_seats(all_allocations, per_slot_room_caps, rooms_info,
                            logger: logging.Logger, output_dir: str):
    """
    Build the overall seating arrangement dataframe and seats-left dataframe,
    and write them to Excel files. Also returns the overall_df (per-student)
    and overall_agg_df (per room/course) for further use.
    """
    if not all_allocations:
        logger.info("No allocations to write.")
        return None, None

    os.makedirs(output_dir, exist_ok=True)

    overall_df = pd.DataFrame(all_allocations)

    # Build "overall seating arrangement" aggregated per room/course
    def join_semi(col_values):
        # Ensure semicolon-separated, no extra spaces
        return ";".join(str(v) for v in col_values)

    grouped = overall_df.groupby(
        ["date", "slot", "building", "room", "coursecode"],
        sort=True
    )

    rows = []
    for (date, slot, building, room, coursecode), g in grouped:
        rolls = join_semi(g["rollno"].tolist())
        names = join_semi(g["name"].tolist())
        rows.append({
            "Date": date,
            "Slot": slot,
            "Building": building,
            "Room": room,
            "CourseCode": coursecode,
            "RollNumbers": rolls,
            "Names": names,
        })

    overall_agg_df = pd.DataFrame(rows)
    overall_path = os.path.join(output_dir, "op_overall_seating_arrangement.xlsx")
    logger.info("Writing overall seating arrangement to %s", overall_path)
    overall_agg_df.to_excel(overall_path, index=False)

    # Seats left: we know per_subject_capacity from rooms_info
    room_base_cap = {}
    for rinfo in rooms_info:
        key = (rinfo["building"], rinfo["room"])
        room_base_cap[key] = rinfo["per_subject_capacity"]

    seats_rows = []
    for key_slot, room_caps in per_slot_room_caps.items():
        date, slot = key_slot
        for (b, rn), remaining in room_caps.items():
            base = room_base_cap.get((b, rn), 0)
            used = base - remaining
            seats_rows.append({
                "Date": date,
                "Slot": slot,
                "Building": b,
                "Room": rn,
                "EffectiveCapacityPerSubject": base,
                "UsedSeats": used,
                "SeatsLeft": remaining,
            })

    seats_df = pd.DataFrame(seats_rows)
    seats_out_path = os.path.join(output_dir, "op_seats_left.xlsx")
    logger.info("Writing seats left to %s", seats_out_path)
    seats_df.to_excel(seats_out_path, index=False)

    # Also write per-slot files in date/slot folders
    for (date, slot), slot_df in overall_agg_df.groupby(["Date", "Slot"]):
        slot_dir = os.path.join(output_dir, date, slot)
        os.makedirs(slot_dir, exist_ok=True)
        slot_path = os.path.join(slot_dir, "seating_arrangement.xlsx")
        logger.info("Writing per-slot seating file to %s", slot_path)
        slot_df.to_excel(slot_path, index=False)

    return overall_df, overall_agg_df


# ---------------------------------------------------------------------------
# Attendance PDF generation
# ---------------------------------------------------------------------------

def draw_attendance_page_header(c, width, height, date_str, slot, room, course):
    """
    Draws the header for each attendance page.
    """
    # Numeric margins in points (1 mm = 2.835... points)
    top_margin = 20 * mm  # ~20 mm
    left_margin = 15 * mm

    c.setFont("Helvetica-Bold", 14)
    c.drawString(left_margin, height - top_margin, "EXAMINATION ATTENDANCE SHEET")

    c.setFont("Helvetica", 11)
    line_y = height - top_margin - 18
    c.drawString(left_margin, line_y, f"Date: {date_str}")
    c.drawString(left_margin + 200, line_y, f"Session: {slot.title()}")

    line_y -= 16
    c.drawString(left_margin, line_y, f"Course Code: {course}")
    c.drawString(left_margin + 200, line_y, f"Room: {room}")

    # Return y position to start the table header from
    return line_y - 24  # leave some gap


def draw_attendance_table_header(c, width, start_y):
    """
    Draws the table header row (column titles) and returns the y position
    for the first data row.
    """
    left_margin = 15 * mm
    right_margin = 15 * mm
    usable_width = width - left_margin - right_margin

    # Define simple column widths
    col_sn = 25
    col_photo = 40
    col_roll = 70
    col_name = usable_width - (col_sn + col_photo + col_roll + 120)
    if col_name < 80:
        col_name = 80  # minimum
    col_sign = 80  # just conceptual, last column ends at right margin

    col_x = [
        left_margin,
        left_margin + col_sn,
        left_margin + col_sn + col_photo,
        left_margin + col_sn + col_photo + col_roll,
        left_margin + col_sn + col_photo + col_roll + col_name,
    ]

    # Header labels
    headers = ["S.No", "Photo", "Roll No.", "Name", "Signature"]

    c.setFont("Helvetica-Bold", 10)
    header_y = start_y
    row_height = 18

    # Draw header texts
    c.drawString(col_x[0] + 2, header_y, headers[0])
    c.drawString(col_x[1] + 2, header_y, headers[1])
    c.drawString(col_x[2] + 2, header_y, headers[2])
    c.drawString(col_x[3] + 2, header_y, headers[3])
    c.drawString(col_x[4] + 2, header_y, headers[4])

    # Draw horizontal line under header
    c.line(left_margin, header_y - 2, width - right_margin, header_y - 2)

    # Return positions needed for data rows
    return {
        "row_start_y": header_y - row_height,
        "row_height": row_height,
        "col_x": col_x,
        "left_margin": left_margin,
        "right_margin": right_margin,
    }


def generate_attendance_pdf_for_group(date_str, slot, room, course,
                                      students_df, out_path,
                                      photos_dir, logger):
    try:
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        c = canvas.Canvas(out_path, pagesize=A4)
        width, height = A4

        margin_x = 15 * mm
        margin_y = 15 * mm
        y = height - margin_y

        # ================= HEADER =================
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(width / 2, y, "IITP Attendance System")
        y -= 20

        c.setFont("Helvetica", 11)
        c.drawString(
            margin_x,
            y,
            f"Date: {date_str} | Shift: {slot.title()} | Room No: {room} | Student count: {len(students_df)}"
        )
        y -= 15

        c.drawString(
            margin_x,
            y,
            f"Subject: {course} | Stud Present:      | Stud Absent:"
        )
        y -= 25

        # ================= CARD GRID =================
        card_w = (width - 2 * margin_x) / 3
        card_h = 45 * mm

        x_positions = [
            margin_x,
            margin_x + card_w,
            margin_x + 2 * card_w
        ]

        col = 0

        for _, stu in students_df.iterrows():
            if col == 0 and y - card_h < margin_y:
                c.showPage()
                y = height - margin_y

            x = x_positions[col]
            c.rect(x, y - card_h, card_w - 5, card_h)

            roll = str(stu["rollno"])
            name = stu["name"] if stu["name"] else "Unknown Name"

            card_top = y
            card_left = x
            card_right = x + card_w - 10

            # Photo
            photo_size = 22 * mm
            photo_x = card_left + 8
            photo_y = card_top - photo_size - 8

            photo_path = (
                os.path.join(photos_dir, f"{roll}.jpg")
                if photos_dir else None
            )

            if photo_path and os.path.exists(photo_path):
                try:
                    c.drawImage(
                        ImageReader(photo_path),
                        photo_x, photo_y,
                        photo_size, photo_size,
                        preserveAspectRatio=True,
                        mask="auto"
                    )
                except Exception:
                    c.rect(photo_x, photo_y, photo_size, photo_size)
            else:
                c.rect(photo_x, photo_y, photo_size, photo_size)
                c.setFont("Helvetica", 7)
                c.drawCentredString(
                    photo_x + photo_size / 2,
                    photo_y + photo_size / 2,
                    "No Image"
                )

            # Name wrapping (no font compromise)
            text_x = photo_x + photo_size + 10
            text_y = card_top - 18
            text_width = card_right - text_x

            c.setFont("Helvetica-Bold", 10)
            words = name.split()
            lines, line = [], ""

            for w in words:
                if c.stringWidth(line + " " + w, "Helvetica-Bold", 10) <= text_width:
                    line = (line + " " + w).strip()
                else:
                    lines.append(line)
                    line = w
            if line:
                lines.append(line)

            for ln in lines[:2]:
                c.drawString(text_x, text_y, ln)
                text_y -= 12

            c.setFont("Helvetica", 9)
            c.drawString(text_x, text_y, f"Roll: {roll}")

            sign_y = card_top - card_h + 12
            c.drawString(text_x, sign_y + 8, "Sign:")
            c.line(text_x + 35, sign_y + 8, card_right, sign_y + 8)

            # Grid movement
            col += 1
            if col == 3:
                col = 0
                y -= card_h + 6

        # ================= CRITICAL FIX =================
        # If cards ended mid-row, force move to next row
        if col != 0:
            y -= card_h + 6
            col = 0

        # Strong visible separation (relative margin)
        y -= 25 * mm

        if y < margin_y + 90:
            c.showPage()
            y = height - margin_y

        # ================= INVIGILATOR SECTION =================
        c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(width / 2, y, "Invigilator Name & Signature")
        y -= 18

        c.setFont("Helvetica-Bold", 10)
        c.drawString(margin_x, y, "Sl No.")
        c.drawString(margin_x + 60, y, "Name")
        c.drawString(width - margin_x - 120, y, "Signature")
        y -= 10

        c.setFont("Helvetica", 10)
        for _ in range(6):
            y -= 16
            c.line(margin_x, y, width - margin_x, y)

        c.save()

    except Exception:
        logger.error(
            "Error generating PDF for %s %s %s %s",
            date_str, slot, room, course,
            exc_info=True
        )







def generate_all_attendance_pdfs(overall_df: pd.DataFrame,
                                 logger: logging.Logger,
                                 attendance_dir: str,
                                 photos_dir: str):
    """
    overall_df: per-student dataframe with columns
        date, slot, building, room, coursecode, rollno, name
    """
    if overall_df is None or overall_df.empty:
        logger.info("No allocations available for PDF generation.")
        return

    for (date, slot, room, course), group in overall_df.groupby(
        ["date", "slot", "room", "coursecode"]
    ):
        # Build filename: YYYY_MM_DD_<SESSION>_R<ROOM>_<SUBCODE>.pdf
        date_clean = date.replace("-", "_")
        session_str = slot.title()  # "morning" -> "Morning"
        room_str = str(room)
        course_str = str(course)

        filename = f"{date_clean}_{session_str}_R{room_str}_{course_str}.pdf"
        out_path = os.path.join(attendance_dir, date, slot, filename)

        # Prepare student dataframe with required columns
        students_df = group[["rollno", "name"]].copy()

        generate_attendance_pdf_for_group(
            date_str=date,
            slot=slot,
            room=room_str,
            course=course_str,
            students_df=students_df,
            out_path=out_path,
            photos_dir=photos_dir,
            logger=logger,
        )


# ---------------------------------------------------------------------------
# Argument parsing and main
# ---------------------------------------------------------------------------

def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Design exam seating arrangement + attendance PDFs (single workbook)."
    )
    parser.add_argument(
        "--input",
        default="input_data_tt.xlsx",
        help="Path to the Excel workbook containing all input sheets "
             "(default: input_data_tt.xlsx)",
    )
    parser.add_argument(
        "--buffer",
        type=int,
        default=0,
        help="Buffer to subtract from classroom capacities (default: 0)",
    )
    parser.add_argument(
        "--mode",
        choices=["sparse", "dense"],
        default="dense",
        help="Seating density mode: sparse (50%%) or dense (10%%) (default: dense)",
    )
    parser.add_argument(
        "--output-dir",
        default="output",
        help="Output directory for Excel files (default: output)",
    )
    parser.add_argument(
        "--attendance-dir",
        default="attendance_pdfs",
        help="Output directory for attendance PDFs (default: attendance_pdfs)",
    )
    parser.add_argument(
        "--photos-dir",
        default="photos",
        help="Directory where photos are stored as ROLL.jpg (default: photos). "
             "If it does not exist, photos are simply skipped.",
    )
    return parser.parse_args(argv)


def main(argv=None):
    logger = setup_logging()
    try:
        args = parse_args(argv)
        logger.info("Starting seating arrangement generation (single workbook).")
        logger.info("Arguments: %s", args)

        reg_df, class_df, roll_to_name = load_inputs_from_workbook(
            args.input, logger
        )

        rooms_info = compute_effective_capacities(
            class_df, buffer=args.buffer, mode=args.mode, logger=logger
        )

        all_allocations = []
        per_slot_room_caps = {}

        # Process each (date, slot)
        for (date, slot), slot_df in reg_df.groupby(["date", "slot"]):
            logger.info("Processing date=%s, slot=%s", date, slot)

            # Check clashes
            check_clashes_for_slot(slot_df, logger)

            allocations, room_caps = allocate_for_slot(
                date, slot, slot_df, rooms_info, roll_to_name, logger
            )
            all_allocations.extend(allocations)
            per_slot_room_caps[(date, slot)] = room_caps

        # Build Excel outputs and get per-student + aggregated dataframes
        overall_df, overall_agg_df = build_overall_and_seats(
            all_allocations,
            per_slot_room_caps,
            rooms_info,
            logger=logger,
            output_dir=args.output_dir,
        )

        # Generate attendance PDFs (if reportlab available)
        if A4 is None or canvas is None:
            logger.error(
                "reportlab is not installed, skipping attendance PDF generation. "
                "Install it with 'pip install reportlab'."
            )
        else:
            photos_dir = args.photos_dir
            if not os.path.isdir(photos_dir):
                logger.warning(
                    "Photos directory '%s' does not exist. "
                    "Attendance PDFs will be generated without photos.",
                    photos_dir,
                )
                photos_dir = None

            generate_all_attendance_pdfs(
                overall_df=overall_df,
                logger=logger,
                attendance_dir=args.attendance_dir,
                photos_dir=photos_dir,
            )

        logger.info("Seating arrangement generation completed successfully.")

    except Exception as e:
        print("An unexpected error occurred. Check errors.txt for details.")
        logger = logging.getLogger("seating")
        logger.error("Unexpected error: %s", e)
        logger.error(traceback.format_exc())


if __name__ == "__main__":
    main()
