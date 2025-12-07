import streamlit as st
import os
import shutil
import zipfile
import tempfile
import subprocess
from pathlib import Path


st.set_page_config(page_title="Exam Seating & Attendance Generator", layout="wide")

st.title("📘 Exam Seating Arrangement & Attendance Generator")
st.markdown("Upload the **Excel timetable** and **Photo folder**, then generate outputs.")

# ---------- UPLOADS ----------
xlsx_file = st.file_uploader("Upload input Excel file (.xlsx)", type=["xlsx"])
photo_zip = st.file_uploader("Upload photos folder (.zip) (ROLL.jpg format)", type=["zip"])

buffer = st.number_input("Buffer seats", min_value=0, value=0)
mode = st.selectbox("Seating Mode", ["dense", "sparse"])

generate = st.button("🚀 Generate Seating & PDFs")

# ---------- PROCESS ----------
if generate:
    if not xlsx_file or not photo_zip:
        st.error("Please upload BOTH Excel file and Photos ZIP.")
        st.stop()

    with st.spinner("Processing... please wait"):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # Paths
            input_excel = tmpdir / "input_data_tt.xlsx"
            photos_dir = tmpdir / "photos"
            output_dir = tmpdir / "output"
            attendance_dir = tmpdir / "attendance_pdfs"
            logs_dir = tmpdir / "logs"

            photos_dir.mkdir()
            output_dir.mkdir()

            # Save Excel
            with open(input_excel, "wb") as f:
                f.write(xlsx_file.read())

            # Extract photos zip
            with zipfile.ZipFile(photo_zip, "r") as zip_ref:
                zip_ref.extractall(photos_dir)

            # Copy backend script
            shutil.copy("seating_arrangement.py", tmpdir)

            # Run backend
            cmd = [
                "python",
                "seating_arrangement.py",
                "--input", "input_data_tt.xlsx",
                "--buffer", str(buffer),
                "--mode", mode,
                "--output-dir", "output",
                "--attendance-dir", "attendance_pdfs",
                "--photos-dir", "photos",
            ]

            result = subprocess.run(
                cmd,
                cwd=tmpdir,
                capture_output=True,
                text=True
            )

            # Zip everything
            zip_path = tmpdir / "final_output.zip"
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                for folder in ["output", "attendance_pdfs", "logs"]:
                    folder_path = tmpdir / folder
                    if folder_path.exists():
                        for file in folder_path.rglob("*"):
                            zipf.write(file, file.relative_to(tmpdir))

            st.success("✅ Generation Completed!")

            # Show logs preview
            if logs_dir.exists():
                st.subheader("📄 Execution Log")
                log_file = logs_dir / "execution.log"
                if log_file.exists():
                    st.text(log_file.read_text()[:2000])

            # Download
            with open(zip_path, "rb") as f:
                st.download_button(
                    "⬇️ Download Output ZIP",
                    f,
                    file_name="exam_seating_output.zip",
                    mime="application/zip"
                )
