# Use official Python image
FROM python:3.10-slim

# Set working directory
WORKDIR /app

# Install OS dependencies (if you need zip, pdf libs)
RUN apt-get update && apt-get install -y \
    zip \
    && rm -rf /var/lib/apt/lists/*

# Copy dependency file
COPY requirements.txt .

# Install Python packages
RUN pip install --no-cache-dir -r requirements.txt

# Copy all project files
COPY . .

# Expose Streamlit default port
EXPOSE 8501

# Run Streamlit
ENTRYPOINT [ "sh", "-c", "streamlit run app.py" ]
