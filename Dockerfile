# Base image with Python
FROM python:3.10-slim

# Install system dependencies (for fonts and docx compatibility)
RUN apt-get update && apt-get install -y \
    fonts-dejavu-core \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy all files
COPY . .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Expose Streamlit port
EXPOSE 8501

# Default command
CMD ["streamlit", "run", "app.py", "--server.port=8080", "--server.address=0.0.0.0"]
