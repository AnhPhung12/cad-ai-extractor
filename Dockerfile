FROM python:3.10-slim

# Install system dependencies needed for ODA File Converter
RUN apt-get update && apt-get install -y wget gdebi-core libgl1 libglib2.0-0

# Copy the ODA File Converter .deb package and install it
COPY ODAFileConverter_QT6_lnxX64_8.3dll_27.1.deb /tmp/oda.deb
RUN gdebi -n /tmp/oda.deb

# Set the working directory
WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy all files
COPY . .

# Expose port and run the app
EXPOSE 8000
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
