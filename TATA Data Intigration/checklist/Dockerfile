# Use an official Python runtime as a parent image
FROM python:3.11-slim

# Set the working directory in the container
WORKDIR /app

# Copy the current directory contents into the container at /app
COPY . /app

# Install system dependencies for some Python packages
RUN apt-get update && apt-get install -y \
    gcc \
    libffi-dev \
    libssl-dev \
    && rm -rf /var/lib/apt/lists/*

# Upgrade pip
RUN pip install --upgrade pip

# Check if requirements.txt exists and install dependencies
RUN ls -la /app && cat /app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Expose the necessary port
EXPOSE 8080

# Command to run the application
CMD ["python", "TATA SCS Vishal.py"]
