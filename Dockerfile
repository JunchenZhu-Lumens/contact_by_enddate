# Start with a Python base image
FROM python:3.8

# Set working directory
WORKDIR /app

# Copy all files from the repo
COPY . .

# Install dependencies
RUN pip install -r requirements.txt

# Expose port 8080
EXPOSE 8080

# Run your script
CMD python contact_by_enddate.py
