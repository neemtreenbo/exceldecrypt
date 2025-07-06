FROM python:3.11-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy your app
COPY . .

EXPOSE 8000

# Start using Gunicorn
CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:8000"]