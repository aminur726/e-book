# FROM python:3.12-slim

# # Install WeasyPrint system deps, Node.js, Bengali fonts
# RUN apt-get update && apt-get install -y \
#     libpango1.0-dev libpangocairo-1.0-0 libcairo2-dev \
#     libgdk-pixbuf2.0-dev libffi-dev shared-mime-info \
#     poppler-utils pandoc fonts-noto fonts-noto-cjk \
#     curl && \
#     curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
#     apt-get install -y nodejs && \
#     npm install -g docx && \
#     rm -rf /var/lib/apt/lists/*

# WORKDIR /app
# COPY requirements.txt .
# RUN pip install --no-cache-dir -r requirements.txt
# COPY . .

# EXPOSE 8501
# CMD ["streamlit", "run", "app.py", \
#      "--server.port=8501", \
#      "--server.headless=true", \
#      "--server.maxUploadSize=200", \
#      "--browser.gatherUsageStats=false"]

# FROM python:3.12-slim

# # Install WeasyPrint system deps, Node.js, Bengali fonts
# RUN apt-get update && apt-get install -y --no-install-recommends \
#     libpango1.0-dev libpangocairo-1.0-0 libcairo2-dev \
#     libgdk-pixbuf-xlib-2.0-dev libffi-dev shared-mime-info \
#     poppler-utils pandoc fonts-noto fonts-noto-cjk \
#     curl && \
#     curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
#     apt-get install -y nodejs && \
#     npm install -g docx && \
#     rm -rf /var/lib/apt/lists/*

# WORKDIR /app
# COPY requirements.txt .
# RUN pip install --no-cache-dir -r requirements.txt
# COPY . .

# EXPOSE 80
# CMD ["streamlit", "run", "app.py", \
#      "--server.port=80", \
#      "--server.headless=true", \
#      "--server.maxUploadSize=200", \
#      "--browser.gatherUsageStats=false"]

FROM python:3.12-slim

# Install WeasyPrint system deps, Node.js, Bengali fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpango1.0-dev libpangocairo-1.0-0 libcairo2-dev \
    libgdk-pixbuf-xlib-2.0-dev libffi-dev shared-mime-info \
    poppler-utils pandoc fonts-noto fonts-noto-cjk \
    curl && \
    curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
    apt-get install -y nodejs && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python deps
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy project
COPY . .

# Install Node dependencies (including docx) locally
RUN npm install docx

EXPOSE 80

CMD ["streamlit", "run", "app.py", \
     "--server.port=80", \
     "--server.headless=true", \
     "--server.maxUploadSize=200", \
     "--browser.gatherUsageStats=false"]

