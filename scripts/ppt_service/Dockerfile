FROM python:3.12-slim

# Install Node.js 20 + system libs required by Playwright Chromium + LibreOffice for PDF conversion
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      curl ca-certificates fonts-liberation \
      libnss3 libnspr4 libatk1.0-0 libatk-bridge2.0-0 libcups2 \
      libdrm2 libxkbcommon0 libxcomposite1 libxdamage1 libxfixes3 \
      libxrandr2 libgbm1 libasound2 libpango-1.0-0 libpangocairo-1.0-0 \
      libcairo2 libatspi2.0-0 libx11-6 libxcb1 libxext6 \
      libreoffice-nogui && \
    curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
    apt-get install -y --no-install-recommends nodejs && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install Chromium for Playwright (layout measurement + PDF export)
RUN python -m playwright install chromium

# Install Node dependencies (pptxgenjs)
COPY package.json .
RUN npm install --production

# Copy application code
COPY . .

# Copy master template
COPY master_template.pptx /app/master_template.pptx

# Output directory
RUN mkdir -p /tmp/ppt_output

EXPOSE 8501

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8501"]
