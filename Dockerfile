FROM python:3.12-slim

# Google Chrome (needed for Selenium - webdriver-manager only fetches a
# matching chromedriver, not the browser itself)
RUN apt-get update && apt-get install -y --no-install-recommends \
        wget ca-certificates \
    && wget -q -O /tmp/chrome.deb https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb \
    && apt-get install -y --no-install-recommends /tmp/chrome.deb \
    && rm /tmp/chrome.deb \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# job_state is a single in-memory dict shared across requests, so this must
# run as exactly one worker process - gthread gives it multiple threads so
# status polling stays responsive while a scrape's background thread runs.
CMD gunicorn -w 1 -k gthread --threads 8 --timeout 120 -b 0.0.0.0:${PORT:-5000} app:app
