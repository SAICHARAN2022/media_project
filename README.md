# YouTube Video Data and Comments Fetcher

This Python script fetches metadata and comments (including replies) for a specified YouTube video using the YouTube Data API v3 and `yt-dlp`. It saves the fetched data to an Excel file with separate sheets for metadata and comments.

---

## Features

- Fetch YouTube video metadata (e.g., title, description, views, likes, comments count, etc.).
- Fetch latest comments and their replies with pagination.
- Save the fetched data into an Excel file.
- Securely handle sensitive API keys using environment variables.

---

## Prerequisites

1. **Python 3.10 or higher** (ensure it's installed on your system).
2. Required Python libraries:
   - `pandas`
   - `google-api-python-client`
   - `yt-dlp`
   - `openpyxl`
3. A valid YouTube Data API key.

---

## Installation

### 1. Clone the Repository
```bash
git clone https://github.com/SAICHARAN2022/media_project.git
cd media_project
