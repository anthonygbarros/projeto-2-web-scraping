# 🕷️ Python Web Scraping & Excel Automation

A professional Python automation script designed to scrape web data using BeautifulSoup and automatically generate styled, formatted Excel reports.

## 🛠️ Tech Stack
* **Python 3.x**
* **Requests** (HTTP Requests)
* **BeautifulSoup4** (HTML Parsing)
* **OpenPyXL** (Excel Spreadsheet Manipulation & Styling)

## 🚀 Key Features
* Connects to target endpoints and parses HTML structure dynamically.
* Extracts structured text data (quotes, authors).
* Export process automatically builds a corporate `.xlsx` spreadsheet.
* Features styled headers (custom background & bold text) and dynamic column width auto-fitting.

## 💻 How to Run Locally

1. Clone this repository: `git clone https://github.com/anthonygbarros/projeto-2-web-scraping.git`

2. Create and activate a virtual environment:
`python -m venv .venv`
`.venv\Scripts\activate`

3. Install required packages:
`pip install requests beautifulsoup4 openpyxl`

4. Run the main script:
`python web-scraper-python.py`