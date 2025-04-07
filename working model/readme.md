# 🧠 Editorial Link Extractor (Google Docs ➝ PDF ➝ Editorial Images)

A Flask web app that extracts **editorial image links** from Google Docs URLs listed in an uploaded Excel file.  
It checks each Google Doc, downloads it as a PDF, and pulls any "editorial use only" image URLs.

## 🚀 Features

- Upload `.xlsx` files with Google Doc URLs
- Downloads docs as PDF automatically
- Extracts editorial-use-only image links
- Beautiful web interface (Bootstrap-based)
- Outputs results in real-time, visually

## 🛠 Tech Stack

- Python 🐍
- Flask 🔥
- PyMuPDF (`fitz`) 📄
- BeautifulSoup 🥣
- Pandas 📊
- Bootstrap 5 💅

## 📁 Expected Excel Format

| Google Docs URL                       | Website URL       |
|--------------------------------------|-------------------|
| `https://docs.google.com/document...`| `example.com/post`|

Both columns are required!

## ⚙️ Setup & Run Locally

```bash
git clone https://github.com/yourusername/editorial-link-extractor.git
cd editorial-link-extractor

# Create a virtual environment (optional but recommended)
python -m venv venv
venv\\Scripts\\activate   # On Windows
# OR
source venv/bin/activate  # On Mac/Linux

# Install dependencies
pip install -r requirements.txt

# Run the app
python app.py
