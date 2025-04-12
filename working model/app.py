import requests
import fitz  # PyMuPDF
import re
import os
import pandas as pd  # For reading Excel files
from bs4 import BeautifulSoup

from flask import Flask, render_template, request
import tempfile

app = Flask(__name__)

def extract_file_id(doc_url):
    """Extracts the Google Docs file ID from the provided URL."""
    if not isinstance(doc_url, str) or pd.isna(doc_url):  # Check if URL is NaN or not a string
        return None

    match = re.search(r"https://docs\.google\.com/document/d/([a-zA-Z0-9-_]+)", doc_url)
    return match.group(1) if match else None

def sanitize_filename(filename):
    """Converts a URL into a safe filename by replacing invalid characters."""
    if not isinstance(filename, str) or pd.isna(filename):  # Check if filename is NaN
        return "Untitled"

    filename = filename.strip().replace("https://", "").replace("http://", "")  # Remove URL scheme
    filename = re.sub(r'[<>:"/\\|?*]', "_", filename)  # Replace invalid characters with "_"
    return filename[:200]  # Ensure filename is not too long

def download_pdf(file_id, file_name, max_retries=3):
    """Downloads a Google Docs document as a PDF file."""
    pdf_url = f"https://docs.google.com/document/d/{file_id}/export?format=pdf"
    file_name = sanitize_filename(file_name)  # Make filename safe
    pdf_filename = f"{file_name}.pdf"

    session = requests.Session()
    retry_count = 0

    while retry_count < max_retries:
        try:
            with session.get(pdf_url, stream=True, timeout=30) as response:
                if response.status_code == 200:
                    with open(pdf_filename, "wb") as f:
                        for chunk in response.iter_content(1024):
                            f.write(chunk)
                    return pdf_filename
                else:
                    print(f"❌ Failed to download {file_name}, Status Code: {response.status_code}")
                    return None
        except requests.exceptions.RequestException:
            retry_count += 1

    return None

def extract_editorial_links_from_pdf(pdf_path):
    """Extracts only editorial-use image URLs from a PDF."""
    doc = fitz.open(pdf_path)
    editorial_links = []
    
    url_pattern = re.compile(r"https?://[^\s]+")  # Detect URLs in text

    for page in doc:
        # Extract clickable links
        for link in page.get_links():
            if "uri" in link:
                url = link["uri"]
                editorial_status = check_editorial_use(url)
                if editorial_status == "[✔] Editorial Use Only":
                    editorial_links.append(url)

        # Extract plain text links that contain "123rf"
        text = page.get_text("text")
        found_links = url_pattern.findall(text)

        for url in found_links:
            if "123rf" in url:  # Only allow 123rf links
                editorial_status = check_editorial_use(url)
                if editorial_status == "[✔] Editorial Use Only":
                    editorial_links.append(url)

    return editorial_links

def check_editorial_use(image_url):
    """Checks whether the given URL is for editorial use only."""
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(image_url, headers=headers, timeout=15)
        if response.status_code != 200:
            return "[❌] Failed to retrieve"

        soup = BeautifulSoup(response.text, "html.parser")
        return "[✔] Editorial Use Only" if "editorial use only" in soup.text.lower() else "[✘] Not Editorial"
    except requests.RequestException:
        return "[❌] Error fetching"

def process_google_docs_from_excel(excel_file):
    """Reads an Excel file, processes Google Docs, extracts editorial links, and saves results to output.txt in real-time."""
    df = pd.read_excel(excel_file)  # Read Excel file

    if 'Google Docs URL' not in df.columns or 'Website URL' not in df.columns:
        print("❌ Excel file must have columns: 'Google Docs URL' and 'Website URL'")
        return

    # with open("done.txt", "w", encoding="utf-8") as output_file:
    #     output_file.write("Final Extracted Editorial Images:\n")
    #     output_file.write("-" * 50 + "\n")
    results = []
    for index, row in df.iterrows():
        doc_url = row['Google Docs URL']
        website_url = row['Website URL']

        if pd.isna(doc_url) or pd.isna(website_url):  # Skip empty rows
            print(f"⚠️ Skipping row {index + 1} due to missing data.")
            results.append({"title": formatted_title,"links": editorial_links})
            continue

        file_id = extract_file_id(str(doc_url))  # Convert to string

        if file_id:
            safe_file_name = sanitize_filename(website_url)  # Convert Website URL to a safe filename
            pdf_file = download_pdf(file_id, safe_file_name)  # Use sanitized Website URL as filename

            if pdf_file:
                editorial_links = extract_editorial_links_from_pdf(pdf_file)
                os.remove(pdf_file)  # Delete the PDF after processing

                if editorial_links:  # Only save results if there are editorial links
                    formatted_title = safe_file_name.replace("_", "/")
                    results.append({"title": formatted_title, "links": editorial_links})
                    formatted_title = safe_file_name.replace("_", "/")  # Convert _ to / for display
                    
                    # Print to console immediately
                    print(f"Title: {formatted_title}")
                    results.append({"title": formatted_title,"links": editorial_links})

                    for url in editorial_links:
                        print(f"Editorial Image URL: {url}")
                        results.append({"title": formatted_title,"links": editorial_links})

                    print("-" * 50)
                    # output_file.write("-" * 50 + "\n")

    print("\n✅ All results have been saved to output.txt in real-time!")
    return results


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        excel_file = request.files.get("excel_file")
        if not excel_file:
            return "❌ Please upload a valid Excel file (.xlsx)."

        file_path = os.path.join(tempfile.gettempdir(), excel_file.filename)
        excel_file.save(file_path)

        results = process_google_docs_from_excel(file_path)
        return render_template("results.html", results=results)

    return render_template("index.html")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

# Run the program with the Excel file
# process_google_docs_from_excel(r"C:\Users\Intel\Desktop\ANUSHKA\main\Done .xlsx")
