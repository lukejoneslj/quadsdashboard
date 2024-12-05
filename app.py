import os
import requests
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# Constants
FOLDER_URL = "https://drive.google.com/drive/folders/1TrnQQGpon33QP0WQ-bFS7WoMLQAuAPyE?usp=sharing"
PROCESSED_FILES_LOG = "processed_files.txt"
DOWNLOAD_DIR = "downloads"
REPORTS_DIR = "reports"

# Ensure directories exist
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

# Step 1: Scrape file links from Google Drive folder
def get_file_links(folder_url):
    response = requests.get(folder_url)
    if response.status_code != 200:
        raise Exception(f"Failed to access folder. Status code: {response.status_code}")
    
    # Parse file links (assuming standard public folder structure)
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(response.text, 'html.parser')
    links = [
        f"https://drive.google.com/uc?id={link['href'].split('/d/')[1].split('/')[0]}&export=download"
        for link in soup.find_all('a', href=True)
        if "file/d/" in link['href']
    ]
    return links

# Step 2: Download and process new files
def download_file(file_url, save_path):
    response = requests.get(file_url)
    if response.status_code == 200:
        with open(save_path, "wb") as file:
            file.write(response.content)
    else:
        print(f"Failed to download file from {file_url}")

def get_processed_files():
    if not os.path.exists(PROCESSED_FILES_LOG):
        return set()
    with open(PROCESSED_FILES_LOG, "r") as f:
        return set(f.read().splitlines())

def update_processed_files(file_name):
    with open(PROCESSED_FILES_LOG, "a") as f:
        f.write(f"{file_name}\n")

# Step 3: Run analysis
def analyze_file(file_path):
    # Load the file
    if file_path.endswith(".xlsx"):
        data = pd.read_excel(file_path, sheet_name="Overall Quad")
    elif file_path.endswith(".csv"):
        data = pd.read_csv(file_path)
    else:
        raise ValueError("Unsupported file type.")
    
    # Example analysis (adapt as needed)
    counts_values = {
        "80 Customer": f"{int(round(data.iloc[5, 1]))}",
        "80 Part": f"{int(round(data.iloc[4, 1]))}",
    }
    counts_df = pd.DataFrame(list(counts_values.items()), columns=["Type", "Count"])
    return counts_df

# Step 4: Generate PDF report
def create_pdf_report(dataframe, output_path):
    pdf = SimpleDocTemplate(output_path, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()

    # Title
    elements.append(Paragraph("Executive Summary Dashboard", styles['Title']))
    elements.append(Spacer(1, 12))

    # Add DataFrame as a table
    table_data = [dataframe.columns.tolist()] + dataframe.values.tolist()
    table = Table(table_data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))
    elements.append(table)

    # Build PDF
    pdf.build(elements)
    print(f"Report saved to {output_path}")

# Main function
def main():
    links = get_file_links(FOLDER_URL)
    processed_files = get_processed_files()

    for idx, link in enumerate(links):
        file_name = f"file_{idx + 1}.xlsx"
        if file_name not in processed_files:
            file_path = os.path.join(DOWNLOAD_DIR, file_name)
            download_file(link, file_path)
            
            # Run analysis
            report_data = analyze_file(file_path)
            
            # Generate report
            report_path = os.path.join(REPORTS_DIR, f"report_{idx + 1}.pdf")
            create_pdf_report(report_data, report_path)
            
            # Mark file as processed
            update_processed_files(file_name)

if __name__ == "__main__":
    main()
