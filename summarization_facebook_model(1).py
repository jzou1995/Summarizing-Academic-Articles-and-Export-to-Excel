import os
import glob
import time
import sys
from transformers import pipeline
import tqdm
from datetime import datetime
import openpyxl

def progress_bar(current, total, message):
    bar_length = 50
    progress = current / total
    arrow = "=" * int(bar_length * progress) + ">"
    spaces = " " * (bar_length - len(arrow))
    sys.stdout.write(f"\r{message}: [{arrow + spaces}] {int(progress * 100)}%")
    sys.stdout.flush()

summarizer = pipeline("summarization", model="facebook/bart-large-cnn") # Changed to use "facebook/bart-large-cnn" model
text_files = glob.glob("*.txt")

for text_file in text_files:
    author, title, year, journal = text_file[:-4].split("-")

    # Check if the summary file already exists
    summary_file = f"summary_of_{author}-{title}-{year}-{journal}.xlsx"
    if os.path.exists(summary_file):
        print(f"Summary file for {text_file} already exists. Skipping summarization.")
        continue
    
    # Create a new Excel file for each text file
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["ID", "Author", "Title", "Year", "Journal", "Summary ID", "Summary", "Date"])

    with open(text_file, "r", encoding="utf-8") as f:
        content = f.read()

    chunk_size = 1024
    chunks = [content[i:i+chunk_size] for i in range(0, len(content), chunk_size)]

    summaries = []
    for i, chunk in enumerate(chunks):
        for j in tqdm.tqdm(range(100)):
            time.sleep(0.02)
            progress_bar(j + 1, 100, f"Summarizing chunk {i+1}/{len(chunks)} of {text_file}")
        summary = summarizer(chunk)
        summaries.append(summary[0]["summary_text"])

    current_date = datetime.today().strftime('%Y-%m-%d')

    file_id = "1"

    for i, summary in enumerate(summaries):
        summary_id = f"{file_id}-{i + 1}"
        new_row = [file_id, author, title, year, journal, summary_id, summary, current_date]
        sheet.append(new_row)

    print(f"Summary of {text_file} added to the Excel file.")

    # Adjust the width of the "Summary" column
    sheet.column_dimensions["G"].width = 200

    # Freeze the first row
    sheet.freeze_panes = "A2"

    # Save the workbook with a unique name based on the text file
    excel_file = summary_file
    workbook.save(excel_file)
