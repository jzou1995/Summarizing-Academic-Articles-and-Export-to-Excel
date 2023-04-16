# Summarization_facebook_model
This code uses the facebook/bart-large-cnn model to summarize text files. Compare to lidiya model in earlier version, the facebook model is more adapt to academic writings whereas lidiya model works better for interviews and conversations. For example, you can use the lidiya model to processes subtitles from Youtube interview videos. But Facebook model outperforms lidiya model in completeness and coherence for processing academic articles. 

# Summarizing-Academic-Articles-and-Export-to-Excel
This code uses the lidiya/bart-large-xsum-samsum model to summarize text files. It exports the summary to an Excel file with columns for author, title, year, journal, and IDs. The article is divided into chunks for summarization, creating more rows for longer articles.

#This Python script summarizes text files and stores the results in an Excel file.

## How it works

The script utilizes the Hugging Face Transformers library to summarize the content of text files located in the same folder as the script. The summaries are then stored in an Excel file along with metadata like the author, title, year, and journal.

## Usage

1. Place your text files in the same folder as the script.
2. Run the script by executing `python Summarizing-Academic-Articles-and-Export-to-Excel.py` in your terminal or command prompt.
3. The script will process each text file and save the summaries to an Excel file named "summaries.xlsx".

## Credit

Created by [Jiajun Zou](https://github.com/jzou1995)
