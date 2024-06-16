# GPT Personal Projects

This repository contains several projects that leverage the power of OpenAI's GPT-4 model. The projects aim to demonstrate the versatility and capabilities of GPT-4 in various applications.

## Project 1: Transforming PDFs with Questions into Excel

The first project in this repository is a Python script named `generate_excel_from_pdfs.py`. This script is designed to extract questions from PDF files and transform them into an Excel format.

### How it works

The script uses the `pdfminer` library to extract text from the PDF files. It then uses the Azure OpenAI GPT-4 model to identify and extract questions and their possible answers from the text.

The extracted questions and answers are then formatted and written into an Excel file using the `openpyxl` library.

### Usage

To use the script, you need to have the following environment variables set:

- `AZURE_OPENAI_DEPLOYMENT`: The deployment ID of your Azure OpenAI instance.
- `AZURE_OPENAI_ENDPOINT`: The endpoint URL of your Azure OpenAI instance.
- `AZURE_OPENAI_API_KEY`: The API key for your Azure OpenAI instance.

Once these are set, you can run the script with a path to a PDF file as an argument. The script will then extract the questions and answers from the PDF and write them into an Excel file.

## Future Projects

Stay tuned for more exciting projects that demonstrate the power of GPT-4 in various applications.