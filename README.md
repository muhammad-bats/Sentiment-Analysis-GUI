# Sentiment-Analysis-GUI
GUI application for Sentiment Analysis

---

# Sentiment Analysis GUI

This Python GUI application enables users to perform sentiment analysis on PDF files within a selected folder using multiple sentiment analysis models. The analyzed results are then saved to an Excel file.

## Pre-Requisites 
For this python script you need these python liberaries

Python version: 3.10.14

tk:                       8.6.14

pypdf2:                   3.0.1

numpy:                    2.0.1 

openpyxl:                 3.1.5

nltk:                     3.8.1

vadersentiment:           3.3.2

textblob:                 0.15.3

flair:                    0.14.0

spacy:                    3.7.5

spacy-legacy:            3.0.12  

spacy-loggers:           1.0.5

torch:                   2.4.0

This script uses a conda environment to run the tracker. 

conda version: 24.7.1

To create a conda environment run the following command in anaconda prompt

`conda create --name pdf`

## Code Explanation

### 1. GUI Layout

The GUI is created using `tkinter`, which includes buttons, labels, and entry fields for user interaction. 

- **Upload Folder**: Allows the user to select a folder containing the PDF files.
- **File Name**: The name of the Excel file where the results will be saved.
- **Sheet Name**: The name of the sheet within the Excel file.
- **Model Select**: Dropdown menu to select the sentiment analysis model (VADER, TextBlob, Flair, SpaCy).
- **Analyze Sentiments**: Button to start the sentiment analysis process.

### 2. Selecting the Sentiment Analysis Model

The `select_model` function is used to choose the sentiment analysis model (VADER, TextBlob, Flair, or SpaCy). The selected model is stored in the `selected_model` variable for later use.

### 3. Extracting Text from PDF Files

The `process` function iterates through all PDF files in the selected folder. For each file, the text is extracted using `PyPDF2`, and the sentences are tokenized using NLTK.

### 4. Sentiment Analysis

Depending on the selected model, the `sentiment_analyze` function applies the respective sentiment analysis technique:

- **VADER**: Analyzes the sentiment of each sentence and calculates average positive, negative, neutral, and compound scores.
- **TextBlob**: Computes the polarity and subjectivity of each sentence and averages the results.
- **Flair**: Uses the Flair model to predict sentiment labels (POSITIVE/NEGATIVE) and confidence scores.
- **SpaCy**: Integrates TextBlob with SpaCy to analyze sentiment and return polarity and subjectivity scores.

The analysis results are then processed and formatted into lists for writing to the Excel file.

### 5. Writing Results to Excel

Two functions are used to write the results to an Excel file:

- **`write_file_excel`**: Writes the name of the processed file to the Excel sheet.
- **`write_excel`**: Writes the sentiment analysis results (category, scores, etc.) to the Excel sheet.

### 6. Error Handling

The `show_error` function is used to display error messages in case no folder or file is selected or if any other error occurs.

## How to Use

1. **Run the Script**: Start the GUI application by running the script.
2. **Upload Folder**: Click "Upload Folder" and select the folder containing the PDF files you want to analyze.
3. **Enter File and Sheet Name**: Provide the name for the Excel file and the sheet where the results will be saved.
4. **Select Model**: Click "Model Select" and choose one of the available sentiment analysis models.
5. **Analyze Sentiments**: Click "Analyze Sentiments!!" to start the sentiment analysis process. The results will be written to the specified Excel file.

## Example Usage

```bash
- Upload Folder: Select folder containing PDF files.
- Enter File Name: analysis_results.xlsx
- Enter Sheet Name: SentimentAnalysis
- Select Model: VADER
- Click "Analyze Sentiments!!" to start the analysis.
```

## Notes

- The program supports four sentiment analysis models: VADER, TextBlob, Flair, and SpaCy with TextBlob integration.
- The code is designed to automatically leave spaces in column to allow use of different models on the same folder and store the data in the same file for comparison between sentiment models
- To use different models on each folder and store the results in the same xlsx file for comparison, just select each model turn by turn and analyze individually
- Ensure that the PDFs contain text, as the script relies on text extraction.
- The code creates a new xlsx file if it doesnot exist
- The script creates a new xlsx sheet if one does not exist in the xlsx file.
- A sample result xlsx file is attched in the repository

---
