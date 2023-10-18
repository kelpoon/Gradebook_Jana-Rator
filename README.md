# Gradebook Jana-Rator

> For Jana!!!!

# About

`Gradebook Jana-Rator` is an automated system that reads highlighted rubrics and writes a score based off the rubric into Google Sheets.

# Dependencies

This program will automatically install all the dependencies needed for this project. Make sure to have python3 downloaded.
These are the dependencies that will be downloaded.

- pydrive
- python-docx
- docx
- termcolor
- google-api

> **Warning**
> The program will install a folder, change a spreadsheet, and automatically install mutliple dependencies

# API Keys

For this program, you will need to use Google Sheets API Keys and Google Drive API Keys. Download the keys from this Google Drive Folder: https://drive.google.com/drive/folders/1gPxW1DvohpBg_AMAYOhd2Ez_e8VGH2Nr?usp=share_link

# Usage

Make a copy of the Gradebook Template:
https://docs.google.com/spreadsheets/d/1PPI7-hc8vr4L_-Gpjdcrn_5qMwZVEXgVVkf6QiFQJRU/edit?usp=sharing

Create a Google Folder to store all of the rubrics. In the folder, make copies of the rubric:
https://docs.google.com/document/d/1H7N8oamMmQuj-0AObqRIpHwQUJdevsY8kCZ2mcbOkDU/edit?usp=sharing

In rubricReaderWrite.py, change the path for the google drive folder link and the spreadsheet link:

```
folderLink = 'CHANGE_THIS'
```

```
spreadsheetLink = 'CHANGE_THIS'
```

Then launch the file in terminal:

```
python3 rubricReaderWriter.py
```

# Next Steps
