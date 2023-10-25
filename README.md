# Gradebook Jana-Rator

> This project was created by Sid Chatterjee, Henry Harms, and Kelly Poon for the Nueva Software Engineering class fall semester taught by Wes Chao.

# About

`Gradebook Jana-Rator` is an automated system that reads highlighted rubrics and writes a score based off the rubric into Google Sheets.

> **Warning**
> The program will install a folder, change a spreadsheet, and automatically install mutliple dependencies

# Current Features

1. Downloads all files in a selected Google Drive folder as docx files into your libary
2. Reads all highlighted lines in the docx file and creates a score based on the highlighted lines
3. Writes the scores into the Google Sheets Gradebook

How to calculate the score:

- Fully highlighted in dark green = 1 point
- Fully highlited in light green = 0.5 points
- Half highlighted in dark green & half highlighted in light green = 0.75 points
- Half highlighted in dark green and half not highlighted = 0.5 points
- Half highlighted in light green and half not highlighted = 0.25 points

# Dependencies

This program will automatically install all the dependencies needed for this project. Make sure to have python3 downloaded.
These are the dependencies that will be downloaded.

- pydrive
- python-docx
- docx
- termcolor
- google-api

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

# Future Features/To-Dos (in order)

1. Testing

- Run unit tests and end-to-end tests. Do both white-box and black-box tests.

2. Create a front end/UI

- Create one button that Jana can push that runs all files (Jana wants the least amount of buttons and interactivity as possible)

3. Write scores into the Nubric based on the scores from the Gradebook

- Talk to Jana about how she converts the Gradebook score to Nubric scores
- Nubric template: https://docs.google.com/spreadsheets/d/1zhUjd6sQ7vHY6O8BgEyjRIbO-U5UTRG-mrrfDENCVKI/edit?usp=sharing

4. Generalize program for all/other Jana classes
