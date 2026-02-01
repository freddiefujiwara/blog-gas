# Google Doc to Markdown API

This project is a Google Apps Script. It turns Google Docs into Markdown text. It works as a Web App that you can call like an API.

The frontend source code that consumes this API can be found at: https://github.com/freddiefujiwara/blog

## What it does

- **List Files**: It shows a list of Google Doc IDs from a specific folder.
- **Convert to Markdown**: It takes a Google Doc ID and returns its content as Markdown. It supports:
  - Headings (like # Title)
  - Lists (Bulleted and Numbered)
  - Tables
  - Bold and Italic text
  - Links

## How to use

The script runs as a Web App.

1. **Get List of IDs**:
   Visit the Web App URL. It will give you a JSON list of all Google Docs in the folder.
2. **Get Markdown Content**:
   Visit `WEB_APP_URL?id=YOUR_DOC_ID`. It will give you the title and the Markdown content.

## Code Explanation (`src/Code.js`)

Here are the main parts of the code:

- `FOLDER_ID`: The ID of the Google Drive folder where your documents are.
- `doGet(e)`: This is the main function. It runs when you visit the Web App URL. It decides to either list files or convert a file based on the `id` parameter.
- `listDocIdsSortedByName_`: This function finds all Google Docs in your folder and sorts them by their names.
- `docBodyToMarkdown_`: This is the main engine. It reads the Google Doc and turns it into Markdown text.
- `elementToMarkdown_`: This part looks at each piece of the document (like a paragraph or a table) and decides how to convert it.
- `paragraphToMarkdown_`: This converts regular text and headings.
- `tableToMarkdown_`: This turns Google Doc tables into Markdown tables.
- `paragraphTextWithInlineStyles_`: This looks for **bold**, *italic*, and [links] inside the text.
- `json_`: A helper to send the data back in JSON format.

## How to setup and deploy

1. **Install dependencies**:
   ```bash
   npm install
   ```
2. **Build the project**:
   ```bash
   npm run build
   ```
   This creates the final code in the `dist/` folder.
3. **Deploy**:
   ```bash
   npm run deploy
   ```
   This pushes the code to your Google Apps Script project.

Note: You may need to add the "Google Drive API" to your Google Apps Script project for the folder listing to work.
