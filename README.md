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

1. **Get List of IDs and Preloaded Articles**:
   Visit the Web App URL without any parameters. It returns a JSON object containing all document IDs and the content of the first 10 documents (for faster loading).

   **Response Structure**:
   ```json
   {
     "ids": ["ID1", "ID2", "..."],
     "article_cache": [
       {
         "id": "ID1",
         "title": "Document Title",
         "markdown": "Markdown content..."
       },
       ... up to 10 items
     ]
   }
   ```

2. **Get Markdown Content for a specific Document**:
   Visit `WEB_APP_URL?id=YOUR_DOC_ID`. It returns the title and Markdown content for the requested ID.

   **Response Structure**:
   ```json
   {
     "id": "YOUR_DOC_ID",
     "title": "Document Title",
     "markdown": "Markdown content..."
   }
   ```

## API Specification

The API is documented using the OpenAPI Specification. You can find the definition in [openapi.yaml](./openapi.yaml).

## Code Explanation (`src/Code.js`)

Here are the main parts of the code:

- `FOLDER_ID`: The ID of the Google Drive folder where your documents are.
- `CACHE_TTL`: Cache time-to-live in seconds (default: 600).
- `CACHE_SIZE_LIMIT`: Maximum size of a cache entry in characters (default: 100,000).
- `preCacheAll()`: A batch processing function intended to be run via a time-driven trigger. It clears old debug logs, saves the list of documents and the content of the top 10 documents to script properties to improve API performance.
- `clearCacheAll()`: Clears all cache entries used by the application, including the document list and individual document contents.
- `doGet(e)`: This is the main function for the Web API. It retrieves content from script properties if available; otherwise, it fetches the data directly and saves it to properties for future use. It decides to either list files or convert a file based on the `id` parameter.
- `listDocIdsSortedByTitle_`: This function finds all Google Docs in your folder and sorts them by their titles.
- `getDocInfoInFolder_`: Verifies if a document exists within the specified folder and retrieves its metadata.
- `docBodyToMarkdown_`: This is the main engine. It reads the Google Doc and turns it into Markdown text.
- `elementToMarkdown_`: This part looks at each piece of the document (like a paragraph or a table) and decides how to convert it.
- `paragraphToMarkdown_`: This converts regular text and headings.
- `listItemToMarkdown_`: Converts Google Doc list items (bulleted or numbered) into Markdown format.
- `tableToMarkdown_`: This turns Google Doc tables into Markdown tables.
- `headingToPrefix_`: Helper to convert paragraph heading levels to Markdown prefixes (#, ##, etc.).
- `isOrderedGlyph_`: Helper to determine if a list item's glyph type suggests an ordered list.
- `paragraphTextWithInlineStyles_`: This looks for **bold**, *italic*, and [links] inside the text.
- `escapeMdInline_`: Escapes special characters like backslashes and backticks in inline text for Markdown.
- `escapeMdTable_`: Escapes pipes in table cells for Markdown.
- `json_`: A helper to send the data back in JSON format.
- `jsonError_`: A helper function to return error messages in JSON format.
- `saveLog_(msg)`: Appends a log message with a JST timestamp to the `DEBUG_LOGS` script property (limited to 9000 characters).
- `log_(msg)`: A wrapper that outputs a log message to both `console.log` and the `DEBUG_LOGS` script property.

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
4. **Set up Trigger**:
   To pre-populate the script properties, you should set up a time-driven trigger in the Google Apps Script editor:
   - Go to **Triggers** (clock icon on the left).
   - Click **Add Trigger**.
   - Select `preCacheAll` as the function to run.
   - Select **Time-driven** as the event source.
   - Select **Minutes timer** and **Every 10 minutes** as the type of time-based trigger.

Note: You may need to add the "Google Drive API" to your Google Apps Script project for the folder listing to work.
