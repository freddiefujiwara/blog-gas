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
  - Inline images (exported to a shared `images` folder in Drive and embedded as Markdown image links)

## Setup

You can use this code for your own blog. You do not need to change the code.
You only set some values in the project settings. These values are called
"Script Properties".

This setup is for beginners. You do everything in your web browser.
You do not need to install anything on your computer.

### Step 1: Make a folder

1. Open Google Drive.
2. Make a new folder.
3. Put your Google Docs in the folder. These Docs are your blog posts.
4. Open the folder. Look at the web address (URL) in your browser.
   The last part of the URL is the folder ID. Copy it.

### Step 2: Make the Apps Script project

1. Open [script.google.com](https://script.google.com).
   Sign in with the same Google account as your folder in Step 1.
2. Click **New project**.
3. The editor opens. There is a file called `Code.gs` with small sample code.
   Select all of the sample code and delete it.
4. Download the ready-made code file from this repository's
   **Releases** page:
   https://github.com/freddiefujiwara/blog-gas/releases/latest
   Download `Code.gs`, open it, and copy all of the code.
5. Go back to the editor. Paste the code into `Code.gs`.
6. Click the save icon (or press Ctrl+S, or Cmd+S on a Mac).
7. Click the project name ("Untitled project") at the top of the page.
   Give the project a name, for example `My Blog API`.

Note: The first time you deploy or run the project, Google asks you
if the project can access your data. This is normal.
The project runs under your own account. Follow the screens and allow it.

### Step 3: Set the Script Properties

1. Click **Project Settings** (the gear icon on the left).
2. Find **Script Properties** and click **Add script property**.
3. Add these four properties, one by one:

| Property | Value | What it is |
| --- | --- | --- |
| `FOLDER_ID` | The folder ID from Step 1 | The folder with your Docs |
| `SITE_TITLE` | Example: `My Blog` | The name of your blog. It is used in the RSS feed |
| `SITE_URL` | Example: `https://example.com/blog` | The web address of your blog. Do not add `/` at the end |
| `SITE_DESCRIPTION` | Example: `My blog posts` | A short text about your blog. It is used in the RSS feed |

4. Click **Save script properties**.

### Step 4: Deploy

1. Click **Deploy** and then **New deployment**.
2. Choose **Web app** as the type.
3. Set **Who has access** to **Anyone**.
4. Click **Deploy**. Now your API is ready.
5. Copy the **Web App URL**. It looks like this:
   `https://script.google.com/macros/s/xxxxx/exec`
   You need this URL when you set up the front side (the web page).

### Step 5: Set up the automatic updates (triggers)

The blog updates itself with two timers. They are called "triggers".

1. Click **Triggers** (the clock icon on the left).
2. Click **Add Trigger**.
   - Choose `preCacheAll` as the function.
   - Choose **Time-driven** as the event source.
   - Choose **Minutes timer** and **Every 10 minutes**.
   - Click **Save**.
3. Click **Add Trigger** again.
   - Choose `dailyRSSCache` as the function.
   - Choose **Time-driven** as the event source.
   - Choose **Day timer** and a time (for example, midnight to 1am).
   - Click **Save**.

The first timer keeps the list of posts new.
The second timer keeps the RSS feed new.

If you change a value later, only change the Script Property.
You do not need to change the code or deploy again.

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

3. **Get RSS Feed**:
   Visit `WEB_APP_URL?o=rss`. It returns an RSS 2.0 XML feed of the latest 50 documents.

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
- `paragraphTextWithInlineStyles_`: This looks for **bold**, *italic*, and [links] inside the text. Inline images are converted via `inlineImageToMarkdown_`.
- `getImageFolder_`: Finds (or creates) the `images` subfolder inside the documents folder, where exported images are stored.
- `exportInlineImage_`: Saves an inline image to the images folder under a content-hash file name (so the same image is exported only once), shares it as anyone-with-the-link, and returns its public URL.
- `inlineImageToMarkdown_`: Converts an inline image into Markdown image syntax using the exported URL and the image's alt description.
- `escapeMdInline_`: Escapes special characters like backslashes and backticks in inline text for Markdown.
- `escapeMdTable_`: Escapes pipes in table cells for Markdown.
- `json_`: A helper to send the data back in JSON format.
- `jsonError_`: A helper function to return error messages in JSON format.
- `saveLog_(msg)`: Appends a log message with a JST timestamp to the `DEBUG_LOGS` script property (limited to 9000 characters).
- `log_(msg)`: A wrapper that outputs a log message to both `console.log` and the `DEBUG_LOGS` script property.

## For developers: setup and deploy with your computer

This section is for people who write code. Beginners can use the Setup
section above and do not need this section.

Note: `dist/` is build output and is not kept in this repository.
Beginners download `Code.gs` from the GitHub Releases page instead,
so when the code changes, publish a new release with the new
`dist/Code.gs` file attached.

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
