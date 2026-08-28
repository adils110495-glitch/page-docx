# Website to DOCX Generator

A Core PHP web application that converts website page content into properly formatted DOCX files.

## Features

- 🌐 **Split-screen UI** with directory browser (left) and form (right)
- 📁 **Project Organization** - Organize files into project-specific folders
- 📊 **Batch Processing** - Handle up to 100 URLs at once
- 📋 **Error Logging** - Automatic error logs for failed URLs
- 🎯 **Optional DIV/CSS class selector** for targeted content extraction
- 📄 **Full body extraction** when no selector is provided
- 📝 **Meta Title and Description** included in DOCX files
- 🎨 **Preserves HTML formatting** in generated documents
- 📦 **Slug-based filenames** derived from URLs
- 🔄 **Real-time directory browsing** - See generated files immediately
- 🔤 **Cyrillic Word Cleaner** - Strip Cyrillic/lookalike Unicode out of English text
- ⚡ **Core PHP** - No frameworks, lightweight and fast

## Requirements

- PHP ≥ 7.4
- Composer
- Docker & Docker Compose (for containerized deployment)

## Installation & Setup

### Using Docker (Recommended)

1. **Build and start the container:**
   ```bash
   docker-compose up -d --build
   ```

2. **Access the application:**
   Open your browser and navigate to:
   ```
   http://localhost:8083
   ```

3. **Stop the container:**
   ```bash
   docker-compose down
   ```

### Manual Installation

1. **Install dependencies:**
   ```bash
   composer install
   ```

2. **Configure your web server** to point to the project directory

3. **Ensure output directory is writable:**
   ```bash
   chmod 755 output
   ```

## Usage

### User Interface Layout

The application features a **split-screen interface**:
- **Left Sidebar**: Directory tree showing all generated files and projects
- **Right Panel**: Form for submitting URLs and configuration

### Step-by-Step Guide

1. **Browse Existing Files (Left Sidebar):**
   - View all generated DOCX files organized by project
   - Click on any file to download it
   - Click on log files (📋) to view error reports
   - Click folder names to expand/collapse

2. **Enter Project Name (Optional):**
   - Enter a project name to organize files into a dedicated folder
   - Example: `skycop-fr`
   - All generated files will be saved under `output/skycop-fr/`
   - If left empty, files will be saved in the root `output/` directory

3. **Enter URLs (Up to 100):**
   - Add one or multiple URLs in the textarea (max 100 per batch)
   - Each URL should be on a separate line
   - URLs must start with `http://` or `https://`

4. **Optional Selector:**
   - Enter a CSS class name (without the dot) to extract content from a specific DIV
   - Example: `your_right_contents`
   - If left empty, the entire `<body>` content will be extracted

5. **Generate:**
   - Click the "Generate DOCX" button
   - Processing happens in the background
   - A progress bar shows the status
   - If errors occur, an error log file will be generated automatically

6. **View Results:**
   - Success/error summary appears below the form
   - Generated files appear immediately in the left sidebar
   - Error log files can be clicked to view failed URLs

## How It Works

### Content Extraction

- **With Selector:** Extracts content from `<div class="your_selector">...</div>`
- **Without Selector:** Extracts full `<body>` content

### DOCX Structure

Each generated DOCX file contains (in order):

1. **Meta Title** (as main heading) - if available
2. **Meta Description** (italic paragraph) - if available
3. **Page Content** (with preserved HTML formatting)

### Filename Generation

Files are saved with slug-based names derived from the **last segment** of the URL path:
- `https://example.com/blog/my-post` → `my-post.docx`
- `https://example.com/abc/xyz/efg` → `efg.docx`
- `https://example.com/products/item-123.html` → `item-123.docx`
- `https://example.com/` → `example-com.docx` (uses hostname if no path)

**Filename Rules:**
- Uses only the last segment of the URL path
- Removes file extensions (.html, .php, .aspx)
- Converts to lowercase
- Replaces special characters with hyphens
- Maximum 100 characters

### Project Organization

When a project name is provided, files are organized into subdirectories:
- **Without project:** `output/my-post.docx`
- **With project "skycop-fr":** `output/skycop-fr/my-post.docx`

This allows you to keep files from different projects organized and separate.

### Batch Processing

The system can process up to **100 URLs** in a single batch:
- All URLs are processed sequentially
- Each URL is handled independently
- One failure doesn't stop processing of other URLs
- Progress is tracked and displayed
- Total processing time depends on number of URLs and page complexity

### Error Handling & Logging

The system provides comprehensive error handling:

- ✅ **Success:** DOCX generated and available for download in sidebar
- ⚠️ **Warning:** Selector not found, invalid content structure
- ❌ **Error:** Invalid URL, fetch failed, generation error

**Automatic Error Logs:**
- When errors occur, a log file is automatically created
- Log filename format: `errors_YYYY-MM-DD_HH-MM-SS.log`
- Logs are saved in the same directory as the DOCX files
- Each log contains:
  - Timestamp of each error
  - Failed URL
  - Error message/reason
  - Summary with total success/failure counts
  - List of all failed URLs for easy retry

**Accessing Logs:**
- Log files (📋) appear in the left sidebar directory tree
- Click any log file to open it in a new tab
- Copy failed URLs from the log to retry processing

## Cyrillic Word Cleaner

Route: **`/cyrillic-cleaner.php`** (linked from the nav bar on every page)

Cleans corrupted English text in which Latin letters have been swapped for Cyrillic
or other lookalike Unicode characters (`fоr`, `саn`, `ԁelаyeԁ`, `Prоtectiоns`, …).
It is a character/word substituter only — it never rewrites, rephrases, or reformats
the text.

### Dictionary pipeline

```
crilic-wordss.csv  →  cyrillic-dictionary.php  →  JSON replacement dictionary  →  cleaner engine (JS)
```

`crilic-wordss.csv` (project root) is the single source of truth. `cyrillic-dictionary.php`
parses it and serves it as JSON:

- The delimiter (tab / comma / semicolon / pipe) is detected automatically.
- The corrupted-word and correct-word columns are detected from the header names,
  falling back to content analysis if the headers are unfamiliar or missing.
  Columns such as *Occurrences*, *Cyrillic codepoints* and *Search-replace pair*
  are explicitly ignored.
- The result is cached in `output/.cyrillic-dictionary.json` and invalidated
  automatically whenever the CSV's modification time or size changes.

The CSV holds **two kinds of entry**, and rows of either kind may appear anywhere
in the file:

**1. Word mappings** — the main table, one corrupted word per row:

| Word as published | Occurrences | Cyrillic codepoints | Correct Latin form |
| ----------------- | ----------- | ------------------- | ------------------ |
| `fоr`             | 241         | о = U+043E          | `for`              |

**2. Character mappings** — a single lookalike and its Latin equivalent:

```
а → a
о → o
ԁ → d
```

`→`, `->`, `-->`, `=>` and `⟶` all work, as does a plain two-column
`а <tab> a`. These extend and override the cleaner's built-in character table,
so a newly discovered lookalike only has to be added here.

For safety the left side must be exactly one non-ASCII character and the right
side plain ASCII — a lookalike can never be rewritten into another Cyrillic
character, and a row like `a -> b` is refused.

**To add new mappings, edit `crilic-wordss.csv` — nothing else needs to change.**
Use the *Reload CSV* link in the tool to pick up changes without restarting.

Any line the parser cannot interpret is reported next to the dictionary status
("⚠ N lines not recognised", hover for the text), so an edit in an unexpected
format is never silently ignored.

### Two layers of detection

1. **CSV word mapping** — exact word replacements from `crilic-wordss.csv`,
   plus a case-insensitive fallback that preserves the original capitalisation.
2. **Generic Unicode mapping** — a character-level scan that fixes lookalike
   characters in words that are *not* in the CSV at all
   (`informаtion` → `information`, `аirline` → `airline`). It uses the
   character mappings from the CSV merged over a built-in table.

Covered scripts: Cyrillic, Greek, Armenian, Cherokee, Roman numeral forms,
fullwidth Latin and IPA lookalikes.

### Safety rules

- A word containing no Latin letter at all (e.g. real Russian text) is left
  untouched and reported instead of guessed at. The *Force-convert all-Cyrillic
  words* option overrides this.
- URLs, email addresses and HTML link attributes are never rewritten — a lookalike
  character there changes where the link points. They are reported instead.
  Toggle off with *Protect URLs & emails*.
- Anything that cannot be confidently mapped is counted under
  **Suspicious Characters Remaining** and listed with its codepoint.

### Rich text vs plain text

The tool has two modes, selected above the editors.

**Rich text (default)** — paste formatted content straight out of the WordPress
editor, Word or a web page. Only *text nodes* are touched: headings, bold, italics,
lists, blockquotes, tables, images, links, classes and inline styles come out exactly
as they went in — an `<h1>` in gives an `<h1>` out. Copying puts real HTML on the
clipboard, so pasting back into WordPress keeps the formatting.

The input pane has a formatting toolbar (paragraph / H1–H4 / quote / preformatted,
bold, italic, underline, bulleted and numbered lists, link, clear formatting), so
blocks can be applied by hand as well as pasted in.

Because attributes are never rewritten, a lookalike character inside an `href`
survives untouched and is reported under *Suspicious Characters Remaining* — changing
it would silently repoint the link.

Pasted markup is arbitrary web content, so `<script>`, `<style>`, `<iframe>`,
`javascript:` URLs and `on*` event handlers are stripped before the result is
rendered back into the page.

**Plain text** — a plain textarea in and out, for when formatting is irrelevant.
Paragraphs, line breaks, spacing and punctuation are preserved exactly.

### Features

- Side-by-side input/output editors (stacked on mobile)
- **Auto Clean** — debounced cleaning as you type or paste
- **Show Changes** — a table of every replacement with its type and count
- Statistics: characters replaced, words corrected, total replacements, suspicious remaining
- Copy Result / Clear
- Optional stripping of invisible characters (zero-width, soft hyphen, NBSP)
- **Clean up markup** (on by default) removes the `<div>` and `<span>` wrappers
  that editors and clipboards add, so what you paste into WordPress is just the
  headings, paragraphs, lists, links and emphasis

All processing happens locally in the browser — no API calls, no text leaves the page.

## File Structure

```
doc-generator/
├── composer.json          # PHP dependencies
├── docker-compose.yml     # Docker configuration
├── Dockerfile            # Docker image definition
├── index.php             # Main UI page
├── generator.php         # DOCX generation logic
├── crilic-wordss.csv     # Cyrillic → Latin word dictionary (source of truth)
├── cyrillic-cleaner.php  # Cyrillic Word Cleaner - UI + cleaning engine
├── cyrillic-dictionary.php # Parses crilic-wordss.csv into a JSON dictionary
├── output/               # Generated DOCX files
├── vendor/               # Composer dependencies
└── README.md            # This file
```

## Dependencies

- **phpoffice/phpword** - For generating DOCX files
- PHP DOM extension - For HTML parsing

## Technical Details

- **HTML Fetching:** Server-side with 30-second timeout
- **No JavaScript Execution:** Static HTML content only
- **HTML Parsing:** DOMDocument and DOMXPath
- **Format Preservation:** HTML-to-DOCX conversion maintains basic formatting

## Troubleshooting

### Permission Issues
```bash
chmod 755 output
chown -R www-data:www-data output
```

### Composer Dependencies
```bash
composer install --optimize-autoloader
```

### Docker Container Logs
```bash
docker-compose logs -f
```

## License

This project is open source and available for use.
