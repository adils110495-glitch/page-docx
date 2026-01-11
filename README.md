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

## File Structure

```
doc-generator/
├── composer.json          # PHP dependencies
├── docker-compose.yml     # Docker configuration
├── Dockerfile            # Docker image definition
├── index.php             # Main UI page
├── generator.php         # DOCX generation logic
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
