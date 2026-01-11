<?php
session_start();
require_once 'vendor/autoload.php';

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Shared\Html;

/**
 * Generate slug from URL for filename
 * Uses only the last segment of the URL path
 * Example: /abc/xyz/efg -> efg
 */
function generateSlug($url) {
    $parsed = parse_url($url);
    $path = isset($parsed['path']) ? $parsed['path'] : '';

    // Remove trailing slash
    $path = rtrim($path, '/');

    // Get the last segment of the path
    if (!empty($path)) {
        $segments = explode('/', $path);
        $lastSegment = end($segments);

        // If last segment is empty or just a slash, use the previous segment
        if (empty($lastSegment)) {
            array_pop($segments);
            $lastSegment = end($segments);
        }

        // If we have a valid last segment, use it
        if (!empty($lastSegment)) {
            $slug = $lastSegment;
        } else {
            // Fallback to hostname if no path segments
            $slug = isset($parsed['host']) ? $parsed['host'] : 'document';
        }
    } else {
        // No path, use hostname
        $slug = isset($parsed['host']) ? $parsed['host'] : 'document';
    }

    // Remove file extension if present (e.g., .html, .php)
    $slug = preg_replace('/\.(html?|php|aspx?)$/i', '', $slug);

    // Remove special characters and convert to lowercase
    $slug = preg_replace('/[^a-z0-9]+/i', '-', $slug);
    $slug = trim($slug, '-');
    $slug = strtolower($slug);

    // Limit length
    if (strlen($slug) > 100) {
        $slug = substr($slug, 0, 100);
    }

    return $slug ?: 'document';
}

/**
 * Fetch HTML content from URL
 */
function fetchHtml($url, $timeout = 30) {
    $context = stream_context_create([
        'http' => [
            'timeout' => $timeout,
            'user_agent' => 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'follow_location' => true,
        ],
        'ssl' => [
            'verify_peer' => false,
            'verify_peer_name' => false,
        ]
    ]);

    $html = @file_get_contents($url, false, $context);

    if ($html === false) {
        return null;
    }

    return $html;
}

/**
 * Extract meta title from HTML
 */
function extractMetaTitle($dom) {
    $titles = $dom->getElementsByTagName('title');
    if ($titles->length > 0) {
        return trim($titles->item(0)->textContent);
    }
    return null;
}

/**
 * Extract meta description from HTML
 */
function extractMetaDescription($dom) {
    $xpath = new DOMXPath($dom);
    $metaTags = $xpath->query('//meta[@name="description"]');

    if ($metaTags->length > 0) {
        return trim($metaTags->item(0)->getAttribute('content'));
    }
    return null;
}

/**
 * Remove elements matching skip selectors from HTML
 */
function removeSkipSelectors($html, $skipSelectors) {
    if (empty($skipSelectors)) {
        return $html;
    }

    // Parse the HTML
    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
    libxml_clear_errors();

    $xpath = new DOMXPath($dom);

    // Parse skip selectors (comma-separated)
    $selectors = array_map('trim', explode(',', $skipSelectors));

    foreach ($selectors as $selector) {
        if (empty($selector)) continue;

        $nodesToRemove = [];

        // Handle different selector types
        if (strpos($selector, '#') === 0) {
            // ID selector (e.g., #sidebar)
            $id = substr($selector, 1);
            $nodes = $xpath->query("//*[@id='{$id}']");
            foreach ($nodes as $node) {
                $nodesToRemove[] = $node;
            }
        } elseif (strpos($selector, '.') === 0) {
            // Class selector (e.g., .header)
            $class = substr($selector, 1);
            $nodes = $xpath->query("//*[contains(concat(' ', normalize-space(@class), ' '), ' {$class} ')]");
            foreach ($nodes as $node) {
                $nodesToRemove[] = $node;
            }
        } else {
            // Element name or class without dot (e.g., header, nav, or sidebar)
            // Try as element name first
            $nodes = $xpath->query("//{$selector}");
            foreach ($nodes as $node) {
                $nodesToRemove[] = $node;
            }

            // Also try as class name
            $nodes = $xpath->query("//*[contains(concat(' ', normalize-space(@class), ' '), ' {$selector} ')]");
            foreach ($nodes as $node) {
                $nodesToRemove[] = $node;
            }
        }

        // Remove the nodes
        foreach ($nodesToRemove as $node) {
            if ($node->parentNode) {
                $node->parentNode->removeChild($node);
            }
        }
    }

    // Get the cleaned HTML
    $body = $dom->getElementsByTagName('body')->item(0);
    if ($body) {
        return getInnerHtml($body);
    }

    return $html;
}

/**
 * Extract content from HTML based on selector
 */
function extractContent($html, $selector = null, $skipSelectors = '') {
    $dom = new DOMDocument();

    // Suppress warnings for malformed HTML
    libxml_use_internal_errors(true);
    $dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
    libxml_clear_errors();

    $xpath = new DOMXPath($dom);

    // Extract meta information first
    $metaTitle = extractMetaTitle($dom);
    $metaDescription = extractMetaDescription($dom);

    // Extract content based on selector
    $contentHtml = '';

    if ($selector && !empty(trim($selector))) {
        // Try to find div with specific class
        $nodes = $xpath->query("//div[contains(concat(' ', normalize-space(@class), ' '), ' $selector ')]");

        if ($nodes->length > 0) {
            // Get inner HTML of the first matching div
            $node = $nodes->item(0);
            $contentHtml = getInnerHtml($node);
        } else {
            return [
                'success' => false,
                'error' => 'Selector not found',
                'metaTitle' => $metaTitle,
                'metaDescription' => $metaDescription
            ];
        }
    } else {
        // Extract full body content
        $bodyNodes = $dom->getElementsByTagName('body');
        if ($bodyNodes->length > 0) {
            $contentHtml = getInnerHtml($bodyNodes->item(0));
        } else {
            return [
                'success' => false,
                'error' => 'No body content found',
                'metaTitle' => $metaTitle,
                'metaDescription' => $metaDescription
            ];
        }
    }

    // Remove skip selectors if provided
    if (!empty($skipSelectors)) {
        $contentHtml = removeSkipSelectors($contentHtml, $skipSelectors);
        debugLog("  After removing skip selectors: " . strlen($contentHtml) . " bytes");
    }

    debugLog("  Extraction result - HTML length: " . strlen($contentHtml) . " bytes, Title: " . ($metaTitle ?: 'none'));

    return [
        'success' => true,
        'html' => $contentHtml,
        'metaTitle' => $metaTitle,
        'metaDescription' => $metaDescription
    ];
}

/**
 * Get inner HTML of a DOMNode
 */
function getInnerHtml($node) {
    $innerHTML = '';
    $children = $node->childNodes;

    foreach ($children as $child) {
        $innerHTML .= $node->ownerDocument->saveHTML($child);
    }

    return $innerHTML;
}

/**
 * Process DOM node and add content to DOCX section
 * Recursively walks through DOM and adds formatted text
 */
function processNodeForDocx($section, $node, $textRun = null) {
    foreach ($node->childNodes as $child) {
        $nodeName = strtolower($child->nodeName);
        $nodeValue = trim($child->nodeValue);

        // Handle text nodes
        if ($child->nodeType === XML_TEXT_NODE) {
            if (!empty($nodeValue)) {
                if ($textRun) {
                    $textRun->addText($nodeValue);
                } else {
                    // Direct text without parent formatting element
                    $section->addText($nodeValue, ['size' => 11, 'name' => 'Arial']);
                }
            }
            continue;
        }

        // Handle element nodes
        if ($child->nodeType === XML_ELEMENT_NODE) {
            switch ($nodeName) {
                case 'h1':
                case 'h2':
                case 'h3':
                case 'h4':
                case 'h5':
                case 'h6':
                    $sizes = ['h1' => 18, 'h2' => 16, 'h3' => 14, 'h4' => 13, 'h5' => 12, 'h6' => 11];
                    $text = getTextContent($child);
                    if (!empty($text)) {
                        $section->addText(
                            $text,
                            ['bold' => true, 'size' => $sizes[$nodeName], 'name' => 'Arial'],
                            ['spaceAfter' => 240]
                        );
                    }
                    break;

                case 'p':
                    $text = getTextContent($child);
                    if (!empty($text)) {
                        $section->addText(
                            $text,
                            ['size' => 11, 'name' => 'Arial'],
                            ['spaceAfter' => 200]
                        );
                    }
                    break;

                case 'strong':
                case 'b':
                    $text = getTextContent($child);
                    if (!empty($text) && $textRun) {
                        $textRun->addText($text, ['bold' => true]);
                    }
                    break;

                case 'em':
                case 'i':
                    $text = getTextContent($child);
                    if (!empty($text) && $textRun) {
                        $textRun->addText($text, ['italic' => true]);
                    }
                    break;

                case 'ul':
                case 'ol':
                    processListForDocx($section, $child, $nodeName);
                    break;

                case 'br':
                    if ($textRun) {
                        $textRun->addTextBreak();
                    }
                    break;

                case 'div':
                case 'section':
                case 'article':
                case 'main':
                    // Recursively process container elements
                    processNodeForDocx($section, $child, $textRun);
                    break;

                default:
                    // For other elements, just extract text content
                    if ($child->hasChildNodes()) {
                        processNodeForDocx($section, $child, $textRun);
                    }
                    break;
            }
        }
    }
}

/**
 * Process list elements (ul/ol) for DOCX
 */
function processListForDocx($section, $listNode, $listType) {
    $depth = 0;
    foreach ($listNode->childNodes as $child) {
        if (strtolower($child->nodeName) === 'li') {
            $text = getTextContent($child);
            if (!empty($text)) {
                $section->addListItem(
                    $text,
                    $depth,
                    ['size' => 11, 'name' => 'Arial'],
                    $listType === 'ol' ? ['listType' => \PhpOffice\PhpWord\Style\ListItem::TYPE_NUMBER] : null,
                    ['spaceAfter' => 120]
                );
            }
        }
    }
}

/**
 * Get all text content from a DOM node
 */
function getTextContent($node) {
    $text = '';
    foreach ($node->childNodes as $child) {
        if ($child->nodeType === XML_TEXT_NODE) {
            $text .= $child->nodeValue;
        } elseif ($child->hasChildNodes()) {
            $text .= getTextContent($child);
        }
    }
    return trim($text);
}

/**
 * Generate DOCX file from HTML content
 */
function generateDocx($content, $filename, $project = null) {
    // Suppress warnings from PHPWord HTML parser
    $oldErrorReporting = error_reporting();
    error_reporting($oldErrorReporting & ~E_WARNING);

    $phpWord = new PhpWord();
    $section = $phpWord->addSection();

    // Add Meta Title if available
    if (!empty($content['metaTitle'])) {
        debugLog("  Adding meta title: " . substr($content['metaTitle'], 0, 50));
        $section->addText(
            htmlspecialchars_decode($content['metaTitle'], ENT_QUOTES),
            ['bold' => true, 'size' => 18, 'name' => 'Arial'],
            ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT, 'spaceAfter' => 240]
        );
    } else {
        debugLog("  No meta title found");
    }

    // Add Meta Description if available
    if (!empty($content['metaDescription'])) {
        debugLog("  Adding meta description: " . substr($content['metaDescription'], 0, 50));
        $section->addText(
            htmlspecialchars_decode($content['metaDescription'], ENT_QUOTES),
            ['italic' => true, 'size' => 11, 'name' => 'Arial', 'color' => '666666'],
            ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT, 'spaceAfter' => 360]
        );
    } else {
        debugLog("  No meta description found");
    }

    // Add main content
    if (!empty($content['html'])) {
        debugLog("  Adding content to DOCX. HTML length: " . strlen($content['html']) . " bytes");

        // Clean HTML for better processing
        $cleanHtml = cleanHtmlForDocx($content['html']);
        debugLog("  After cleaning: " . strlen($cleanHtml) . " bytes");

        // Convert HTML to formatted text for DOCX
        // PHPWord's HTML parser has limitations, so we'll extract and format text properly
        $dom = new DOMDocument();
        libxml_use_internal_errors(true);
        $dom->loadHTML(mb_convert_encoding($cleanHtml, 'HTML-ENTITIES', 'UTF-8'));
        libxml_clear_errors();

        // Process the DOM and add content with formatting
        $body = $dom->getElementsByTagName('body')->item(0);
        if ($body) {
            processNodeForDocx($section, $body);
            debugLog("  Content added to DOCX by processing DOM nodes");
        } else {
            // Ultimate fallback
            $textContent = strip_tags($cleanHtml);
            $paragraphs = preg_split('/\n\s*\n/', trim($textContent));
            foreach ($paragraphs as $para) {
                $para = trim($para);
                if (!empty($para)) {
                    $section->addText($para, ['size' => 11, 'name' => 'Arial'], ['spaceAfter' => 200]);
                }
            }
            debugLog("  Content added as plain text fallback");
        }
    } else {
        debugLog("  WARNING: No HTML content to add!");
    }

    // Determine output directory
    $outputDir = __DIR__ . '/output';

    // If project name is provided, create project subdirectory
    if (!empty($project)) {
        $projectSlug = sanitizeProjectName($project);
        $outputDir .= '/' . $projectSlug;
    }

    // Create directory if it doesn't exist
    if (!is_dir($outputDir)) {
        mkdir($outputDir, 0777, true);
        chmod($outputDir, 0777);
    }

    $filepath = $outputDir . '/' . $filename . '.docx';

    $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
    $objWriter->save($filepath);

    // Restore error reporting
    error_reporting($oldErrorReporting);

    return $filepath;
}

/**
 * Sanitize project name for use as directory name
 */
function sanitizeProjectName($project) {
    // Remove special characters and convert to lowercase
    $slug = preg_replace('/[^a-z0-9\-_]+/i', '-', $project);
    $slug = trim($slug, '-');
    $slug = strtolower($slug);

    // Limit length
    if (strlen($slug) > 50) {
        $slug = substr($slug, 0, 50);
    }

    return $slug ?: 'default';
}

/**
 * Clean HTML for better DOCX conversion
 */
function cleanHtmlForDocx($html) {
    // Remove script and style tags
    $html = preg_replace('/<script\b[^>]*>(.*?)<\/script>/is', '', $html);
    $html = preg_replace('/<style\b[^>]*>(.*?)<\/style>/is', '', $html);

    // Remove comments
    $html = preg_replace('/<!--(.|\s)*?-->/', '', $html);

    // Remove problematic tags that might cause null node issues
    $html = preg_replace('/<svg\b[^>]*>(.*?)<\/svg>/is', '', $html);
    $html = preg_replace('/<noscript\b[^>]*>(.*?)<\/noscript>/is', '', $html);
    $html = preg_replace('/<iframe\b[^>]*>(.*?)<\/iframe>/is', '', $html);

    // Remove empty tags that can cause issues
    $html = preg_replace('/<(\w+)[^>]*>\s*<\/\1>/', '', $html);

    // Convert common HTML entities
    $html = html_entity_decode($html, ENT_QUOTES | ENT_HTML5, 'UTF-8');

    // Trim whitespace
    $html = trim($html);

    return $html;
}

/**
 * Validate URL
 */
function isValidUrl($url) {
    $url = trim($url);

    if (empty($url)) {
        return false;
    }

    // Check if URL starts with http or https
    if (!preg_match('/^https?:\/\//i', $url)) {
        return false;
    }

    // Validate URL format
    return filter_var($url, FILTER_VALIDATE_URL) !== false;
}

/**
 * Process single URL
 */
function processUrl($url, $selector, $project = null, $skipSelectors = '') {
    // Validate URL
    if (!isValidUrl($url)) {
        debugLog("  Invalid URL format");
        return [
            'type' => 'error',
            'message' => 'Invalid URL format',
            'url' => $url
        ];
    }

    // Fetch HTML
    debugLog("  Fetching HTML...");
    $html = fetchHtml($url);

    if ($html === null) {
        debugLog("  Failed to fetch HTML");
        return [
            'type' => 'error',
            'message' => 'Failed to fetch URL',
            'url' => $url
        ];
    }
    debugLog("  HTML fetched: " . strlen($html) . " bytes");

    // Extract content
    debugLog("  Extracting content...");
    $extracted = extractContent($html, $selector, $skipSelectors);

    if (!$extracted['success']) {
        debugLog("  Extraction failed: " . $extracted['error']);
        return [
            'type' => 'warning',
            'message' => 'Skipped: ' . $extracted['error'],
            'url' => $url
        ];
    }
    debugLog("  Content extracted: " . strlen($extracted['html']) . " bytes");

    // Generate slug for filename
    $slug = generateSlug($url);
    debugLog("  Slug: $slug");

    // Generate DOCX
    try {
        debugLog("  Generating DOCX...");
        $filepath = generateDocx($extracted, $slug, $project);
        debugLog("  DOCX saved to: $filepath");

        // Build relative filepath
        if (!empty($project)) {
            $projectSlug = sanitizeProjectName($project);
            $relativeFilepath = 'output/' . $projectSlug . '/' . basename($filepath);
        } else {
            $relativeFilepath = 'output/' . basename($filepath);
        }

        return [
            'type' => 'success',
            'message' => 'Successfully generated DOCX',
            'url' => $url,
            'file' => $relativeFilepath
        ];
    } catch (Exception $e) {
        return [
            'type' => 'error',
            'message' => 'Failed to generate DOCX: ' . $e->getMessage(),
            'url' => $url
        ];
    }
}

/**
 * Create log file for errors
 */
function createLogFile($project = null) {
    $outputDir = __DIR__ . '/output';

    if (!empty($project)) {
        $projectSlug = sanitizeProjectName($project);
        $outputDir .= '/' . $projectSlug;
    }

    if (!is_dir($outputDir)) {
        mkdir($outputDir, 0777, true);
    }

    $timestamp = date('Y-m-d_H-i-s');
    $logFilename = 'errors_' . $timestamp . '.log';
    $logPath = $outputDir . '/' . $logFilename;

    return $logPath;
}

/**
 * Write to log file
 */
function writeLog($logPath, $message) {
    $timestamp = date('Y-m-d H:i:s');
    $logMessage = "[{$timestamp}] {$message}\n";
    file_put_contents($logPath, $logMessage, FILE_APPEND);
}

// Debug logging function
function debugLog($message) {
    $logFile = __DIR__ . '/output/debug.log';
    $timestamp = date('Y-m-d H:i:s');
    @file_put_contents($logFile, "[{$timestamp}] {$message}\n", FILE_APPEND);
}

// Main processing
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    debugLog("=== Starting new processing request ===");
    $urlsInput = isset($_POST['urls']) ? $_POST['urls'] : '';
    $selector = isset($_POST['selector']) ? trim($_POST['selector']) : '';
    $skipSelectors = isset($_POST['skip_selectors']) ? trim($_POST['skip_selectors']) : '';
    $project = isset($_POST['project']) ? trim($_POST['project']) : '';
    debugLog("Project: " . ($project ?: 'none') . ", Selector: " . ($selector ?: 'none') . ", Skip: " . ($skipSelectors ?: 'none'));

    // Parse URLs (one per line)
    $urls = array_filter(
        array_map('trim', explode("\n", $urlsInput)),
        function($url) {
            return !empty($url);
        }
    );

    if (empty($urls)) {
        $_SESSION['status'] = [
            'type' => 'error',
            'message' => 'No URLs provided'
        ];
        header('Location: index.php');
        exit;
    }

    // Limit to 100 URLs
    if (count($urls) > 100) {
        $urls = array_slice($urls, 0, 100);
        $_SESSION['status'] = [
            'type' => 'error',
            'message' => 'Maximum 100 URLs allowed. Only first 100 URLs will be processed.'
        ];
    }

    // Initialize counters and log file
    $totalUrls = count($urls);
    $successCount = 0;
    $errorCount = 0;
    $logPath = null;
    $errors = [];

    // Process each URL
    foreach ($urls as $index => $url) {
        debugLog("Processing URL " . ($index + 1) . "/{$totalUrls}: $url");
        $result = processUrl($url, $selector, $project, $skipSelectors);
        debugLog("Result type: " . $result['type'] . ", Message: " . $result['message']);

        if ($result['type'] === 'success') {
            $successCount++;
            if (isset($result['file'])) {
                debugLog("File created: " . $result['file']);
            }
        } else {
            $errorCount++;

            // Create log file on first error
            if ($logPath === null) {
                $logPath = createLogFile($project);
                writeLog($logPath, "=== DOCX Generation Error Log ===");
                writeLog($logPath, "Project: " . ($project ?: 'No project'));
                writeLog($logPath, "Selector: " . ($selector ?: 'Full body'));
                writeLog($logPath, "Total URLs: {$totalUrls}");
                writeLog($logPath, "=====================================\n");
            }

            // Log the error
            $errorMsg = "URL: {$url}\nError: {$result['message']}\n";
            writeLog($logPath, $errorMsg);
            $errors[] = $url;
        }
    }

    // Prepare status message
    $statusMessage = "Processed {$totalUrls} URLs: {$successCount} successful, {$errorCount} failed";

    $status = [
        'type' => $errorCount > 0 ? 'error' : 'success',
        'message' => $statusMessage,
        'processed' => $totalUrls,
        'total' => $totalUrls
    ];

    if ($logPath !== null) {
        // Make log path relative
        $relativeLogPath = str_replace(__DIR__ . '/', '', $logPath);
        $status['log_file'] = $relativeLogPath;

        // Write summary to log
        writeLog($logPath, "\n=== Summary ===");
        writeLog($logPath, "Total URLs: {$totalUrls}");
        writeLog($logPath, "Successful: {$successCount}");
        writeLog($logPath, "Failed: {$errorCount}");
        writeLog($logPath, "\n=== Failed URLs List ===");
        foreach ($errors as $errorUrl) {
            writeLog($logPath, $errorUrl);
        }
    }

    $_SESSION['status'] = $status;

    // Redirect back to index
    header('Location: index.php');
    exit;
} else {
    // If accessed directly, redirect to index
    header('Location: index.php');
    exit;
}
