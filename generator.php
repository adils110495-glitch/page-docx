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
    debugLog("[SKIP] Processing skip selectors: " . implode(', ', $selectors));

    foreach ($selectors as $selector) {
        if (empty($selector)) continue;

        $nodesToRemove = [];

        // Handle different selector types
        if (strpos($selector, '#') === 0) {
            // ID selector (e.g., #sidebar)
            $id = substr($selector, 1);
            $nodes = $xpath->query("//*[@id='{$id}']");
            debugLog("[SKIP] Looking for ID selector: #{$id}, found: " . $nodes->length . " elements");
            foreach ($nodes as $node) {
                debugLog("[SKIP] Removing element: <" . $node->nodeName . "> with id='{$id}'");
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
            debugLog("[SKIP] Looking for element name: {$selector}, found: " . $nodes->length . " elements");
            foreach ($nodes as $node) {
                debugLog("[SKIP] Removing element by name: <{$node->nodeName}>");
                $nodesToRemove[] = $node;
            }

            // Also try as class name
            $nodes = $xpath->query("//*[contains(concat(' ', normalize-space(@class), ' '), ' {$selector} ')]");
            debugLog("[SKIP] Looking for class name: {$selector}, found: " . $nodes->length . " elements");
            foreach ($nodes as $node) {
                debugLog("[SKIP] Removing element by class: <{$node->nodeName}> with class='{$node->getAttribute('class')}'");
                $nodesToRemove[] = $node;
            }

            // Also try as ID without # prefix
            $nodes = $xpath->query("//*[@id='{$selector}']");
            debugLog("[SKIP] Looking for ID (without #): {$selector}, found: " . $nodes->length . " elements");
            foreach ($nodes as $node) {
                debugLog("[SKIP] Removing element by ID: <{$node->nodeName}> with id='{$selector}'");
                $nodesToRemove[] = $node;
            }
        }

        // Remove the nodes
        $removedCount = 0;
        foreach ($nodesToRemove as $node) {
            if ($node->parentNode) {
                $node->parentNode->removeChild($node);
                $removedCount++;
            }
        }
        if ($removedCount > 0) {
            debugLog("[SKIP] Removed {$removedCount} node(s) for selector: {$selector}");
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
function processNodeForDocx($section, $node, $textRun = null, $depth = 0) {
    foreach ($node->childNodes as $child) {
        $nodeName = strtolower($child->nodeName);
        $nodeValue = trim($child->nodeValue);

        // Debug: Log every element node we encounter
        if ($child->nodeType === XML_ELEMENT_NODE && in_array($nodeName, ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])) {
            $childClass = $child->hasAttribute('class') ? $child->getAttribute('class') : 'no-class';
            debugLog("  [DEPTH $depth] Found element: {$nodeName} (class: " . substr($childClass, 0, 40) . ")");
        }

        // Handle text nodes
        if ($child->nodeType === XML_TEXT_NODE) {
            $trimmedValue = trim($nodeValue);
            // Only add text if it's not just whitespace
            if (!empty($trimmedValue)) {
                if ($textRun) {
                    $textRun->addText($nodeValue);
                } else {
                    // Direct text without parent formatting element - only add if substantial
                    if (strlen($trimmedValue) > 2) {
                        $section->addText($trimmedValue, ['size' => 11, 'name' => 'Arial']);
                    }
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
                        // Get the class attribute if it exists
                        $headingClass = $child->hasAttribute('class') ? $child->getAttribute('class') : '';

                        // Decode HTML entities to prevent issues with special characters
                        $text = html_entity_decode($text, ENT_QUOTES | ENT_HTML5, 'UTF-8');
                        debugLog("  Adding heading {$nodeName}" . ($headingClass ? " (class: " . substr($headingClass, 0, 30) . ")" : "") . ": " . substr($text, 0, 50));

                        // Use slightly larger size for accordion titles
                        $size = $sizes[$nodeName];
                        if (strpos($headingClass, 'accordion__title') !== false && $nodeName === 'h3') {
                            $size = 13; // Make FAQ questions more prominent
                        }

                        $section->addText(
                            $text,
                            ['bold' => true, 'size' => $size, 'name' => 'Arial'],
                            ['spaceAfter' => 240]
                        );
                    } else {
                        debugLog("  Empty heading {$nodeName} skipped");
                    }
                    break;

                case 'p':
                    $text = getTextContent($child);
                    if (!empty($text)) {
                        // Decode HTML entities to prevent issues with special characters
                        $text = html_entity_decode($text, ENT_QUOTES | ENT_HTML5, 'UTF-8');
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

                case 'table':
                    processTableForDocx($section, $child);
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
                    // Check if div has class that indicates it's a heading
                    $divClass = $child->hasAttribute('class') ? $child->getAttribute('class') : '';

                    // Map common heading-like classes to heading styles
                    $headingClasses = [
                        'title1' => ['size' => 16, 'bold' => true],  // Large heading
                        'title2' => ['size' => 14, 'bold' => true],  // Medium heading
                        'title3' => ['size' => 13, 'bold' => true],  // Small heading
                        'your-rights-faq__question' => ['size' => 13, 'bold' => true],  // FAQ questions
                        'your-rights-compensation__title' => ['size' => 16, 'bold' => true],
                        'bordered-card__title' => ['size' => 14, 'bold' => true],
                    ];

                    $isHeadingDiv = false;
                    $headingStyle = null;

                    foreach ($headingClasses as $className => $style) {
                        if (strpos($divClass, $className) !== false) {
                            $isHeadingDiv = true;
                            $headingStyle = $style;
                            break;
                        }
                    }

                    if ($isHeadingDiv && $headingStyle) {
                        // Treat this div as a heading
                        $text = getTextContent($child);
                        if (!empty($text)) {
                            $text = html_entity_decode($text, ENT_QUOTES | ENT_HTML5, 'UTF-8');
                            debugLog("  Adding div heading (class: " . substr($divClass, 0, 30) . "): " . substr($text, 0, 50));
                            $section->addText(
                                $text,
                                array_merge(['name' => 'Arial'], $headingStyle),
                                ['spaceAfter' => 240]
                            );
                        }
                    } else {
                        // Recursively process container elements
                        processNodeForDocx($section, $child, $textRun, $depth + 1);
                    }
                    break;

                default:
                    // For other elements, just extract text content
                    if ($child->hasChildNodes()) {
                        processNodeForDocx($section, $child, $textRun, $depth + 1);
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
            // Check if LI contains heading elements (H1-H6) at any depth
            $hasHeading = containsHeading($child);

            // If LI contains headings, process it recursively to preserve heading formatting
            if ($hasHeading) {
                debugLog("  [LIST] LI contains headings, processing recursively");
                processNodeForDocx($section, $child, null, 0);
            } else {
                // Regular list item - extract text
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
}

/**
 * Check if a node contains any heading elements (H1-H6) at any depth
 */
function containsHeading($node) {
    if ($node->nodeType === XML_ELEMENT_NODE) {
        $nodeName = strtolower($node->nodeName);
        if (in_array($nodeName, ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])) {
            return true;
        }
    }

    // Recursively check children
    if ($node->hasChildNodes()) {
        foreach ($node->childNodes as $child) {
            if (containsHeading($child)) {
                return true;
            }
        }
    }

    return false;
}

/**
 * Process table elements for DOCX
 */
function processTableForDocx($section, $tableNode) {
    // Count columns from first row
    $firstRow = null;
    $columnCount = 0;

    foreach ($tableNode->childNodes as $child) {
        if (strtolower($child->nodeName) === 'tbody' || strtolower($child->nodeName) === 'thead') {
            foreach ($child->childNodes as $row) {
                if (strtolower($row->nodeName) === 'tr') {
                    $firstRow = $row;
                    break 2;
                }
            }
        } elseif (strtolower($child->nodeName) === 'tr') {
            $firstRow = $child;
            break;
        }
    }

    if (!$firstRow) return;

    // Count columns
    foreach ($firstRow->childNodes as $cell) {
        $cellName = strtolower($cell->nodeName);
        if ($cellName === 'td' || $cellName === 'th') {
            $columnCount++;
        }
    }

    if ($columnCount === 0) return;

    // Create table
    $table = $section->addTable([
        'borderSize' => 6,
        'borderColor' => '999999',
        'width' => 100 * 50, // 100% width
        'unit' => \PhpOffice\PhpWord\SimpleType\TblWidth::PERCENT
    ]);

    // Process table rows
    $isFirstRow = true;
    foreach ($tableNode->childNodes as $section) {
        $sectionName = strtolower($section->nodeName);

        if ($sectionName === 'thead' || $sectionName === 'tbody' || $sectionName === 'tfoot') {
            foreach ($section->childNodes as $row) {
                if (strtolower($row->nodeName) === 'tr') {
                    processTableRow($table, $row, $isFirstRow);
                    $isFirstRow = false;
                }
            }
        } elseif ($sectionName === 'tr') {
            processTableRow($table, $section, $isFirstRow);
            $isFirstRow = false;
        }
    }
}

/**
 * Process a table row for DOCX
 */
function processTableRow($table, $rowNode, $isHeader = false) {
    $table->addRow();

    foreach ($rowNode->childNodes as $cell) {
        $cellName = strtolower($cell->nodeName);

        if ($cellName === 'td' || $cellName === 'th') {
            $isHeaderCell = ($cellName === 'th' || $isHeader);
            $cellText = getTextContent($cell);

            $cellStyle = [
                'valign' => 'center',
                'bgColor' => $isHeaderCell ? 'E8E8E8' : null
            ];

            $textStyle = [
                'size' => 10,
                'name' => 'Arial',
                'bold' => $isHeaderCell
            ];

            $paragraphStyle = [
                'spaceAfter' => 0,
                'spaceBefore' => 0
            ];

            $cell = $table->addCell(null, $cellStyle);
            $cell->addText($cellText, $textStyle, $paragraphStyle);
        }
    }
}

/**
 * Get all text content from a DOM node
 * Properly handles text extraction while preserving spaces between elements
 */
function getTextContent($node) {
    $text = '';
    foreach ($node->childNodes as $child) {
        if ($child->nodeType === XML_TEXT_NODE) {
            $text .= $child->nodeValue;
        } elseif ($child->hasChildNodes()) {
            $childText = getTextContent($child);
            // Add space between inline elements if needed
            if (!empty($childText) && !empty($text) && !preg_match('/\s$/', $text)) {
                $text .= ' ';
            }
            $text .= $childText;
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
        $textRun = $section->addTextRun(['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT, 'spaceAfter' => 240]);
        $textRun->addText(
            'Meta Title: ',
            ['bold' => true, 'size' => 11, 'name' => 'Arial', 'color' => '666666']
        );
        $textRun->addText(
            htmlspecialchars_decode($content['metaTitle'], ENT_QUOTES),
            ['bold' => true, 'size' => 18, 'name' => 'Arial']
        );
    } else {
        debugLog("  No meta title found");
    }

    // Add Meta Description if available
    if (!empty($content['metaDescription'])) {
        debugLog("  Adding meta description: " . substr($content['metaDescription'], 0, 50));
        $textRun = $section->addTextRun(['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT, 'spaceAfter' => 360]);
        $textRun->addText(
            'Meta Description: ',
            ['bold' => true, 'size' => 11, 'name' => 'Arial', 'color' => '666666']
        );
        $textRun->addText(
            htmlspecialchars_decode($content['metaDescription'], ENT_QUOTES),
            ['italic' => true, 'size' => 11, 'name' => 'Arial', 'color' => '666666']
        );
    } else {
        debugLog("  No meta description found");
    }

    // Add main content
    if (!empty($content['html'])) {
        debugLog("  Adding content to DOCX. HTML length: " . strlen($content['html']) . " bytes");

        // Count H3s before cleaning
        $h3CountBefore = substr_count($content['html'], '<h3');
        debugLog("  H3 tags found before cleaning: " . $h3CountBefore);

        // Clean HTML for better processing
        $cleanHtml = cleanHtmlForDocx($content['html']);
        debugLog("  After cleaning: " . strlen($cleanHtml) . " bytes");

        // Count H3s after cleaning
        $h3CountAfter = substr_count($cleanHtml, '<h3');
        debugLog("  H3 tags found after cleaning: " . $h3CountAfter);

        // Save cleaned HTML for debugging
        @file_put_contents(__DIR__ . '/output/cleaned_html_debug.html', $cleanHtml);
        debugLog("  Cleaned HTML saved to output/cleaned_html_debug.html for inspection");

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

    // Remove empty tags that can cause issues, but preserve heading tags
    // First, temporarily mark headings to protect them
    $html = preg_replace('/<(h[1-6])\b([^>]*)>\s*<\/\1>/', '<$1$2>__PRESERVE__</$1>', $html);
    // Remove other empty tags
    $html = preg_replace('/<(\w+)[^>]*>\s*<\/\1>/', '', $html);
    // Restore preserved headings (they'll be extracted later even if empty)
    $html = str_replace('__PRESERVE__', '', $html);

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
