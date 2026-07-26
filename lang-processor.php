<?php
ob_start();
session_start();
require_once 'vendor/autoload.php';

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;

// Base URL used to make anchor hrefs absolute in generated DOCX files
if (!defined('LANG_DOCX_LINK_BASE_URL')) {
    define('LANG_DOCX_LINK_BASE_URL', 'http://localhost:8085/');
}

// ─── Utility Functions ────────────────────────────────────────────────────────

function langDebugLog($message) {
    $logFile = __DIR__ . '/output/lang-debug.log';
    $ts = date('Y-m-d H:i:s');
    @file_put_contents($logFile, "[{$ts}] {$message}\n", FILE_APPEND);
}

function langSanitizeProjectName($project) {
    $slug = preg_replace('/[^a-z0-9\-_]+/i', '-', $project);
    $slug = trim($slug, '-');
    $slug = strtolower($slug);
    if (strlen($slug) > 50) $slug = substr($slug, 0, 50);
    return $slug ?: 'default';
}

function langIsValidUrl($url) {
    $url = trim($url);
    if (empty($url)) return false;
    if (!preg_match('/^https?:\/\//i', $url)) return false;
    return filter_var($url, FILTER_VALIDATE_URL) !== false;
}

function langFetchHtml($url, $timeout = 30) {
    $context = stream_context_create([
        'http' => [
            'timeout' => $timeout,
            'user_agent' => 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'follow_location' => true,
        ],
        'ssl' => ['verify_peer' => false, 'verify_peer_name' => false]
    ]);
    $html = @file_get_contents($url, false, $context);
    return $html === false ? null : $html;
}

function langExtractMetaTitle($dom) {
    $titles = $dom->getElementsByTagName('title');
    if ($titles->length > 0) return trim($titles->item(0)->textContent);
    return null;
}

function langExtractMetaDescription($dom) {
    $xpath = new DOMXPath($dom);
    $metaTags = $xpath->query('//meta[@name="description"]');
    if ($metaTags->length > 0) return trim($metaTags->item(0)->getAttribute('content'));
    return null;
}

function langGetInnerHtml($node) {
    $innerHTML = '';
    foreach ($node->childNodes as $child) {
        $innerHTML .= $node->ownerDocument->saveHTML($child);
    }
    return $innerHTML;
}

/**
 * Normalize a single selector entry
 * Accepts a tag (article or <article>), a class (.my-class) or an ID (#my-id)
 */
function langNormalizeSelector($selector) {
    $selector = trim(str_replace(['<', '>', '/'], '', trim($selector)));
    if ($selector === '') return '';

    $prefix = '';
    if ($selector[0] === '#' || $selector[0] === '.') {
        $prefix = $selector[0];
        $selector = substr($selector, 1);
    }

    $selector = preg_replace('/[^a-zA-Z0-9_\-]/', '', $selector);
    return $selector === '' ? '' : $prefix . $selector;
}

function langRemoveSkipSelectors($html, $skipSelectors) {
    if (empty($skipSelectors)) return $html;
    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
    libxml_clear_errors();
    $xpath = new DOMXPath($dom);
    $selectors = array_filter(array_map('langNormalizeSelector', explode(',', $skipSelectors)));
    foreach ($selectors as $selector) {
        if (empty($selector)) continue;
        $nodesToRemove = [];
        if (strpos($selector, '#') === 0) {
            $id = substr($selector, 1);
            $nodes = $xpath->query("//*[@id='{$id}']");
            foreach ($nodes as $n) $nodesToRemove[] = $n;
        } elseif (strpos($selector, '.') === 0) {
            $class = substr($selector, 1);
            $nodes = $xpath->query("//*[contains(concat(' ', normalize-space(@class), ' '), ' {$class} ')]");
            foreach ($nodes as $n) $nodesToRemove[] = $n;
        } else {
            $nodes = $xpath->query("//{$selector}");
            foreach ($nodes as $n) $nodesToRemove[] = $n;
            $nodes = $xpath->query("//*[contains(concat(' ', normalize-space(@class), ' '), ' {$selector} ')]");
            foreach ($nodes as $n) $nodesToRemove[] = $n;
            $nodes = $xpath->query("//*[@id='{$selector}']");
            foreach ($nodes as $n) $nodesToRemove[] = $n;
        }
        foreach ($nodesToRemove as $node) {
            if ($node->parentNode) $node->parentNode->removeChild($node);
        }
    }
    $body = $dom->getElementsByTagName('body')->item(0);
    return $body ? langGetInnerHtml($body) : $html;
}

function langExtractContent($html, $selector = null, $skipSelectors = '') {
    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
    libxml_clear_errors();
    $xpath = new DOMXPath($dom);
    $metaTitle = langExtractMetaTitle($dom);
    $metaDescription = langExtractMetaDescription($dom);
    $contentHtml = '';
    if ($selector && langNormalizeSelector($selector) !== '') {
        // Accepts a tag (<article> or article), a class (.my-class) or an ID (#my-id)
        $selector = langNormalizeSelector($selector);
        if (strpos($selector, '#') === 0) {
            $id = substr($selector, 1);
            $nodes = $xpath->query("//*[@id='$id']");
        } elseif (strpos($selector, '.') === 0) {
            $class = substr($selector, 1);
            $nodes = $xpath->query("//*[contains(concat(' ', normalize-space(@class), ' '), ' $class ')]");
        } else {
            $tagNodes = $xpath->query("//$selector");
            if ($tagNodes !== false && $tagNodes->length > 0) {
                $nodes = $tagNodes;
            } else {
                $nodes = $xpath->query("//*[contains(concat(' ', normalize-space(@class), ' '), ' $selector ')]");
            }
        }
        if ($nodes !== false && $nodes->length > 0) {
            $contentHtml = langGetInnerHtml($nodes->item(0));
        } else {
            return ['success' => false, 'error' => 'Selector not found', 'metaTitle' => $metaTitle, 'metaDescription' => $metaDescription];
        }
    } else {
        $bodyNodes = $dom->getElementsByTagName('body');
        if ($bodyNodes->length > 0) {
            $contentHtml = langGetInnerHtml($bodyNodes->item(0));
        } else {
            return ['success' => false, 'error' => 'No body content found', 'metaTitle' => $metaTitle, 'metaDescription' => $metaDescription];
        }
    }
    if (!empty($skipSelectors)) {
        $contentHtml = langRemoveSkipSelectors($contentHtml, $skipSelectors);
    }
    return ['success' => true, 'html' => $contentHtml, 'metaTitle' => $metaTitle, 'metaDescription' => $metaDescription];
}

function langSanitizeText($text) {
    $text = html_entity_decode($text, ENT_QUOTES | ENT_HTML5, 'UTF-8');
    $text = preg_replace('/[\x{10000}-\x{10FFFF}]/u', '', $text);
    $text = preg_replace('/[\x{FFFE}\x{FFFF}]/u', '', $text);
    $text = preg_replace('/[\x{0080}-\x{009F}]/u', '', $text);
    $cyrillicMap = [
        'а'=>'a','с'=>'c','е'=>'e','о'=>'o','р'=>'p','х'=>'x','у'=>'y','і'=>'i','ј'=>'j','ѕ'=>'s',
        'А'=>'A','В'=>'B','С'=>'C','Е'=>'E','Н'=>'H','К'=>'K','М'=>'M','О'=>'O','Р'=>'P','Т'=>'T','Х'=>'X','І'=>'I',
    ];
    $text = strtr($text, $cyrillicMap);
    // Non-breaking / unicode spaces -> plain space (runs of them show as wide gaps)
    $text = preg_replace('/[\x{00A0}\x{1680}\x{2000}-\x{200A}\x{202F}\x{205F}\x{3000}]/u', ' ', $text);
    // Drop zero-width characters
    $text = preg_replace('/[\x{200B}-\x{200D}\x{2060}\x{FEFF}]/u', '', $text);
    $text = preg_replace('/\s+/u', ' ', $text);
    $text = str_replace(['<', '>'], ['', ''], $text);
    $text = str_replace('&', 'and', $text);
    $text = preg_replace('/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/', '', $text);
    return trim($text);
}

/**
 * Sanitize an inline text node, preserving the single leading/trailing space
 * that separates it from neighbouring inline elements (e.g. links)
 */
/**
 * Check whether the last text added to a run already ends with a space
 * Prevents double spaces around inline elements such as links
 */
function langRunEndsWithSpace($textRun) {
    $elements = $textRun->getElements();
    if (empty($elements)) return false;

    $last = end($elements);
    if ($last instanceof \PhpOffice\PhpWord\Element\Text || $last instanceof \PhpOffice\PhpWord\Element\Link) {
        return preg_match('/\s$/u', $last->getText()) === 1;
    }
    return true;
}

function langSanitizeInlineText($text) {
    $leading  = preg_match('/^\s/u', $text) ? ' ' : '';
    $trailing = preg_match('/\s$/u', $text) ? ' ' : '';
    $clean    = langSanitizeText($text);
    if ($clean === '') return '';
    return $leading . $clean . $trailing;
}

function langGetTextContent($node) {
    $text = '';
    foreach ($node->childNodes as $child) {
        if ($child->nodeType === XML_TEXT_NODE) {
            $text .= $child->nodeValue;
        } elseif ($child->nodeType === XML_ELEMENT_NODE) {
            $nodeName = strtolower($child->nodeName);
            if ($nodeName === 'br') {
                $text .= ' ';
            } elseif ($child->hasChildNodes()) {
                $childText = langGetTextContent($child);
                if (!empty($childText) && !empty($text) && !preg_match('/\s$/', $text)) {
                    $text .= ' ';
                }
                $text .= $childText;
            }
        }
    }
    return trim($text);
}

/**
 * Resolve an anchor href into an absolute URL for DOCX hyperlinks
 * Relative hrefs are resolved against http://localhost:8085/
 * Returns null for links that should not become hyperlinks
 */
function langResolveLinkUrl($href) {
    $href = trim(html_entity_decode((string) $href, ENT_QUOTES | ENT_HTML5, 'UTF-8'));
    $href = preg_replace('/[\x00-\x20\x7F]/', '', $href);

    if ($href === '' || $href === '#') return null;
    if (preg_match('/^(javascript|data|vbscript):/i', $href)) return null;
    if (preg_match('/^[a-z][a-z0-9+.\-]*:/i', $href)) return $href;
    if (strpos($href, '//') === 0) return 'http:' . $href;

    return rtrim(LANG_DOCX_LINK_BASE_URL, '/') . '/' . ltrim($href, '/');
}

function langHasBlockChild($node) {
    $blockElements = ['div','p','ul','ol','li','table','tr','td','th','h1','h2','h3','h4','h5','h6','blockquote','section','article','header','footer','nav','aside'];
    foreach ($node->childNodes as $child) {
        if ($child->nodeType === XML_ELEMENT_NODE && in_array(strtolower($child->nodeName), $blockElements)) return true;
    }
    return false;
}

function langContainsAnchor($node) {
    if ($node->nodeType === XML_ELEMENT_NODE && strtolower($node->nodeName) === 'a') return true;
    if ($node->hasChildNodes()) {
        foreach ($node->childNodes as $child) {
            if (langContainsAnchor($child)) return true;
        }
    }
    return false;
}

function langContainsInlineFormatting($node) {
    if ($node->nodeType === XML_ELEMENT_NODE && in_array(strtolower($node->nodeName), ['b','strong','em','i','u'])) return true;
    if ($node->hasChildNodes()) {
        foreach ($node->childNodes as $child) {
            if (langContainsInlineFormatting($child)) return true;
        }
    }
    return false;
}

function langNeedsInlineRun($node) {
    if (langContainsBrTag($node)) return true;
    // Inline processing drops block-level children, so only use it when there are none
    if (langHasBlockChild($node)) return false;

    return langContainsAnchor($node) || langContainsInlineFormatting($node);
}

function langAddLinkToTextRun($textRun, $node, $fontStyle = []) {
    $text = langSanitizeText(langGetTextContent($node));
    if ($text === '') return;

    $url = $node->hasAttribute('href') ? langResolveLinkUrl($node->getAttribute('href')) : null;
    if ($url === null) {
        $textRun->addText($text, $fontStyle);
        return;
    }

    $linkStyle = array_merge($fontStyle, ['color' => '0563C1', 'underline' => 'single']);
    try {
        $textRun->addLink($url, $text, $linkStyle);
    } catch (Exception $e) {
        $textRun->addText($text, $fontStyle);
    }
}

function langContainsBrTag($node) {
    if ($node->nodeType === XML_ELEMENT_NODE && strtolower($node->nodeName) === 'br') return true;
    if ($node->hasChildNodes()) {
        foreach ($node->childNodes as $child) {
            if (langContainsBrTag($child)) return true;
        }
    }
    return false;
}

function langContainsHeading($node) {
    if ($node->nodeType === XML_ELEMENT_NODE && in_array(strtolower($node->nodeName), ['h1','h2','h3','h4','h5','h6'])) return true;
    if ($node->hasChildNodes()) {
        foreach ($node->childNodes as $child) {
            if (langContainsHeading($child)) return true;
        }
    }
    return false;
}

function langContainsTitle1Class($node) {
    if ($node->nodeType === XML_ELEMENT_NODE) {
        $class = $node->hasAttribute('class') ? $node->getAttribute('class') : '';
        if (strpos($class, 'title1') !== false) return true;
    }
    if ($node->hasChildNodes()) {
        foreach ($node->childNodes as $child) {
            if (langContainsTitle1Class($child)) return true;
        }
    }
    return false;
}

function langProcessInlineContent($textRun, $node, $fontStyle = [], $suppressBold = false) {
    $blockElements = ['div','p','ul','ol','li','table','tr','td','th','h1','h2','h3','h4','h5','h6','blockquote','section','article','header','footer','nav','aside'];
    foreach ($node->childNodes as $child) {
        if ($child->nodeType === XML_TEXT_NODE) {
            $raw = $child->nodeValue;
            if (trim($raw) === '') {
                // Whitespace-only node between inline elements - keep one space
                if ($raw !== '' && $textRun->countElements() > 0 && !langRunEndsWithSpace($textRun)) $textRun->addText(' ', $fontStyle);
            } else {
                $text = langSanitizeInlineText($raw);
                if ($textRun->countElements() === 0 || langRunEndsWithSpace($textRun)) $text = ltrim($text);
                if ($text !== '') $textRun->addText($text, $fontStyle);
            }
        } elseif ($child->nodeType === XML_ELEMENT_NODE) {
            $nodeName = strtolower($child->nodeName);
            if (in_array($nodeName, $blockElements)) continue;
            switch ($nodeName) {
                case 'br': $textRun->addTextBreak(); break;
                case 'strong': case 'b':
                    // Bold, unless we are inside a heading (headings are never bold)
                    $boldStyle = $suppressBold ? $fontStyle : array_merge($fontStyle, ['bold' => true]);
                    langProcessInlineContent($textRun, $child, $boldStyle, $suppressBold); break;
                case 'em': case 'i':
                    langProcessInlineContent($textRun, $child, array_merge($fontStyle, ['italic' => true]), $suppressBold); break;
                case 'u':
                    langProcessInlineContent($textRun, $child, array_merge($fontStyle, ['underline' => 'single']), $suppressBold); break;
                case 'a':
                    langAddLinkToTextRun($textRun, $child, $fontStyle); break;
                default:
                    langProcessInlineContent($textRun, $child, $fontStyle, $suppressBold); break;
            }
        }
    }
}

/**
 * Add a heading as a real Word heading. Page h1-h6 map to Heading 2-6,
 * because Heading 1 is reserved for the language tab label.
 */
function langAddHeadingContent($section, $node, $htmlLevel) {
    $level = max(2, min(6, (int) $htmlLevel + 1));

    if (langNeedsInlineRun($node)) {
        // Bold suppressed so <strong>/<b> inside a heading stays unbolded
        $textRun = $section->addTextRun('Heading' . $level);
        langProcessInlineContent($textRun, $node, [], true);
        return $textRun->countElements() > 0;
    }

    $text = langSanitizeText(langGetTextContent($node));
    if ($text === '') return false;

    $section->addTitle($text, $level);
    return true;
}

function langAddElementContent($section, $node, $fontStyle = [], $paragraphStyle = []) {
    try {
        if (langNeedsInlineRun($node)) {
            $textRun = $section->addTextRun($paragraphStyle);
            langProcessInlineContent($textRun, $node, $fontStyle);
        } else {
            $text = langGetTextContent($node);
            if (!empty($text)) $section->addText(langSanitizeText($text), $fontStyle, $paragraphStyle);
        }
    } catch (Exception $e) {
        $text = langGetTextContent($node);
        if (!empty($text)) $section->addText(langSanitizeText($text), $fontStyle, $paragraphStyle);
    }
}

function langGetChildLists($node) {
    $lists = [];
    foreach ($node->childNodes as $child) {
        if ($child->nodeType === XML_ELEMENT_NODE && in_array(strtolower($child->nodeName), ['ul', 'ol'])) {
            $lists[] = $child;
        }
    }
    return $lists;
}

function langProcessListForDocx($section, $listNode, $listType, $depth = 0) {
    $listStyle = $listType === 'ol' ? ['listType' => \PhpOffice\PhpWord\Style\ListItem::TYPE_NUMBER] : null;

    foreach ($listNode->childNodes as $child) {
        if (strtolower($child->nodeName) === 'li') {
            // Nested lists are rendered separately, one level deeper
            $nestedLists = langGetChildLists($child);
            $item = $child;

            if (!empty($nestedLists)) {
                $item = $child->cloneNode(true);
                foreach (langGetChildLists($item) as $nested) {
                    $item->removeChild($nested);
                }
            }

            if (langContainsHeading($item) || langContainsTitle1Class($item)) {
                langProcessNodeForDocx($section, $item, null, 0);
            } elseif (langNeedsInlineRun($item)) {
                // Preserve bold/italic/links inside the bullet
                $listItemRun = $section->addListItemRun($depth, $listStyle, ['spaceAfter' => 120]);
                langProcessInlineContent($listItemRun, $item, ['size' => 11, 'name' => 'Arial']);
            } else {
                $text = langGetTextContent($item);
                if (!empty($text)) {
                    $text = langSanitizeText($text);
                    $section->addListItem(
                        $text,
                        $depth,
                        ['size' => 11, 'name' => 'Arial'],
                        $listStyle,
                        ['spaceAfter' => 120]
                    );
                }
            }

            foreach ($nestedLists as $nested) {
                langProcessListForDocx($section, $nested, strtolower($nested->nodeName), $depth + 1);
            }
        }
    }
}

function langProcessTableRow($table, $rowNode, $isHeader = false) {
    $table->addRow();
    foreach ($rowNode->childNodes as $cellNode) {
        $cellName = strtolower($cellNode->nodeName);
        if ($cellName === 'td' || $cellName === 'th') {
            $isHeaderCell = ($cellName === 'th' || $isHeader);
            $cellStyle = ['valign' => 'center', 'bgColor' => $isHeaderCell ? 'E8E8E8' : null];
            $textStyle = ['size' => 10, 'name' => 'Arial', 'bold' => $isHeaderCell];
            $paragraphStyle = ['spaceAfter' => 0, 'spaceBefore' => 0];
            $cell = $table->addCell(null, $cellStyle);
            if (langNeedsInlineRun($cellNode)) {
                $textRun = $cell->addTextRun($paragraphStyle);
                langProcessInlineContent($textRun, $cellNode, $textStyle);
            } else {
                $cell->addText(langSanitizeText(langGetTextContent($cellNode)), $textStyle, $paragraphStyle);
            }
        }
    }
}

function langProcessTableForDocx($section, $tableNode) {
    $firstRow = null;
    foreach ($tableNode->childNodes as $child) {
        $cn = strtolower($child->nodeName);
        if ($cn === 'tbody' || $cn === 'thead') {
            foreach ($child->childNodes as $row) {
                if (strtolower($row->nodeName) === 'tr') { $firstRow = $row; break 2; }
            }
        } elseif ($cn === 'tr') { $firstRow = $child; break; }
    }
    if (!$firstRow) return;
    $columnCount = 0;
    foreach ($firstRow->childNodes as $cell) {
        $cellName = strtolower($cell->nodeName);
        if ($cellName === 'td' || $cellName === 'th') $columnCount++;
    }
    if ($columnCount === 0) return;
    $table = $section->addTable([
        'borderSize' => 6,
        'borderColor' => '999999',
        'width' => 100 * 50,
        'unit' => \PhpOffice\PhpWord\SimpleType\TblWidth::PERCENT
    ]);
    $isFirstRow = true;
    foreach ($tableNode->childNodes as $sec) {
        $sn = strtolower($sec->nodeName);
        if ($sn === 'thead' || $sn === 'tbody' || $sn === 'tfoot') {
            foreach ($sec->childNodes as $row) {
                if (strtolower($row->nodeName) === 'tr') {
                    langProcessTableRow($table, $row, $isFirstRow);
                    $isFirstRow = false;
                }
            }
        } elseif ($sn === 'tr') {
            langProcessTableRow($table, $sec, $isFirstRow);
            $isFirstRow = false;
        }
    }
}

function langProcessNodeForDocx($section, $node, $textRun = null, $depth = 0) {
    foreach ($node->childNodes as $child) {
        $nodeName = strtolower($child->nodeName);
        $nodeValue = trim($child->nodeValue);

        if ($child->nodeType === XML_TEXT_NODE) {
            $trimmedValue = trim($nodeValue);
            if (!empty($trimmedValue)) {
                if ($textRun) {
                    $textRun->addText(langSanitizeText($nodeValue));
                } elseif (strlen($trimmedValue) > 2) {
                    $section->addText(langSanitizeText($trimmedValue), ['size' => 11, 'name' => 'Arial']);
                }
            }
            continue;
        }

        if ($child->nodeType === XML_ELEMENT_NODE) {
            $elementClass = $child->hasAttribute('class') ? $child->getAttribute('class') : '';

            if (strpos($elementClass, 'title1') !== false) {
                if (langAddHeadingContent($section, $child, 3)) {
                    $section->addTextBreak();
                }
                continue;
            }

            switch ($nodeName) {
                case 'h1': case 'h2': case 'h3': case 'h4': case 'h5': case 'h6':
                    $headingClass = $child->hasAttribute('class') ? $child->getAttribute('class') : '';
                    if (strpos($headingClass, 'title1') !== false) {
                        if (langAddHeadingContent($section, $child, 3)) {
                            $section->addTextBreak();
                        }
                        break;
                    }
                    langAddHeadingContent($section, $child, (int) substr($nodeName, 1));
                    break;

                case 'p':
                    if (langNeedsInlineRun($child)) {
                        langAddElementContent($section, $child, ['size' => 11, 'name' => 'Arial'], ['spaceAfter' => 200]);
                    } else {
                        $text = langGetTextContent($child);
                        if (!empty($text)) {
                            $section->addText(langSanitizeText($text), ['size' => 11, 'name' => 'Arial'], ['spaceAfter' => 200]);
                        }
                    }
                    break;

                case 'strong': case 'b':
                    // Bold content must never be dropped - start a paragraph if needed
                    if ($textRun) {
                        langProcessInlineContent($textRun, $child, ['bold' => true]);
                    } elseif (langGetTextContent($child) !== '') {
                        $boldRun = $section->addTextRun(['spaceAfter' => 200]);
                        langProcessInlineContent($boldRun, $child, ['size' => 11, 'name' => 'Arial', 'bold' => true]);
                    }
                    break;

                case 'em': case 'i':
                    if ($textRun) {
                        langProcessInlineContent($textRun, $child, ['italic' => true]);
                    } elseif (langGetTextContent($child) !== '') {
                        $italicRun = $section->addTextRun(['spaceAfter' => 200]);
                        langProcessInlineContent($italicRun, $child, ['size' => 11, 'name' => 'Arial', 'italic' => true]);
                    }
                    break;

                case 'a':
                    if ($textRun) {
                        langAddLinkToTextRun($textRun, $child);
                    } else {
                        $linkText = langGetTextContent($child);
                        if (!empty($linkText)) {
                            $linkRun = $section->addTextRun(['spaceAfter' => 200]);
                            langAddLinkToTextRun($linkRun, $child, ['size' => 11, 'name' => 'Arial']);
                        }
                    }
                    break;

                case 'ul': case 'ol':
                    langProcessListForDocx($section, $child, $nodeName);
                    break;

                case 'table':
                    langProcessTableForDocx($section, $child);
                    break;

                case 'br':
                    if ($textRun) $textRun->addTextBreak();
                    else $section->addTextBreak();
                    break;

                case 'div': case 'section': case 'article': case 'main':
                    $divClass = $child->hasAttribute('class') ? $child->getAttribute('class') : '';
                    if (strpos($divClass, 'title1') !== false) {
                        if (langAddHeadingContent($section, $child, 3)) {
                            $section->addTextBreak();
                        }
                        break;
                    }
                    $headingClasses = [
                        'title2' => 3,
                        'title3' => 4,
                        'your-rights-faq__question' => 4,
                        'your-rights-compensation__title' => 2,
                        'bordered-card__title' => 3,
                    ];
                    $headingLevel = null;
                    foreach ($headingClasses as $className => $classLevel) {
                        if (strpos($divClass, $className) !== false) {
                            $headingLevel = $classLevel;
                            break;
                        }
                    }
                    if ($headingLevel !== null) {
                        langAddHeadingContent($section, $child, $headingLevel);
                    } else {
                        langProcessNodeForDocx($section, $child, $textRun, $depth + 1);
                    }
                    break;

                default:
                    if ($child->hasChildNodes()) langProcessNodeForDocx($section, $child, $textRun, $depth + 1);
                    break;
            }
        }
    }
}

function langCleanHtml($html) {
    $html = preg_replace('/<script\b[^>]*>(.*?)<\/script>/is', '', $html);
    $html = preg_replace('/<style\b[^>]*>(.*?)<\/style>/is', '', $html);
    $html = preg_replace('/<!--(.|\s)*?-->/', '', $html);
    $html = preg_replace('/<svg\b[^>]*>(.*?)<\/svg>/is', '', $html);
    $html = preg_replace('/<noscript\b[^>]*>(.*?)<\/noscript>/is', '', $html);
    $html = preg_replace('/<iframe\b[^>]*>(.*?)<\/iframe>/is', '', $html);
    $html = preg_replace('/<(h[1-6])\b([^>]*)>\s*<\/\1>/', '<$1$2>__PRESERVE__</$1>', $html);
    $html = preg_replace('/<(\w+)[^>]*>\s*<\/\1>/', '', $html);
    $html = str_replace('__PRESERVE__', '', $html);
    return trim($html);
}

// ─── Language Code Detection ──────────────────────────────────────────────────

$ISO_LANG_CODES = [
    'ab','aa','af','ak','sq','am','ar','an','hy','as','av','ae','ay','az',
    'bm','ba','eu','be','bn','bh','bi','bs','br','bg','my','ca','ch','ce',
    'ny','zh','cv','kw','co','cr','hr','cs','da','dv','nl','dz','en','eo',
    'et','ee','fo','fj','fi','fr','ff','gl','ka','de','el','gn','gu','ht',
    'ha','he','hz','hi','ho','hu','ia','id','ie','ga','ig','ik','io','is',
    'it','iu','ja','jv','kl','kn','kr','ks','kk','km','ki','rw','ky','kv',
    'kg','ko','ku','kj','la','lb','lg','li','ln','lo','lt','lu','lv','gv',
    'mk','mg','ms','ml','mt','mi','mr','mh','mn','na','nv','nd','ne','ng',
    'nb','nn','no','ii','nr','oc','oj','cu','om','or','os','pa','pi','fa',
    'pl','ps','pt','qu','rm','rn','ro','ru','sa','sc','sd','se','sm','sg',
    'sr','gd','sn','si','sk','sl','so','st','es','su','sw','ss','sv','ta',
    'te','tg','th','ti','bo','tk','tl','tn','to','tr','ts','tt','tw','ty',
    'ug','uk','ur','uz','ve','vi','vo','wa','cy','wo','fy','xh','yi','yo',
    'za','zu'
];

// Convert URL to a readable slug for use as a fallback heading label
function langUrlToSlug($url) {
    $parsed = parse_url($url);
    $path   = isset($parsed['path']) ? trim($parsed['path'], '/') : '';
    if (empty($path)) {
        return isset($parsed['host']) ? $parsed['host'] : $url;
    }
    $segments = array_filter(explode('/', $path), function($s) { return $s !== ''; });
    $slug = implode(' › ', array_map(function($s) {
        return ucwords(str_replace(['-', '_'], ' ', rawurldecode($s)));
    }, $segments));
    return $slug ?: $url;
}

function detectLanguageCode($url, $validCodes) {
    $parsed = parse_url($url);
    $path   = isset($parsed['path']) ? $parsed['path'] : '';
    $path   = trim($path, '/');

    if (!empty($path)) {
        $segments = array_values(array_filter(explode('/', $path), function($s) { return $s !== ''; }));

        if (!empty($segments)) {
            $first = strtolower($segments[0]);

            // Any 2-letter code (language or region, e.g. gb, fr, de, us)
            if (preg_match('/^[a-z]{2}$/', $first)) {
                return $first;
            }

            // Locale code: en-us, pt-br, zh-cn, en-gb, etc.
            if (preg_match('/^[a-z]{2}[-_][a-z]{2,4}$/i', $first)) {
                return strtolower(str_replace('_', '-', $first));
            }
        }
    }

    return 'en';
}

// ─── Combined DOCX Generator ──────────────────────────────────────────────────

function generateLangDocx($languageGroups, $filename, $project, $selector, $skipSelectors, $langNames) {
    $phpWord = new PhpWord();

    // Define Heading styles so they appear in Word's Navigation Pane (left sidebar)
    // Heading 1 = Language tab (EN, ES, FR …)
    $phpWord->addTitleStyle(1, [
        'bold'      => true,
        'size'      => 22,
        'name'      => 'Arial',
        'color'     => '4472C4',
    ], [
        'spaceAfter'  => 120,
        'spaceBefore' => 240,
        'keepNext'    => true,
    ]);

    // Heading 2-6 = the page's own h1-h6 (shifted down one level, since
    // Heading 1 is used for the language tab label)
    foreach ([2 => 18, 3 => 16, 4 => 14, 5 => 13, 6 => 12] as $level => $size) {
        // Headings are sized, not bold - only <b>/<strong> inside them turns bold
        $phpWord->addTitleStyle(
            $level,
            ['bold' => false, 'size' => $size, 'name' => 'Arial'],
            ['spaceAfter' => 240, 'spaceBefore' => 120]
        );
    }


    // Sort language groups alphabetically
    ksort($languageGroups);

    $isFirstSection = true;

    foreach ($languageGroups as $langCode => $urls) {
        if (empty($urls)) continue;

        if ($isFirstSection) {
            $section = $phpWord->addSection();
            $isFirstSection = false;
        } else {
            $section->addPageBreak();
        }

        // ── Heading 1: Language tab label ──────────────────────────────────
        $displayCode = strtoupper($langCode);
        $langLabel   = isset($langNames[$langCode]) ? $langNames[$langCode] : '';
        $headerText  = $langLabel ? "{$displayCode} — {$langLabel}" : $displayCode;

        $section->addTitle($headerText, 1);

        foreach ($urls as $urlIndex => $url) {
            langDebugLog("Processing [{$langCode}] URL: {$url}");

            if (!langIsValidUrl($url)) {
                $section->addText('Skipped invalid URL: ' . langSanitizeText($url), ['italic' => true, 'size' => 10, 'name' => 'Arial', 'color' => 'cc0000'], ['spaceAfter' => 200]);
                continue;
            }

            // Fetch page
            $html = langFetchHtml($url);
            if ($html === null) {
                langDebugLog("  Fetch failed for: {$url}");
                continue; // skip silently
            }

            $extracted = langExtractContent($html, $selector, $skipSelectors);

            if (!$extracted['success']) {
                langDebugLog("  Extraction failed: " . $extracted['error']);
                continue; // skip silently — no error text in document
            }

            // Page content
            if (!empty($extracted['html'])) {
                $cleanHtml = langCleanHtml($extracted['html']);
                $dom = new DOMDocument();
                libxml_use_internal_errors(true);
                $dom->loadHTML(mb_convert_encoding($cleanHtml, 'HTML-ENTITIES', 'UTF-8'));
                libxml_clear_errors();
                $body = $dom->getElementsByTagName('body')->item(0);
                if ($body) {
                    langProcessNodeForDocx($section, $body);
                } else {
                    foreach (preg_split('/\n\s*\n/', trim(strip_tags($cleanHtml))) as $para) {
                        $para = trim($para);
                        if (!empty($para)) {
                            $section->addText(langSanitizeText($para), ['size' => 11, 'name' => 'Arial'], ['spaceAfter' => 200]);
                        }
                    }
                }
            }

            // Separator between pages within the same language section
            if ($urlIndex < count($urls) - 1) {
                $section->addTextBreak(1);
                $section->addText(
                    str_repeat('─ ', 35),
                    ['size' => 8, 'color' => 'dddddd', 'name' => 'Arial'],
                    ['spaceAfter' => 200]
                );
            }
        }
    }

    // Determine output directory
    $outputDir = __DIR__ . '/output';
    if (!empty($project)) {
        $projectSlug = langSanitizeProjectName($project);
        $outputDir .= '/' . $projectSlug;
    }
    if (!is_dir($outputDir)) {
        mkdir($outputDir, 0777, true);
        @chmod($outputDir, 0777);
    }

    $filepath = $outputDir . '/' . $filename . '.docx';
    $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
    $objWriter->save($filepath);

    return $filepath;
}

// ─── Main Processing ──────────────────────────────────────────────────────────

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    header('Location: lang-generator.php');
    exit;
}

global $ISO_LANG_CODES;

$urlsInput   = isset($_POST['urls']) ? $_POST['urls'] : '';
$selector    = isset($_POST['selector']) ? trim($_POST['selector']) : '';
$skipSels    = isset($_POST['skip_selectors']) ? trim($_POST['skip_selectors']) : '';
$project     = isset($_POST['project']) ? trim($_POST['project']) : '';

langDebugLog("=== Language Generator — new request ===");
langDebugLog("Project: " . ($project ?: 'none') . " | Selector: " . ($selector ?: 'none'));

// Parse URLs
$urls = array_values(array_filter(
    array_map('trim', explode("\n", $urlsInput)),
    function($u) { return !empty($u); }
));

if (empty($urls)) {
    $_SESSION['lang_status'] = ['type' => 'error', 'message' => 'No URLs provided.'];
    header('Location: lang-generator.php');
    exit;
}

// Cap at 200 URLs
if (count($urls) > 200) {
    $urls = array_slice($urls, 0, 200);
}

// Detect language code for each URL and group
$languageGroups = [];
foreach ($urls as $url) {
    $lang = detectLanguageCode($url, $ISO_LANG_CODES);
    $key  = $lang ?? '__other__';
    if (!isset($languageGroups[$key])) $languageGroups[$key] = [];
    $languageGroups[$key][] = $url;
}

$langNames = [
    'en'=>'English','es'=>'Spanish','fr'=>'French','de'=>'German','it'=>'Italian',
    'pt'=>'Portuguese','nl'=>'Dutch','pl'=>'Polish','ru'=>'Russian','ar'=>'Arabic',
    'zh'=>'Chinese','ja'=>'Japanese','ko'=>'Korean','sv'=>'Swedish','da'=>'Danish',
    'fi'=>'Finnish','nb'=>'Norwegian','no'=>'Norwegian','cs'=>'Czech','sk'=>'Slovak',
    'ro'=>'Romanian','hu'=>'Hungarian','bg'=>'Bulgarian','hr'=>'Croatian','sr'=>'Serbian',
    'uk'=>'Ukrainian','el'=>'Greek','tr'=>'Turkish','he'=>'Hebrew','fa'=>'Persian',
    'hi'=>'Hindi','bn'=>'Bengali','th'=>'Thai','vi'=>'Vietnamese','id'=>'Indonesian',
    'ms'=>'Malay','ca'=>'Catalan','eu'=>'Basque','gl'=>'Galician','af'=>'Afrikaans',
    'sq'=>'Albanian','hy'=>'Armenian','ka'=>'Georgian','lv'=>'Latvian','lt'=>'Lithuanian',
    'et'=>'Estonian','sl'=>'Slovenian','mk'=>'Macedonian','is'=>'Icelandic',
    'ga'=>'Irish','cy'=>'Welsh','mt'=>'Maltese','lb'=>'Luxembourgish',
];

$langCount = count(array_filter(array_keys($languageGroups), function($k) { return $k !== '__other__'; }));
$totalUrls = count($urls);

// ── Determine document filename from slug ─────────────────────────────────────
// Priority: first URL with NO language code → strip lang code from first URL
function langBuildDocName($urls, $validCodes) {
    $noLangUrl = null;
    $firstUrl  = null;

    foreach ($urls as $url) {
        $parsed   = parse_url($url);
        $path     = trim(isset($parsed['path']) ? $parsed['path'] : '', '/');
        $segments = array_values(array_filter(explode('/', $path), function($s) { return $s !== ''; }));
        if (empty($segments)) continue;

        if ($firstUrl === null) $firstUrl = $url;

        $first = strtolower($segments[0]);
        $isLang = (preg_match('/^[a-z]{2}$/', $first) && in_array($first, $validCodes))
               || preg_match('/^[a-z]{2}[-_][a-z]{2,4}$/i', $first);

        if (!$isLang) {
            $noLangUrl = $url;
            break;
        }
    }

    $slugSource = $noLangUrl ?? $firstUrl;
    if (!$slugSource) return 'lang-combined-' . date('Y-m-d_H-i-s');

    $parsed   = parse_url($slugSource);
    $path     = trim(isset($parsed['path']) ? $parsed['path'] : '', '/');
    $segments = array_values(array_filter(explode('/', $path), function($s) { return $s !== ''; }));

    // If first segment is a language code, remove it
    if (!empty($segments)) {
        $first = strtolower($segments[0]);
        $isLang = (preg_match('/^[a-z]{2}$/', $first) && in_array($first, $validCodes))
               || preg_match('/^[a-z]{2}[-_][a-z]{2,4}$/i', $first);
        if ($isLang) array_shift($segments);
    }

    $slug = implode('-', $segments);
    $slug = preg_replace('/[^a-z0-9\-]+/i', '-', rawurldecode($slug));
    $slug = strtolower(trim($slug, '-'));
    if (strlen($slug) > 80) $slug = substr($slug, 0, 80);

    return $slug ?: 'lang-combined-' . date('Y-m-d_H-i-s');
}

$docName     = 'lang-' . langBuildDocName($urls, $ISO_LANG_CODES);
$outputMode  = isset($_POST['output_mode']) && $_POST['output_mode'] === 'google_drive' ? 'google_drive' : 'docx';
$langList    = implode(', ', array_map('strtoupper', array_filter(array_keys($languageGroups), function($k) { return $k !== '__other__'; })));

langDebugLog("Detected {$langCount} language(s), {$totalUrls} total URLs | mode: {$outputMode}");

// ── Helper: read a key from root .env ─────────────────────────────────────
function langReadEnv(string $key): string {
    $file = __DIR__ . '/.env';
    if (!file_exists($file)) return '';
    foreach (file($file, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES) as $line) {
        if ($line === '' || $line[0] === '#') continue;
        [$k, $v] = array_pad(explode('=', $line, 2), 2, '');
        if (trim($k) === $key) return trim($v);
    }
    return '';
}

// ══════════════════════════════════════════════════════════════════════════
// PATH A — Save to Google Drive
// ══════════════════════════════════════════════════════════════════════════
if ($outputMode === 'google_drive') {
    try {
        if (!class_exists('App\GoogleAuthHelper')) {
            throw new RuntimeException('Google API client not installed. Run: composer install');
        }

        $googleAuth = new \App\GoogleAuthHelper();
        if (!$googleAuth->isAuthenticated()) {
            throw new RuntimeException('Google account not connected. Please authenticate first.');
        }

        $googleClient = $googleAuth->createClient();

        $rawTitle = isset($_POST['google_doc_title']) ? trim($_POST['google_doc_title']) : '';
        $docTitle = $rawTitle !== '' ? $rawTitle : ltrim($docName, 'lang-');
        $folderId = langReadEnv('GOOGLE_DRIVE_FOLDER_ID');

        $exporter = new \App\GoogleDocsExporter($googleClient);
        $result   = $exporter->export($languageGroups, $docTitle, $selector, $skipSels, $folderId);

        langDebugLog("Google Doc created: " . $result['url']);

        $tabCount = count($result['langs']);
        $msg = "Google Doc created with {$tabCount} language tab(s)" . ($langList ? " ({$langList})" : '') . " — {$totalUrls} URL(s) processed.";

        $_SESSION['lang_status'] = [
            'type'              => 'success',
            'message'           => $msg,
            'processed'         => $totalUrls,
            'total'             => $totalUrls,
            'google_doc_url'    => $result['url'],
            'google_folder_url' => $result['folderUrl'],
        ];

    } catch (Exception $e) {
        langDebugLog("Google Drive ERROR: " . $e->getMessage());
        $_SESSION['lang_status'] = [
            'type'             => 'error',
            'message'          => 'Google Drive export failed: ' . $e->getMessage(),
            'google_docs_error'=> $e->getMessage(),
        ];
    }

    header('Location: lang-generator.php');
    exit;
}

// ══════════════════════════════════════════════════════════════════════════
// PATH B — Save as DOCX (default)
// ══════════════════════════════════════════════════════════════════════════
try {
    $filepath = generateLangDocx($languageGroups, $docName, $project, $selector, $skipSels, $langNames);
    langDebugLog("Combined DOCX saved: {$filepath}");

    $msg = "Generated combined document for {$langCount} language(s)" . ($langList ? " ({$langList})" : '') . " — {$totalUrls} URL(s) processed.";

    $_SESSION['lang_status'] = [
        'type'      => 'success',
        'message'   => $msg,
        'processed' => $totalUrls,
        'total'     => $totalUrls,
    ];
} catch (Exception $e) {
    langDebugLog("ERROR: " . $e->getMessage());
    $_SESSION['lang_status'] = [
        'type'    => 'error',
        'message' => 'Failed to generate document: ' . $e->getMessage(),
    ];
}

header('Location: lang-generator.php');
exit;
