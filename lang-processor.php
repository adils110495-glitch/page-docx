<?php
ob_start();
session_start();
require_once 'vendor/autoload.php';

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;

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

function langRemoveSkipSelectors($html, $skipSelectors) {
    if (empty($skipSelectors)) return $html;
    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
    libxml_clear_errors();
    $xpath = new DOMXPath($dom);
    $selectors = array_map('trim', explode(',', $skipSelectors));
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
    if ($selector && !empty(trim($selector))) {
        $selector = trim($selector);
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
    $text = preg_replace('/\s+/', ' ', $text);
    $text = str_replace(['<', '>'], ['', ''], $text);
    $text = str_replace('&', 'and', $text);
    $text = preg_replace('/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/', '', $text);
    return trim($text);
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

function langProcessInlineContent($textRun, $node, $fontStyle = []) {
    $blockElements = ['div','p','ul','ol','li','table','tr','td','th','h1','h2','h3','h4','h5','h6','blockquote','section','article','header','footer','nav','aside'];
    foreach ($node->childNodes as $child) {
        if ($child->nodeType === XML_TEXT_NODE) {
            $text = langSanitizeText($child->nodeValue);
            if (!empty($text)) $textRun->addText($text, $fontStyle);
        } elseif ($child->nodeType === XML_ELEMENT_NODE) {
            $nodeName = strtolower($child->nodeName);
            if (in_array($nodeName, $blockElements)) continue;
            switch ($nodeName) {
                case 'br': $textRun->addTextBreak(); break;
                case 'strong': case 'b':
                    langProcessInlineContent($textRun, $child, array_merge($fontStyle, ['bold' => true])); break;
                case 'em': case 'i':
                    langProcessInlineContent($textRun, $child, array_merge($fontStyle, ['italic' => true])); break;
                case 'u':
                    langProcessInlineContent($textRun, $child, array_merge($fontStyle, ['underline' => 'single'])); break;
                default:
                    langProcessInlineContent($textRun, $child, $fontStyle); break;
            }
        }
    }
}

function langAddElementContent($section, $node, $fontStyle = [], $paragraphStyle = []) {
    try {
        if (langContainsBrTag($node)) {
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

function langProcessListForDocx($section, $listNode, $listType) {
    foreach ($listNode->childNodes as $child) {
        if (strtolower($child->nodeName) === 'li') {
            if (langContainsHeading($child) || langContainsTitle1Class($child)) {
                langProcessNodeForDocx($section, $child, null, 0);
            } else {
                $text = langGetTextContent($child);
                if (!empty($text)) {
                    $text = langSanitizeText($text);
                    $section->addListItem(
                        $text,
                        0,
                        ['size' => 11, 'name' => 'Arial'],
                        $listType === 'ol' ? ['listType' => \PhpOffice\PhpWord\Style\ListItem::TYPE_NUMBER] : null,
                        ['spaceAfter' => 120]
                    );
                }
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
            if (langContainsBrTag($cellNode)) {
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
                $text = langGetTextContent($child);
                if (!empty($text)) {
                    $section->addText(langSanitizeText($text), ['bold' => true, 'size' => 14, 'name' => 'Arial'], ['spaceAfter' => 240]);
                    $section->addTextBreak();
                }
                continue;
            }

            switch ($nodeName) {
                case 'h1': case 'h2': case 'h3': case 'h4': case 'h5': case 'h6':
                    $sizes = ['h1'=>18,'h2'=>16,'h3'=>14,'h4'=>13,'h5'=>12,'h6'=>11];
                    $headingClass = $child->hasAttribute('class') ? $child->getAttribute('class') : '';
                    if (strpos($headingClass, 'title1') !== false) {
                        $text = langGetTextContent($child);
                        if (!empty($text)) {
                            $section->addText(langSanitizeText($text), ['bold' => true, 'size' => 14, 'name' => 'Arial'], ['spaceAfter' => 240]);
                            $section->addTextBreak();
                        }
                        break;
                    }
                    $size = $sizes[$nodeName];
                    if (langContainsBrTag($child)) {
                        langAddElementContent($section, $child, ['bold' => true, 'size' => $size, 'name' => 'Arial'], ['spaceAfter' => 240]);
                    } else {
                        $text = langGetTextContent($child);
                        if (!empty($text)) {
                            $section->addText(langSanitizeText($text), ['bold' => true, 'size' => $size, 'name' => 'Arial'], ['spaceAfter' => 240]);
                        }
                    }
                    break;

                case 'p':
                    if (langContainsBrTag($child)) {
                        langAddElementContent($section, $child, ['size' => 11, 'name' => 'Arial'], ['spaceAfter' => 200]);
                    } else {
                        $text = langGetTextContent($child);
                        if (!empty($text)) {
                            $section->addText(langSanitizeText($text), ['size' => 11, 'name' => 'Arial'], ['spaceAfter' => 200]);
                        }
                    }
                    break;

                case 'strong': case 'b':
                    $text = langGetTextContent($child);
                    if (!empty($text) && $textRun) $textRun->addText(langSanitizeText($text), ['bold' => true]);
                    break;

                case 'em': case 'i':
                    $text = langGetTextContent($child);
                    if (!empty($text) && $textRun) $textRun->addText(langSanitizeText($text), ['italic' => true]);
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
                        $text = langGetTextContent($child);
                        if (!empty($text)) {
                            $section->addText(langSanitizeText($text), ['bold' => true, 'size' => 14, 'name' => 'Arial'], ['spaceAfter' => 240]);
                            $section->addTextBreak();
                        }
                        break;
                    }
                    $headingClasses = [
                        'title2' => ['size' => 14, 'bold' => true],
                        'title3' => ['size' => 13, 'bold' => true],
                        'your-rights-faq__question' => ['size' => 13, 'bold' => true],
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
                        $text = langGetTextContent($child);
                        if (!empty($text)) {
                            $section->addText(langSanitizeText($text), array_merge(['name' => 'Arial'], $headingStyle), ['spaceAfter' => 240]);
                        }
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
    $path = isset($parsed['path']) ? $parsed['path'] : '';
    $path = trim($path, '/');

    if (!empty($path)) {
        $segments = array_values(array_filter(explode('/', $path), function($s) { return $s !== ''; }));

        if (!empty($segments)) {
            $first = strtolower($segments[0]);

            // 2-letter ISO 639-1
            if (preg_match('/^[a-z]{2}$/', $first) && in_array($first, $validCodes)) {
                return $first;
            }

            // Locale: en-us, pt-br, zh-cn, en-GB, etc.
            if (preg_match('/^([a-z]{2})[-_]([a-z]{2,4})$/i', $first, $matches)) {
                $base = strtolower($matches[1]);
                if (in_array($base, $validCodes)) {
                    return strtolower($first);
                }
            }
        }
    }

    return 'en'; // default to English when no language code detected in URL
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

$docName = 'lang-' . langBuildDocName($urls, $ISO_LANG_CODES);

langDebugLog("Detected {$langCount} language(s), {$totalUrls} total URLs");

try {
    $filepath = generateLangDocx($languageGroups, $docName, $project, $selector, $skipSels, $langNames);
    langDebugLog("Combined DOCX saved: {$filepath}");

    $langList = implode(', ', array_map('strtoupper', array_filter(array_keys($languageGroups), function($k) { return $k !== '__other__'; })));
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
