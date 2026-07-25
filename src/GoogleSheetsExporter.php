<?php

declare(strict_types=1);

namespace App;

use Google\Client;

/**
 * Exports multi-language content into a single Google Spreadsheet.
 *
 * One sheet tab per language is created at spreadsheet-creation time
 * (Google Sheets API supports defining all sheets in the initial POST),
 * so no separate "insertTab" calls are required.
 *
 * Content layout per sheet:
 *   Row 1 : URL (bold, blue)
 *   Row 2+: Extracted text rows (headings bold, paragraphs normal, lists bulleted)
 *   ─────── separator between URLs ───────────────────────────────────
 */
class GoogleSheetsExporter
{
    private const SHEETS_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';
    private const DRIVE_BASE  = 'https://www.googleapis.com/drive/v3/files';

    private \GuzzleHttp\ClientInterface $http;
    private array $log = [];

    private static array $LANG_NAMES = [
        'en' => 'English',    'es' => 'Spanish',     'fr' => 'French',
        'de' => 'German',     'it' => 'Italian',     'pt' => 'Portuguese',
        'nl' => 'Dutch',      'pl' => 'Polish',      'ru' => 'Russian',
        'ar' => 'Arabic',     'zh' => 'Chinese',     'ja' => 'Japanese',
        'ko' => 'Korean',     'sv' => 'Swedish',     'da' => 'Danish',
        'fi' => 'Finnish',    'nb' => 'Norwegian',   'no' => 'Norwegian',
        'cs' => 'Czech',      'sk' => 'Slovak',      'ro' => 'Romanian',
        'hu' => 'Hungarian',  'bg' => 'Bulgarian',   'hr' => 'Croatian',
        'sr' => 'Serbian',    'uk' => 'Ukrainian',   'el' => 'Greek',
        'tr' => 'Turkish',    'he' => 'Hebrew',      'fa' => 'Persian',
        'hi' => 'Hindi',      'bn' => 'Bengali',     'th' => 'Thai',
        'vi' => 'Vietnamese', 'id' => 'Indonesian',  'ms' => 'Malay',
        'ca' => 'Catalan',    'eu' => 'Basque',      'gl' => 'Galician',
        'af' => 'Afrikaans',  'sq' => 'Albanian',    'hy' => 'Armenian',
        'ka' => 'Georgian',   'lv' => 'Latvian',     'lt' => 'Lithuanian',
        'et' => 'Estonian',   'sl' => 'Slovenian',   'mk' => 'Macedonian',
        'is' => 'Icelandic',  'ga' => 'Irish',       'cy' => 'Welsh',
        'mt' => 'Maltese',    'lb' => 'Luxembourgish',
        'gb' => 'United Kingdom', 'us' => 'United States',
        'au' => 'Australia',  'ca' => 'Canada',
    ];

    public function __construct(Client $client)
    {
        $this->http = $client->authorize();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Public API
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Creates a Google Spreadsheet with one sheet tab per language.
     *
     * @param  array  $languageGroups ['en' => ['url1', ...], 'fr' => [...], ...]
     * @param  string $title          Spreadsheet title
     * @param  string $selector       CSS selector (empty = full body)
     * @param  string $skipSelectors  Comma-separated selectors to exclude
     * @param  string $folderId       Drive folder ID (empty = root My Drive)
     * @return array  ['spreadsheetId', 'url', 'folderUrl', 'langs', 'log']
     */
    public function export(
        array  $languageGroups,
        string $title,
        string $selector,
        string $skipSelectors,
        string $folderId = ''
    ): array {
        ksort($languageGroups);
        $langCodes = array_keys($languageGroups);

        // ── Create spreadsheet with all language tabs in one request ──────
        $sheetsDef = array_map(fn($code) => [
            'properties' => ['title' => strtoupper($code)],
        ], $langCodes);

        $this->info("Creating spreadsheet: {$title} (" . count($langCodes) . " tabs)");
        $spreadsheet = $this->api('POST', self::SHEETS_BASE, [
            'json' => [
                'properties' => ['title' => $title],
                'sheets'     => $sheetsDef,
            ],
        ]);

        $spreadsheetId = $spreadsheet['spreadsheetId'];
        $this->info("Spreadsheet ID: {$spreadsheetId}");

        // Build sheetId map: tab title → sheetId (needed for formatting)
        $sheetIdMap = [];
        foreach ($spreadsheet['sheets'] ?? [] as $sheet) {
            $sheetIdMap[$sheet['properties']['title']] = $sheet['properties']['sheetId'];
        }

        // ── Move to folder ────────────────────────────────────────────────
        if ($folderId !== '') {
            $this->moveToFolder($spreadsheetId, $folderId);
            $this->info("Moved to folder: {$folderId}");
        }

        // ── Populate each language tab ────────────────────────────────────
        foreach ($languageGroups as $langCode => $urls) {
            $tabTitle = strtoupper($langCode);
            $sheetId  = $sheetIdMap[$tabTitle] ?? null;
            $this->info("Populating tab [{$tabTitle}]");
            $this->populateSheet($spreadsheetId, $tabTitle, $sheetId, $langCode, $urls, $selector, $skipSelectors);
        }

        $folderUrl = $folderId !== ''
            ? "https://drive.google.com/drive/folders/{$folderId}"
            : 'https://drive.google.com/drive/my-drive';

        return [
            'spreadsheetId' => $spreadsheetId,
            'url'           => "https://docs.google.com/spreadsheets/d/{$spreadsheetId}/edit",
            'folderUrl'     => $folderUrl,
            'langs'         => $langCodes,
            'log'           => $this->log,
        ];
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Sheet population
    // ─────────────────────────────────────────────────────────────────────────

    private function populateSheet(
        string  $spreadsheetId,
        string  $tabTitle,
        ?int    $sheetId,
        string  $langCode,
        array   $urls,
        string  $selector,
        string  $skipSelectors
    ): void {
        $rows           = [];
        $boldRows       = [];   // 0-based row indices that should be bold
        $urlHeaderRows  = [];   // 0-based row indices that are URL headers
        $headingRows    = [];   // 0-based row indices that are H1–H6

        // ── Language header row ───────────────────────────────────────────
        $langName = self::$LANG_NAMES[$langCode] ?? strtoupper($langCode);
        $rows[]        = [strtoupper($langCode) . ' — ' . $langName];
        $boldRows[]    = 0;
        $headingRows[] = 0;

        foreach ($urls as $url) {
            $this->info("  [{$langCode}] {$url}");

            $html = $this->fetchHtml($url);
            if ($html === null) {
                $this->info("  fetch failed — skipped");
                continue;
            }

            $extracted = $this->extractContent($html, $selector, $skipSelectors);
            if (!$extracted['success']) {
                $this->info("  extract failed: " . $extracted['error']);
                continue;
            }

            // URL header row
            $urlRowIdx       = count($rows);
            $rows[]          = [$url];
            $urlHeaderRows[] = $urlRowIdx;
            $rows[]          = [''];

            // Content rows
            $contentRows = $this->htmlToRows($extracted['html'], $headingRowOffsets);
            foreach ($headingRowOffsets as $offset) {
                $headingRows[] = count($rows) + $offset;
            }
            $rows = array_merge($rows, $contentRows);

            // Separator
            $rows[] = [''];
            $rows[] = [str_repeat('─', 80)];
            $rows[] = [''];
        }

        if (count($rows) <= 1) return; // only language header, nothing to write

        // ── Write values ──────────────────────────────────────────────────
        $range = "'{$tabTitle}'!A1";
        $this->api(
            'PUT',
            self::SHEETS_BASE . "/{$spreadsheetId}/values/" . rawurlencode($range)
            . '?valueInputOption=RAW',
            [
                'json' => [
                    'range'          => $range,
                    'majorDimension' => 'ROWS',
                    'values'         => $rows,
                ],
            ]
        );

        // ── Apply formatting ──────────────────────────────────────────────
        if ($sheetId !== null) {
            $this->applyFormatting(
                $spreadsheetId,
                $sheetId,
                $headingRows,
                $urlHeaderRows,
                count($rows)
            );
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // HTML → rows
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Converts HTML to a 2D array of cell values.
     * Also populates $headingOffsets with row indices (relative to start of
     * $rows) that correspond to H1–H6 elements.
     */
    private function htmlToRows(string $html, array &$headingOffsets = []): array
    {
        $html = $this->cleanHtml($html);
        if (trim($html) === '') return [];

        $dom = new \DOMDocument();
        libxml_use_internal_errors(true);
        $dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
        libxml_clear_errors();

        $rows           = [];
        $headingOffsets = [];

        $body = $dom->getElementsByTagName('body')->item(0);
        if ($body) {
            $this->nodeToRows($body, $rows, $headingOffsets);
        }

        return $rows;
    }

    private function nodeToRows(\DOMNode $node, array &$rows, array &$headingOffsets): void
    {
        foreach ($node->childNodes as $child) {
            if ($child->nodeType === XML_TEXT_NODE) {
                $text = trim($child->nodeValue);
                if ($text !== '') $rows[] = [$this->sanitize($text)];
                continue;
            }
            if ($child->nodeType !== XML_ELEMENT_NODE) continue;

            $tag = strtolower($child->nodeName);

            if (in_array($tag, ['script', 'style', 'svg', 'noscript', 'iframe', 'nav', 'header', 'footer'], true)) {
                continue;
            }

            switch ($tag) {
                case 'h1': case 'h2': case 'h3':
                case 'h4': case 'h5': case 'h6':
                    $text = $this->sanitize($this->textOf($child));
                    if ($text !== '') {
                        $headingOffsets[] = count($rows);
                        $rows[] = [$text];
                    }
                    break;

                case 'p':
                    $text = $this->sanitize($this->textOf($child));
                    if ($text !== '') $rows[] = [$text];
                    break;

                case 'li':
                    $text = $this->sanitize($this->textOf($child));
                    if ($text !== '') $rows[] = ['• ' . $text];
                    break;

                case 'tr':
                    $cells = [];
                    foreach ($child->childNodes as $cell) {
                        $ct = strtolower($cell->nodeName);
                        if ($ct === 'td' || $ct === 'th') {
                            $cells[] = $this->sanitize($this->textOf($cell));
                        }
                    }
                    if (!empty(array_filter($cells))) $rows[] = $cells;
                    break;

                case 'br':
                    $rows[] = [''];
                    break;

                case 'table':
                    $this->nodeToRows($child, $rows, $headingOffsets);
                    $rows[] = [''];
                    break;

                default:
                    if ($child->hasChildNodes()) {
                        $this->nodeToRows($child, $rows, $headingOffsets);
                    }
                    break;
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Formatting
    // ─────────────────────────────────────────────────────────────────────────

    private function applyFormatting(
        string $spreadsheetId,
        int    $sheetId,
        array  $headingRows,
        array  $urlHeaderRows,
        int    $totalRows
    ): void {
        $requests = [];

        // Auto-resize column A to fit content
        $requests[] = [
            'autoResizeDimensions' => [
                'dimensions' => [
                    'sheetId'   => $sheetId,
                    'dimension' => 'COLUMNS',
                    'startIndex' => 0,
                    'endIndex'   => 1,
                ],
            ],
        ];

        // Bold + larger font for heading rows
        foreach ($headingRows as $rowIdx) {
            $requests[] = $this->boldRow($sheetId, $rowIdx, 12);
        }

        // Bold + blue for URL header rows
        foreach ($urlHeaderRows as $rowIdx) {
            $requests[] = [
                'repeatCell' => [
                    'range' => [
                        'sheetId'        => $sheetId,
                        'startRowIndex'  => $rowIdx,
                        'endRowIndex'    => $rowIdx + 1,
                        'startColumnIndex' => 0,
                        'endColumnIndex'   => 10,
                    ],
                    'cell' => [
                        'userEnteredFormat' => [
                            'textFormat' => [
                                'bold'            => true,
                                'foregroundColor' => ['red' => 0.1, 'green' => 0.46, 'blue' => 0.82],
                            ],
                        ],
                    ],
                    'fields' => 'userEnteredFormat.textFormat',
                ],
            ];
        }

        if (!empty($requests)) {
            $this->api(
                'POST',
                self::SHEETS_BASE . "/{$spreadsheetId}:batchUpdate",
                ['json' => ['requests' => $requests]]
            );
        }
    }

    private function boldRow(int $sheetId, int $rowIdx, int $fontSize = 11): array
    {
        return [
            'repeatCell' => [
                'range' => [
                    'sheetId'          => $sheetId,
                    'startRowIndex'    => $rowIdx,
                    'endRowIndex'      => $rowIdx + 1,
                    'startColumnIndex' => 0,
                    'endColumnIndex'   => 10,
                ],
                'cell' => [
                    'userEnteredFormat' => [
                        'textFormat' => ['bold' => true, 'fontSize' => $fontSize],
                    ],
                ],
                'fields' => 'userEnteredFormat.textFormat',
            ],
        ];
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Drive helpers
    // ─────────────────────────────────────────────────────────────────────────

    private function moveToFolder(string $fileId, string $folderId): void
    {
        $meta    = $this->api('GET', self::DRIVE_BASE . "/{$fileId}?fields=parents");
        $parents = implode(',', $meta['parents'] ?? []);

        $url = self::DRIVE_BASE . "/{$fileId}"
             . "?addParents={$folderId}"
             . ($parents !== '' ? "&removeParents={$parents}" : '')
             . '&fields=id,parents';

        $this->api('PATCH', $url);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // HTML fetching & extraction
    // ─────────────────────────────────────────────────────────────────────────

    private function fetchHtml(string $url, int $timeout = 30): ?string
    {
        $context = stream_context_create([
            'http' => [
                'timeout'         => $timeout,
                'user_agent'      => 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                'follow_location' => true,
            ],
            'ssl' => ['verify_peer' => false, 'verify_peer_name' => false],
        ]);
        $html = @file_get_contents($url, false, $context);
        return $html === false ? null : $html;
    }

    /**
     * Normalize a selector entry: a tag (article or <article>),
     * a class (.my-class) or an ID (#my-id).
     */
    private function normalizeSelector(string $selector): string
    {
        $selector = trim(str_replace(['<', '>', '/'], '', trim($selector)));
        if ($selector === '') return '';

        $prefix = '';
        if ($selector[0] === '#' || $selector[0] === '.') {
            $prefix   = $selector[0];
            $selector = substr($selector, 1);
        }

        $selector = preg_replace('/[^a-zA-Z0-9_\-]/', '', $selector);
        return $selector === '' ? '' : $prefix . $selector;
    }

    private function extractContent(string $html, string $selector, string $skipSelectors): array
    {
        $dom = new \DOMDocument();
        libxml_use_internal_errors(true);
        $dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
        libxml_clear_errors();
        $xpath = new \DOMXPath($dom);

        $sel = $this->normalizeSelector($selector);
        if ($sel !== '') {
            if ($sel[0] === '#') {
                $nodes = $xpath->query("//*[@id='" . substr($sel, 1) . "']");
            } elseif ($sel[0] === '.') {
                $cls   = substr($sel, 1);
                $nodes = $xpath->query("//*[contains(concat(' ',normalize-space(@class),' '),' {$cls} ')]");
            } else {
                $nodes = $xpath->query("//{$sel}");
                if (!$nodes || $nodes->length === 0) {
                    $nodes = $xpath->query("//*[contains(concat(' ',normalize-space(@class),' '),' {$sel} ')]");
                }
            }
            if (!$nodes || $nodes->length === 0) {
                return ['success' => false, 'error' => 'Selector not found'];
            }
            $contentHtml = $this->innerHtml($nodes->item(0));
        } else {
            $body = $dom->getElementsByTagName('body')->item(0);
            if (!$body) return ['success' => false, 'error' => 'No body element'];
            $contentHtml = $this->innerHtml($body);
        }

        if ($skipSelectors !== '') {
            $contentHtml = $this->stripSelectors($contentHtml, $skipSelectors);
        }

        return ['success' => true, 'html' => $contentHtml];
    }

    private function stripSelectors(string $html, string $skipSelectors): string
    {
        $dom = new \DOMDocument();
        libxml_use_internal_errors(true);
        $dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
        libxml_clear_errors();
        $xpath = new \DOMXPath($dom);

        foreach (explode(',', $skipSelectors) as $rawSel) {
            $sel = $this->normalizeSelector($rawSel);
            if ($sel === '') continue;
            $toRemove = [];
            if ($sel[0] === '#') {
                foreach ($xpath->query("//*[@id='" . substr($sel, 1) . "']") as $n) $toRemove[] = $n;
            } elseif ($sel[0] === '.') {
                $cls = substr($sel, 1);
                foreach ($xpath->query("//*[contains(concat(' ',normalize-space(@class),' '),' {$cls} ')]") as $n) $toRemove[] = $n;
            } else {
                foreach ($xpath->query("//{$sel}") as $n) $toRemove[] = $n;
            }
            foreach ($toRemove as $n) {
                if ($n->parentNode) $n->parentNode->removeChild($n);
            }
        }

        $body = $dom->getElementsByTagName('body')->item(0);
        return $body ? $this->innerHtml($body) : $html;
    }

    private function innerHtml(\DOMNode $node): string
    {
        $out = '';
        foreach ($node->childNodes as $child) {
            $out .= $node->ownerDocument->saveHTML($child);
        }
        return $out;
    }

    private function cleanHtml(string $html): string
    {
        $html = preg_replace('/<script\b[^>]*>.*?<\/script>/is', '', $html);
        $html = preg_replace('/<style\b[^>]*>.*?<\/style>/is', '', $html);
        $html = preg_replace('/<!--.*?-->/s', '', $html);
        $html = preg_replace('/<svg\b[^>]*>.*?<\/svg>/is', '', $html);
        $html = preg_replace('/<noscript\b[^>]*>.*?<\/noscript>/is', '', $html);
        $html = preg_replace('/<iframe\b[^>]*>.*?<\/iframe>/is', '', $html);
        return trim($html);
    }

    private function sanitize(string $text): string
    {
        $text = html_entity_decode($text, ENT_QUOTES | ENT_HTML5, 'UTF-8');
        $text = preg_replace('/[\x{10000}-\x{10FFFF}]/u', '', $text);
        $text = preg_replace('/[\x{0080}-\x{009F}]/u', '', $text);
        $text = preg_replace('/\s+/', ' ', $text);
        $text = preg_replace('/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/', '', $text);
        return trim($text);
    }

    private function textOf(\DOMNode $node): string
    {
        $text = '';
        foreach ($node->childNodes as $child) {
            if ($child->nodeType === XML_TEXT_NODE) {
                $text .= $child->nodeValue;
            } elseif ($child->nodeType === XML_ELEMENT_NODE && $child->hasChildNodes()) {
                $text .= ' ' . $this->textOf($child);
            }
        }
        return trim($text);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // HTTP helper
    // ─────────────────────────────────────────────────────────────────────────

    private function api(string $method, string $url, array $options = []): array
    {
        try {
            $response = $this->http->request($method, $url, array_merge(
                ['http_errors' => false],
                $options
            ));
            $body = (string)$response->getBody();
            $data = json_decode($body, true) ?? [];

            if ($response->getStatusCode() >= 400) {
                $msg = $data['error']['message'] ?? "HTTP {$response->getStatusCode()}";
                throw new \RuntimeException("Google Sheets API error: {$msg}");
            }
            return $data;
        } catch (\GuzzleHttp\Exception\GuzzleException $e) {
            throw new \RuntimeException('HTTP request failed: ' . $e->getMessage(), 0, $e);
        }
    }

    private function info(string $message): void
    {
        $ts          = date('Y-m-d H:i:s');
        $this->log[] = "[{$ts}] {$message}";
        @file_put_contents(
            dirname(__DIR__) . '/output/lang-debug.log',
            "[{$ts}] [GoogleSheets] {$message}\n",
            FILE_APPEND
        );
    }
}
