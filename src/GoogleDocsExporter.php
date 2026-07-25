<?php

declare(strict_types=1);

namespace App;

use Google\Client;

/**
 * Exports multi-language content scraped from URLs into a Google Document.
 *
 * Strategy:
 *   • Probes for Google Docs Tabs API support at runtime.
 *   • If tabs are available: creates one tab per language and inserts content there.
 *   • If tabs are unavailable: falls back to page-break-separated H1 sections in
 *     the default tab, keeping the code ready for tab support when it arrives.
 *
 * Content insertion uses sequential index tracking so heading/paragraph/bold/italic
 * styles are applied in the same batchUpdate call as the text insertions.
 */
class GoogleDocsExporter
{
    private const DOCS_BASE  = 'https://docs.googleapis.com/v1/documents';
    private const DRIVE_BASE = 'https://www.googleapis.com/drive/v3/files';

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
    ];

    public function __construct(Client $client)
    {
        $this->http = $client->authorize();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Public API
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Creates a single Google Document with one tab per language.
     *
     * Attempts to create real Document tabs via insertTab (batchUpdate with
     * ?includeTabsContent=true). If the API rejects it, falls back to
     * page-break-separated H1 sections inside the default tab.
     *
     * @param  array  $languageGroups ['en' => ['url1', ...], 'fr' => [...], ...]
     * @param  string $title          Document title
     * @param  string $selector       CSS selector (empty = full body)
     * @param  string $skipSelectors  Comma-separated selectors to exclude
     * @param  string $folderId       Drive folder ID (empty = root "My Drive")
     * @return array  ['docId', 'url', 'folderUrl', 'langs', 'log']
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

        $this->info("Creating document: {$title}");
        $docId = $this->createDocument($title);
        $this->info("Document ID: {$docId}");

        if ($folderId !== '') {
            $this->moveToFolder($docId, $folderId);
            $this->info("Moved to folder: {$folderId}");
        }

        // Try real tabs — fall back to sections if API rejects insertTab
        $tabsOk = $this->tryExportWithTabs($docId, $languageGroups, $selector, $skipSelectors, $langCodes);

        if (!$tabsOk) {
            $this->info("Tab creation not supported — using sections fallback");
            $this->exportWithSections($docId, $languageGroups, $selector, $skipSelectors);
        }

        $folderUrl = $folderId !== ''
            ? "https://drive.google.com/drive/folders/{$folderId}"
            : 'https://drive.google.com/drive/my-drive';

        return [
            'docId'     => $docId,
            'url'       => "https://docs.google.com/document/d/{$docId}/edit",
            'folderUrl' => $folderUrl,
            'langs'     => $langCodes,
            'log'       => $this->log,
        ];
    }

    /**
     * Attempts to create one real Document tab per language.
     * Returns true if all tabs were created and populated successfully.
     * Returns false if insertTab is rejected (caller falls back to sections).
     */
    private function tryExportWithTabs(
        string $docId,
        array  $languageGroups,
        string $selector,
        string $skipSelectors,
        array  $langCodes
    ): bool {
        $tabIds = [];

        foreach ($langCodes as $i => $code) {
            $langName = self::$LANG_NAMES[$code] ?? strtoupper($code);
            $tabTitle = strtoupper($code) . ' — ' . $langName;

            // Snapshot current tab IDs so we can identify the new one
            $beforeIds = $this->getCurrentTabIds($docId);

            try {
                // Pass includeTabsContent=true — required for tab-write operations
                $this->api(
                    'POST',
                    self::DOCS_BASE . "/{$docId}:batchUpdate?includeTabsContent=true",
                    [
                        'json' => [
                            'requests' => [[
                                'insertTab' => [
                                    'insertionIndex' => $i,
                                    'tab'            => ['tabProperties' => ['title' => $tabTitle]],
                                ],
                            ]],
                        ],
                    ]
                );
            } catch (\RuntimeException $e) {
                $this->info("insertTab rejected: " . $e->getMessage());
                return false;
            }

            // Find the newly created tab by diff
            $afterIds = $this->getCurrentTabIds($docId);
            $newId    = $this->diffTabIds($beforeIds, $afterIds);

            if ($newId === '') {
                $this->info("Could not resolve new tab ID for [{$code}]");
                return false;
            }

            $tabIds[$code] = $newId;
            $this->info("Tab created: {$tabTitle} → {$newId}");
        }

        // Populate each tab with its URL content
        foreach ($languageGroups as $code => $urls) {
            $tabId = $tabIds[$code] ?? null;
            if ($tabId) {
                $this->populateTab($docId, $tabId, $code, $urls, $selector, $skipSelectors);
            }
        }

        return true;
    }

    /** Returns all current tabId strings for the document. */
    private function getCurrentTabIds(string $docId): array
    {
        $doc  = $this->api('GET', self::DOCS_BASE . "/{$docId}?includeTabsContent=true");
        $ids  = [];
        foreach ($doc['tabs'] ?? [] as $tab) {
            $id = $tab['tabProperties']['tabId'] ?? '';
            if ($id !== '') $ids[] = $id;
        }
        return $ids;
    }

    /** Returns the first ID present in $after but not in $before. */
    private function diffTabIds(array $before, array $after): string
    {
        foreach ($after as $id) {
            if (!in_array($id, $before, true)) return $id;
        }
        return '';
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Single-document content population
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Fills a single Google Doc with content scraped from $urls.
     * Each URL is flushed as its own batchUpdate so indices always match
     * the live document state.
     */
    private function populateDocument(
        string $docId,
        string $langCode,
        array  $urls,
        string $selector,
        string $skipSelectors
    ): void {
        $idx = 1;

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

            $urlReqs = [];
            $this->htmlToRequests($extracted['html'], $urlReqs, $idx, '');

            if (!empty($urlReqs)) {
                foreach (array_chunk($urlReqs, 200) as $batch) {
                    $this->batchUpdate($docId, $batch);
                }
            }
        }
    }


    // ─────────────────────────────────────────────────────────────────────────
    // Document creation
    // ─────────────────────────────────────────────────────────────────────────

    private function createDocument(string $title): string
    {
        $data = $this->api('POST', self::DOCS_BASE, ['json' => ['title' => $title]]);
        return $data['documentId'];
    }

    /**
     * Moves a Drive file into the given folder.
     * Uses Drive API v3 files.update with addParents / removeParents.
     *
     * @throws \RuntimeException if the Drive API call fails
     */
    private function moveToFolder(string $fileId, string $folderId): void
    {
        // Fetch current parents so we can remove them (avoids duplicate locations)
        $meta    = $this->api('GET', self::DRIVE_BASE . "/{$fileId}?fields=parents");
        $parents = implode(',', $meta['parents'] ?? []);

        $url = self::DRIVE_BASE . "/{$fileId}"
             . "?addParents={$folderId}"
             . ($parents !== '' ? "&removeParents={$parents}" : '')
             . '&fields=id,parents';

        $this->api('PATCH', $url);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Tabs API support probe
    // ─────────────────────────────────────────────────────────────────────────

    private function probeTabsApi(string $docId): bool
    {
        // The Google Docs API currently exposes tab data on GET requests but does
        // NOT support insertTab in batchUpdate ("Unknown name" error). Until Google
        // adds write support we always use the section-based fallback. When they
        // do add it, remove the early return below and the test will run live.
        return false;

        // phpcs:ignore -- dead code kept intentionally for future activation
        try {
            $doc = $this->api('GET', self::DOCS_BASE . "/{$docId}?includeTabsContent=true");
            if (!isset($doc['tabs'])) return false;

            // Test whether insertTab batchUpdate is actually accepted
            $this->batchUpdate($docId, [[
                'insertTab' => [
                    'insertionIndex' => 0,
                    'tab'            => ['tabProperties' => ['title' => '__probe__']],
                ],
            ]]);
            return true;
        } catch (\Exception $e) {
            return false;
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Tab-based export (primary path)
    // ─────────────────────────────────────────────────────────────────────────

    private function exportWithTabs(
        string $docId,
        array  $languageGroups,
        string $selector,
        string $skipSelectors,
        array  $langCodes
    ): void {
        // Insert a freshly-named tab for every language.
        // We do NOT rename the auto-created default tab because
        // updateTabProperties is not yet supported by the API.
        // The default tab remains empty and harmless.
        $tabIds = [];

        foreach ($langCodes as $i => $code) {
            $title = strtoupper($code);
            $tabId = $this->insertTab($docId, $title, $i);
            $tabIds[$code] = $tabId;
            $this->info("Tab created: {$title} → {$tabId}");
        }

        foreach ($languageGroups as $code => $urls) {
            $tabId = $tabIds[$code] ?? null;
            if ($tabId === null || $tabId === '') {
                $this->info("No tab ID for [{$code}] — skipping");
                continue;
            }
            $this->populateTab($docId, $tabId, $code, $urls, $selector, $skipSelectors);
        }
    }

    private function insertTab(string $docId, string $title, int $insertionIndex): string
    {
        // Snapshot existing tab IDs so we can identify the newly created one
        $before    = $this->api('GET', self::DOCS_BASE . "/{$docId}?includeTabsContent=true");
        $beforeIds = [];
        foreach ($before['tabs'] ?? [] as $t) {
            $beforeIds[] = $t['tabProperties']['tabId'] ?? '';
        }

        $this->batchUpdate($docId, [[
            'insertTab' => [
                'insertionIndex' => $insertionIndex,
                'tab'            => ['tabProperties' => ['title' => $title]],
            ],
        ]]);

        // Find the tab that was not present before insertion
        $after = $this->api('GET', self::DOCS_BASE . "/{$docId}?includeTabsContent=true");
        foreach ($after['tabs'] ?? [] as $tab) {
            $tabId = $tab['tabProperties']['tabId'] ?? '';
            if ($tabId !== '' && !in_array($tabId, $beforeIds, true)) {
                return $tabId;
            }
        }
        return '';
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Section-based fallback export
    // ─────────────────────────────────────────────────────────────────────────

    private function exportWithSections(
        string $docId,
        array  $languageGroups,
        string $selector,
        string $skipSelectors
    ): void {
        $idx   = 1;
        $codes = array_keys($languageGroups);
        $last  = end($codes);

        foreach ($languageGroups as $code => $urls) {
            // ── Language heading ───────────────────────────────────────────
            $header = strtoupper($code);
            if (isset(self::$LANG_NAMES[$code])) {
                $header .= ' — ' . self::$LANG_NAMES[$code];
            }
            $headingReqs = [];
            $this->appendHeading($headingReqs, $idx, '', $header, 1);
            $this->batchUpdate($docId, $headingReqs);

            // ── One batchUpdate per URL so indices always match live state ─
            foreach ($urls as $url) {
                $this->info("  [{$code}] fetching {$url}");

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

                $urlReqs = [];
                $this->htmlToRequests($extracted['html'], $urlReqs, $idx, '');

                if (!empty($urlReqs)) {
                    foreach (array_chunk($urlReqs, 200) as $batch) {
                        $this->batchUpdate($docId, $batch);
                    }
                }
            }

            // ── Page break between language sections ───────────────────────
            if ($code !== $last) {
                $breakReqs = [];
                $this->appendRaw($breakReqs, $idx, '', "\n");
                $breakReqs[] = ['insertPageBreak' => ['location' => ['index' => $idx]]];
                $idx++;
                $this->batchUpdate($docId, $breakReqs);
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Tab content population
    // ─────────────────────────────────────────────────────────────────────────

    private function populateTab(
        string $docId,
        string $tabId,
        string $code,
        array  $urls,
        string $selector,
        string $skipSelectors
    ): void {
        $this->info("Populating [{$code}] — " . count($urls) . ' URL(s)');

        $idx      = 1;
        $requests = [];

        $header = strtoupper($code);
        if (isset(self::$LANG_NAMES[$code])) {
            $header .= ' — ' . self::$LANG_NAMES[$code];
        }
        $this->appendHeading($requests, $idx, $tabId, $header, 1);

        $lastIndex = count($urls) - 1;

        foreach ($urls as $i => $url) {
            $this->info("  [{$code}] {$url}");

            $html = $this->fetchHtml($url);
            if ($html === null) {
                $this->info("  Fetch failed: {$url}");
                continue;
            }

            $extracted = $this->extractContent($html, $selector, $skipSelectors);
            if (!$extracted['success']) {
                $this->info('  Extract failed: ' . $extracted['error']);
                continue;
            }

            $this->htmlToRequests($extracted['html'], $requests, $idx, $tabId);

            if ($i < $lastIndex) {
                $sep = "\n" . str_repeat('─', 50) . "\n";
                $this->appendRaw($requests, $idx, $tabId, $sep);
            }
        }

        foreach (array_chunk($requests, 500) as $batch) {
            $this->batchUpdate($docId, $batch);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // HTML → Docs requests
    // ─────────────────────────────────────────────────────────────────────────

    private function htmlToRequests(string $html, array &$requests, int &$idx, string $tabId): void
    {
        $html = $this->cleanHtml($html);
        if (trim($html) === '') return;

        $dom = new \DOMDocument();
        libxml_use_internal_errors(true);
        $dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
        libxml_clear_errors();

        $body = $dom->getElementsByTagName('body')->item(0);
        if ($body) {
            $this->nodeToRequests($body, $requests, $idx, $tabId, []);
        }
    }

    private function nodeToRequests(
        \DOMNode $node,
        array    &$requests,
        int      &$idx,
        string   $tabId,
        array    $inlineStyle
    ): void {
        foreach ($node->childNodes as $child) {
            if ($child->nodeType === XML_TEXT_NODE) {
                $text = $this->sanitize($child->nodeValue);
                if ($text !== '') {
                    $this->appendStyled($requests, $idx, $tabId, $text . "\n", 'NORMAL_TEXT', $inlineStyle);
                }
                continue;
            }

            if ($child->nodeType !== XML_ELEMENT_NODE) continue;

            $tag = strtolower($child->nodeName);

            // Skip decorative/non-content elements
            if (in_array($tag, ['script', 'style', 'svg', 'noscript', 'iframe', 'nav', 'header', 'footer'], true)) {
                continue;
            }

            switch ($tag) {
                case 'h1': case 'h2': case 'h3':
                case 'h4': case 'h5': case 'h6':
                    $text = $this->sanitize($this->textOf($child));
                    if ($text !== '') {
                        $this->appendHeading($requests, $idx, $tabId, $text, (int)$tag[1]);
                    }
                    break;

                case 'p':
                    $this->paragraphToRequests($child, $requests, $idx, $tabId, $inlineStyle);
                    break;

                case 'strong': case 'b':
                    $this->nodeToRequests($child, $requests, $idx, $tabId, array_merge($inlineStyle, ['bold' => true]));
                    break;

                case 'em': case 'i':
                    $this->nodeToRequests($child, $requests, $idx, $tabId, array_merge($inlineStyle, ['italic' => true]));
                    break;

                case 'u':
                    $this->nodeToRequests($child, $requests, $idx, $tabId, array_merge($inlineStyle, ['underline' => true]));
                    break;

                case 'a':
                    $href = $child->getAttribute('href');
                    $text = $this->sanitize($this->textOf($child));
                    if ($text !== '') {
                        $style = $inlineStyle;
                        if ($href !== '') $style['link'] = ['url' => $href];
                        $this->appendStyled($requests, $idx, $tabId, $text, 'NORMAL_TEXT', $style);
                    }
                    break;

                case 'ul':
                    $this->listToRequests($child, $requests, $idx, $tabId, false);
                    break;

                case 'ol':
                    $this->listToRequests($child, $requests, $idx, $tabId, true);
                    break;

                case 'table':
                    $this->tableToRequests($child, $requests, $idx, $tabId);
                    break;

                case 'br':
                    $this->appendRaw($requests, $idx, $tabId, "\n");
                    break;

                default:
                    if ($child->hasChildNodes()) {
                        $this->nodeToRequests($child, $requests, $idx, $tabId, $inlineStyle);
                    }
                    break;
            }
        }
    }

    /**
     * Converts a <p> element, preserving inline bold / italic / links.
     */
    private function paragraphToRequests(
        \DOMNode $node,
        array    &$requests,
        int      &$idx,
        string   $tabId,
        array    $inherited
    ): void {
        $segments = [];
        $this->gatherInline($node, $segments, $inherited);

        $hasText = false;
        foreach ($segments as $seg) {
            if (trim($seg['text']) !== '') { $hasText = true; break; }
        }
        if (!$hasText) return;

        foreach ($segments as $seg) {
            if ($seg['text'] === '') continue;
            $this->appendStyled($requests, $idx, $tabId, $seg['text'], 'NORMAL_TEXT', $seg['style']);
        }
        $this->appendRaw($requests, $idx, $tabId, "\n");
    }

    /**
     * Recursively collects inline segments with cumulative styles.
     */
    private function gatherInline(\DOMNode $node, array &$segments, array $style): void
    {
        static $blockTags = ['div','p','ul','ol','li','table','tr','td','th',
                             'h1','h2','h3','h4','h5','h6','blockquote','section','article'];

        foreach ($node->childNodes as $child) {
            if ($child->nodeType === XML_TEXT_NODE) {
                $text = $this->sanitize($child->nodeValue);
                if ($text !== '') {
                    $segments[] = ['text' => $text, 'style' => $style];
                }
                continue;
            }
            if ($child->nodeType !== XML_ELEMENT_NODE) continue;

            $tag = strtolower($child->nodeName);
            if (in_array($tag, $blockTags, true)) continue;

            switch ($tag) {
                case 'strong': case 'b':
                    $this->gatherInline($child, $segments, array_merge($style, ['bold' => true]));
                    break;
                case 'em': case 'i':
                    $this->gatherInline($child, $segments, array_merge($style, ['italic' => true]));
                    break;
                case 'u':
                    $this->gatherInline($child, $segments, array_merge($style, ['underline' => true]));
                    break;
                case 'a':
                    $href = $child->getAttribute('href');
                    $s    = $style;
                    if ($href !== '') $s['link'] = ['url' => $href];
                    $this->gatherInline($child, $segments, $s);
                    break;
                case 'br':
                    $segments[] = ['text' => "\n", 'style' => $style];
                    break;
                default:
                    $this->gatherInline($child, $segments, $style);
                    break;
            }
        }
    }

    private function listToRequests(
        \DOMNode $listNode,
        array    &$requests,
        int      &$idx,
        string   $tabId,
        bool     $ordered
    ): void {
        foreach ($listNode->childNodes as $child) {
            if (strtolower($child->nodeName) !== 'li') continue;
            $text = $this->sanitize($this->textOf($child));
            if ($text === '') continue;

            $start    = $idx;
            $fullText = $text . "\n";
            $this->appendRaw($requests, $idx, $tabId, $fullText);
            $end = $idx; // idx already advanced

            $preset = $ordered ? 'NUMBERED_DECIMAL_ALPHA_ROMAN' : 'BULLET_DISC_CIRCLE_SQUARE';
            $requests[] = [
                'createParagraphBullets' => [
                    'range'        => $this->range($start, $end, $tabId),
                    'bulletPreset' => $preset,
                ],
            ];
        }
    }

    /**
     * Renders a <table> as plain-text rows (header row in bold).
     * True insertTable support requires a separate document-fetch to get cell indices
     * and is omitted here to keep batch sizes manageable.
     */
    private function tableToRequests(\DOMNode $tableNode, array &$requests, int &$idx, string $tabId): void
    {
        $rows = [];
        $this->gatherTableRows($tableNode, $rows);
        if (empty($rows)) return;

        $this->appendRaw($requests, $idx, $tabId, "\n");

        foreach ($rows as $ri => $cells) {
            $rowText = implode(' | ', array_map([$this, 'sanitize'], $cells));
            if ($rowText === '') continue;
            $bold = ($ri === 0);
            $this->appendStyled($requests, $idx, $tabId, $rowText . "\n", 'NORMAL_TEXT', $bold ? ['bold' => true] : []);
        }

        $this->appendRaw($requests, $idx, $tabId, "\n");
    }

    private function gatherTableRows(\DOMNode $node, array &$rows): void
    {
        $tag = strtolower($node->nodeName);
        if ($tag === 'tr') {
            $cells = [];
            foreach ($node->childNodes as $child) {
                $ct = strtolower($child->nodeName);
                if ($ct === 'td' || $ct === 'th') {
                    $cells[] = $this->textOf($child);
                }
            }
            if (!empty($cells)) $rows[] = $cells;
            return;
        }
        foreach ($node->childNodes as $child) {
            if ($child->nodeType === XML_ELEMENT_NODE) {
                $this->gatherTableRows($child, $rows);
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Request-building helpers
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Inserts a heading paragraph and applies the named heading style.
     */
    private function appendHeading(array &$requests, int &$idx, string $tabId, string $text, int $level): void
    {
        if ($text === '') return;
        $text    = $this->sanitize($text);
        $start   = $idx;
        $full    = $text . "\n";
        $textLen = mb_strlen($full, 'UTF-8');

        $requests[] = ['insertText' => [
            'location' => $this->location($idx, $tabId),
            'text'     => $full,
        ]];

        $end = $idx + $textLen;

        $requests[] = ['updateParagraphStyle' => [
            'range'          => $this->range($start, $end, $tabId),
            'paragraphStyle' => ['namedStyleType' => 'HEADING_' . $level],
            'fields'         => 'namedStyleType',
        ]];

        $idx = $end;
    }

    /**
     * Inserts text with optional paragraph style and inline text style.
     * All characters in $text must be BMP Unicode (non-BMP stripped by sanitize()).
     */
    private function appendStyled(
        array  &$requests,
        int    &$idx,
        string $tabId,
        string $text,
        string $paraStyle,
        array  $textStyle
    ): void {
        if ($text === '') return;

        $start   = $idx;
        $textLen = mb_strlen($text, 'UTF-8');

        $requests[] = ['insertText' => [
            'location' => $this->location($idx, $tabId),
            'text'     => $text,
        ]];

        $end = $idx + $textLen;

        if ($paraStyle !== 'NORMAL_TEXT') {
            $requests[] = ['updateParagraphStyle' => [
                'range'          => $this->range($start, $end, $tabId),
                'paragraphStyle' => ['namedStyleType' => $paraStyle],
                'fields'         => 'namedStyleType',
            ]];
        }

        if (!empty($textStyle)) {
            // Don't apply text style over the trailing newline
            $styleEnd = (substr($text, -1) === "\n") ? $end - 1 : $end;
            if ($styleEnd > $start) {
                $apiStyle    = [];
                $styleFields = [];

                foreach ($textStyle as $key => $value) {
                    $apiStyle[$key]  = $value;
                    $styleFields[]   = $key;
                }

                $requests[] = ['updateTextStyle' => [
                    'range'     => $this->range($start, $styleEnd, $tabId),
                    'textStyle' => $apiStyle,
                    'fields'    => implode(',', $styleFields),
                ]];
            }
        }

        $idx = $end;
    }

    /**
     * Inserts raw text with no styling (separators, line-breaks, etc.).
     */
    private function appendRaw(array &$requests, int &$idx, string $tabId, string $text): void
    {
        if ($text === '') return;
        $requests[] = ['insertText' => [
            'location' => $this->location($idx, $tabId),
            'text'     => $text,
        ]];
        $idx += mb_strlen($text, 'UTF-8');
    }

    private function location(int $index, string $tabId): array
    {
        $loc = ['index' => $index];
        if ($tabId !== '') $loc['tabId'] = $tabId;
        return $loc;
    }

    private function range(int $start, int $end, string $tabId): array
    {
        $r = ['startIndex' => $start, 'endIndex' => $end];
        if ($tabId !== '') $r['tabId'] = $tabId;
        return $r;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Content fetching & HTML processing
    // ─────────────────────────────────────────────────────────────────────────

    private function fetchHtml(string $url, int $timeout = 30): ?string
    {
        $context = stream_context_create([
            'http' => [
                'timeout'         => $timeout,
                'user_agent'      => 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                'follow_location' => true,
            ],
            'ssl' => [
                'verify_peer'      => false,
                'verify_peer_name' => false,
            ],
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
                foreach ($xpath->query("//*[contains(concat(' ',normalize-space(@class),' '),' {$sel} ')]") as $n) $toRemove[] = $n;
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
        $html = preg_replace('/<(\w+)[^>]*>\s*<\/\1>/', '', $html);
        return trim($html);
    }

    private function sanitize(string $text): string
    {
        $text = html_entity_decode($text, ENT_QUOTES | ENT_HTML5, 'UTF-8');
        $text = preg_replace('/[\x{10000}-\x{10FFFF}]/u', '', $text);   // remove non-BMP (2 UTF-16 code units each)
        $text = preg_replace('/[\x{FFFE}\x{FFFF}]/u', '', $text);
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
    // HTTP / API helpers
    // ─────────────────────────────────────────────────────────────────────────

    private function batchUpdate(string $docId, array $requests): void
    {
        $this->api(
            'POST',
            self::DOCS_BASE . "/{$docId}:batchUpdate",
            ['json' => ['requests' => $requests]]
        );
    }

    /**
     * Makes an authorized HTTP request to the Google Docs API.
     *
     * @throws \RuntimeException on HTTP error
     */
    private function api(string $method, string $url, array $options = []): array
    {
        try {
            $response = $this->http->request($method, $url, array_merge(
                ['http_errors' => false],
                $options
            ));

            $body = (string)$response->getBody();
            $data = json_decode($body, true) ?? [];

            $statusCode = $response->getStatusCode();
            if ($statusCode >= 400) {
                $msg = $data['error']['message'] ?? "HTTP {$statusCode}";
                throw new \RuntimeException("Google API error: {$msg}");
            }

            return $data;
        } catch (\GuzzleHttp\Exception\GuzzleException $e) {
            throw new \RuntimeException('HTTP request failed: ' . $e->getMessage(), 0, $e);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Logging
    // ─────────────────────────────────────────────────────────────────────────

    private function info(string $message): void
    {
        $ts          = date('Y-m-d H:i:s');
        $this->log[] = "[{$ts}] {$message}";

        $logFile = dirname(__DIR__) . '/output/lang-debug.log';
        @file_put_contents($logFile, "[{$ts}] [GoogleDocs] {$message}\n", FILE_APPEND);
    }
}
