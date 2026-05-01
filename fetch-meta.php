<?php
header('Content-Type: application/json');
set_time_limit(120);

// ── Helpers ───────────────────────────────────────────────────────────────────

function isValidUrl(string $url): bool {
    $url = trim($url);
    if (empty($url)) return false;
    if (!preg_match('/^https?:\/\//i', $url)) return false;
    return filter_var($url, FILTER_VALIDATE_URL) !== false;
}

function sanitizeProjectName(string $project): string {
    $slug = preg_replace('/[^a-z0-9\-_]+/i', '-', $project);
    $slug = trim($slug, '-');
    $slug = strtolower($slug);
    if (strlen($slug) > 50) $slug = substr($slug, 0, 50);
    return $slug ?: 'default';
}

/**
 * Fetch multiple URLs in parallel using cURL multi.
 * Falls back to sequential file_get_contents when cURL is unavailable.
 */
function fetchParallel(array $urls, int $timeout = 10): array {
    if (!function_exists('curl_multi_init')) {
        $results = [];
        foreach ($urls as $url) {
            $ctx  = stream_context_create([
                'http' => ['timeout' => $timeout, 'user_agent' => 'Mozilla/5.0', 'follow_location' => true],
                'ssl'  => ['verify_peer' => false, 'verify_peer_name' => false],
            ]);
            $html = @file_get_contents($url, false, $ctx);
            $results[$url] = $html !== false && !empty(trim($html))
                ? ['html' => $html, 'error' => null]
                : ['html' => null,  'error' => 'Could not fetch URL'];
        }
        return $results;
    }

    $mh      = curl_multi_init();
    $handles = [];

    foreach ($urls as $url) {
        $ch = curl_init();
        curl_setopt_array($ch, [
            CURLOPT_URL            => $url,
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_TIMEOUT        => $timeout,
            CURLOPT_CONNECTTIMEOUT => 5,
            CURLOPT_FOLLOWLOCATION => true,
            CURLOPT_MAXREDIRS      => 5,
            CURLOPT_SSL_VERIFYPEER => false,
            CURLOPT_SSL_VERIFYHOST => false,
            CURLOPT_USERAGENT      => 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            CURLOPT_ENCODING       => '',
            CURLOPT_HTTPHEADER     => ['Accept-Language: en-US,en;q=0.9', 'Accept: text/html,application/xhtml+xml'],
        ]);
        curl_multi_add_handle($mh, $ch);
        $handles[$url] = $ch;
    }

    do {
        $status = curl_multi_exec($mh, $running);
        if ($running) curl_multi_select($mh, 0.5);
    } while ($running > 0 && $status === CURLM_OK);

    $results = [];
    foreach ($handles as $url => $ch) {
        $html  = curl_multi_getcontent($ch);
        $errno = curl_errno($ch);
        $error = curl_error($ch);
        curl_multi_remove_handle($mh, $ch);
        curl_close($ch);

        if ($errno !== 0 || $html === false || empty(trim($html))) {
            $results[$url] = ['html' => null, 'error' => $error ?: 'Empty or failed response'];
        } else {
            $results[$url] = ['html' => $html, 'error' => null];
        }
    }

    curl_multi_close($mh);
    return $results;
}

/**
 * Fetch a batch with retry logic.
 * Each failed URL is retried up to $maxRetries times with progressive delays.
 * Uses parallel fetching within each attempt round.
 */
function fetchBatchWithRetry(array $urls, int $maxRetries = 3, int $timeout = 10): array {
    $finalResults = [];
    $toFetch      = $urls;

    for ($attempt = 1; $attempt <= $maxRetries; $attempt++) {
        if (empty($toFetch)) break;

        // Progressive delay: 0s on first attempt, 1s on second, 2s on third
        if ($attempt > 1) sleep($attempt - 1);

        $batchResults = fetchParallel($toFetch, $timeout);
        $stillFailing = [];

        foreach ($batchResults as $url => $result) {
            if ($result['html'] !== null) {
                $finalResults[$url] = [
                    'html'     => $result['html'],
                    'error'    => null,
                    'attempts' => $attempt,
                ];
            } else {
                $stillFailing[] = $url;
                // On final attempt, record the failure
                if ($attempt === $maxRetries) {
                    $finalResults[$url] = [
                        'html'     => null,
                        'error'    => $result['error'] ?: 'Failed after ' . $maxRetries . ' attempts',
                        'attempts' => $attempt,
                    ];
                }
            }
        }

        $toFetch = $stillFailing;
    }

    return $finalResults;
}

/**
 * Extract meta title and description from an HTML string.
 * Falls back to og:description when name="description" is absent.
 */
function extractMeta(string $html): array {
    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $dom->loadHTML(mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'));
    libxml_clear_errors();

    // <title>
    $title = '';
    $titleTags = $dom->getElementsByTagName('title');
    if ($titleTags->length > 0) {
        $title = trim($titleTags->item(0)->textContent);
    }

    // <meta name="description"> — case-insensitive
    $description = '';
    $xpath = new DOMXPath($dom);
    $metaTags = $xpath->query(
        '//meta[translate(@name,"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz")="description"]'
    );
    if ($metaTags->length > 0) {
        $description = trim($metaTags->item(0)->getAttribute('content'));
    }

    // Fallback: og:description
    if ($description === '') {
        $ogMeta = $xpath->query('//meta[@property="og:description"]');
        if ($ogMeta->length > 0) {
            $description = trim($ogMeta->item(0)->getAttribute('content'));
        }
    }

    return ['title' => $title, 'description' => $description];
}

/**
 * Write CSV to the output directory and return a web-relative path.
 */
function saveCsv(array $rows, string $project = ''): string {
    $outputDir   = __DIR__ . '/output';
    $projectSlug = '';

    if (!empty($project)) {
        $projectSlug = sanitizeProjectName($project);
        $outputDir  .= '/' . $projectSlug;
    }

    if (!is_dir($outputDir)) {
        mkdir($outputDir, 0777, true);
    }

    $filename = 'meta_' . date('Ymd_His') . '.csv';
    $filepath = $outputDir . '/' . $filename;

    $fp = fopen($filepath, 'w');
    fputs($fp, "\xEF\xBB\xBF"); // UTF-8 BOM for Excel
    fputcsv($fp, ['Website URL', 'Meta Title', 'Meta Description']);
    foreach ($rows as $row) {
        fputcsv($fp, [$row['url'], $row['meta_title'], $row['meta_description']]);
    }
    fclose($fp);

    // Return a clean forward-slash relative path
    $relative = 'output'
        . ($projectSlug !== '' ? '/' . $projectSlug : '')
        . '/' . $filename;

    return $relative;
}

// ── Router ────────────────────────────────────────────────────────────────────

$raw   = file_get_contents('php://input');
$input = json_decode($raw, true);

if (!is_array($input) || !isset($input['action'])) {
    http_response_code(400);
    echo json_encode(['success' => false, 'message' => 'Expected JSON body with an action field.']);
    exit;
}

// ── action = process ──────────────────────────────────────────────────────────
if ($input['action'] === 'process') {

    $urls       = $input['urls'] ?? [];
    $maxRetries = min(max((int)($input['max_retries'] ?? 3), 1), 3);

    if (!is_array($urls) || empty($urls)) {
        echo json_encode(['success' => false, 'message' => 'No URLs supplied in this batch.']);
        exit;
    }

    // Cap per-request batch at 20 to prevent overlong requests
    $urls = array_values(array_filter(
        array_map('trim', array_slice($urls, 0, 20)),
        fn($u) => $u !== ''
    ));

    // Separate valid from invalid up front so we don't waste network calls
    $validUrls   = array_values(array_filter($urls, 'isValidUrl'));
    $invalidUrls = array_diff($urls, $validUrls);

    $fetchResults = fetchBatchWithRetry($validUrls, $maxRetries);

    $results = [];
    foreach ($urls as $url) {
        if (!isValidUrl($url)) {
            $results[] = [
                'url'              => $url,
                'meta_title'       => 'Error: Invalid URL format',
                'meta_description' => '',
                'status'           => 'error',
                'attempts'         => 0,
            ];
            continue;
        }

        $fr = $fetchResults[$url] ?? ['html' => null, 'error' => 'Not processed', 'attempts' => 0];

        if ($fr['html'] === null) {
            $results[] = [
                'url'              => $url,
                'meta_title'       => 'Error: ' . ($fr['error'] ?: 'Could not fetch'),
                'meta_description' => '',
                'status'           => 'error',
                'attempts'         => (int)$fr['attempts'],
            ];
        } else {
            $meta = extractMeta($fr['html']);
            $results[] = [
                'url'              => $url,
                'meta_title'       => $meta['title'],
                'meta_description' => $meta['description'],
                'status'           => 'success',
                'attempts'         => (int)$fr['attempts'],
            ];
        }
    }

    echo json_encode(['success' => true, 'results' => $results]);
    exit;
}

// ── action = finalize ─────────────────────────────────────────────────────────
if ($input['action'] === 'finalize') {

    $results = $input['results'] ?? [];
    $project = trim((string)($input['project'] ?? ''));

    if (!is_array($results) || empty($results)) {
        echo json_encode(['success' => false, 'message' => 'No results provided to save.']);
        exit;
    }

    $rows = [];
    foreach ($results as $r) {
        if (!isset($r['url'])) continue;
        $rows[] = [
            'url'              => substr(strip_tags((string)$r['url']), 0, 2048),
            'meta_title'       => substr((string)($r['meta_title']       ?? ''), 0, 512),
            'meta_description' => substr((string)($r['meta_description'] ?? ''), 0, 2048),
        ];
    }

    if (empty($rows)) {
        echo json_encode(['success' => false, 'message' => 'No valid rows to write.']);
        exit;
    }

    try {
        $csvFile = saveCsv($rows, $project);
        echo json_encode(['success' => true, 'csv_file' => $csvFile]);
    } catch (Exception $e) {
        http_response_code(500);
        echo json_encode(['success' => false, 'message' => 'Failed to write CSV: ' . $e->getMessage()]);
    }

    exit;
}

http_response_code(400);
echo json_encode(['success' => false, 'message' => 'Unknown action: ' . htmlspecialchars((string)$input['action'])]);
