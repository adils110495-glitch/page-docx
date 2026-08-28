<?php
/**
 * Cyrillic Word Cleaner - replacement dictionary endpoint.
 *
 * Pipeline:  crilic-wordss.csv  ->  replacement dictionary (JSON)  ->  cleaner engine (JS)
 *
 * The CSV is the single source of truth. This endpoint parses it and also keeps a
 * JSON cache next to the tool so a large CSV is not re-parsed on every page load.
 * The cache is invalidated automatically whenever the CSV's mtime or size changes,
 * so editing crilic-wordss.csv is enough to update the cleaner - no code changes.
 *
 * Response: { ok, source, count, phraseCount, columns:{source,target}, delimiter,
 *             map:{ "<corrupted>":"<correct>" }, skipped, warnings[] }
 */

header('Content-Type: application/json; charset=utf-8');
header('X-Content-Type-Options: nosniff');

/* ------------------------------------------------------------------ *
 * Locate the CSV. A copy in output/ wins, so the dictionary can be
 * overridden per-deployment without editing the checked-in file.
 * ------------------------------------------------------------------ */
$candidates = [
    __DIR__ . '/output/crilic-wordss.csv',
    __DIR__ . '/crilic-wordss.csv',
];
$csvPath = null;
foreach ($candidates as $c) {
    if (is_file($c) && is_readable($c)) { $csvPath = $c; break; }
}

if ($csvPath === null) {
    http_response_code(404);
    echo json_encode([
        'ok'     => false,
        'error'  => 'crilic-wordss.csv was not found.',
        'looked' => array_map('basename', $candidates),
        'map'    => new stdClass(),
        'count'  => 0,
    ], JSON_UNESCAPED_UNICODE);
    exit;
}

$stat      = stat($csvPath);
$signature = $stat['mtime'] . '-' . $stat['size'];
// output/ is already writable and is only listed for .docx/.log, so a dot-file
// cache there stays out of the way - same convention as output/.settings.json.
$cacheDir  = __DIR__ . '/output';
$cacheFile = $cacheDir . '/.cyrillic-dictionary.json';

/* ---- Conditional GET: nothing changed since the browser last asked ---- */
$etag = '"' . md5($signature) . '"';
header('ETag: ' . $etag);
header('Cache-Control: no-cache, must-revalidate');
if (isset($_SERVER['HTTP_IF_NONE_MATCH']) && trim($_SERVER['HTTP_IF_NONE_MATCH']) === $etag) {
    http_response_code(304);
    exit;
}

/* ---- Serve from the JSON cache when the CSV has not changed ---- */
if (empty($_GET['refresh']) && is_file($cacheFile)) {
    $cached = json_decode((string)file_get_contents($cacheFile), true);
    if (is_array($cached) && isset($cached['signature']) && $cached['signature'] === $signature) {
        unset($cached['signature']);
        $cached['cached'] = true;
        echo json_encode($cached, JSON_UNESCAPED_UNICODE);
        exit;
    }
}

/* ------------------------------------------------------------------ *
 * Helpers
 * ------------------------------------------------------------------ */

/** Strip a UTF-8 BOM from the start of a string. */
function cc_strip_bom($s) {
    return (substr($s, 0, 3) === "\xEF\xBB\xBF") ? substr($s, 3) : $s;
}

/** Does the string contain any character outside plain ASCII? */
function cc_has_non_ascii($s) {
    return (bool)preg_match('/[^\x00-\x7F]/', $s);
}

/** Guess the delimiter from a data line. */
function cc_detect_delimiter($line) {
    $best = "\t";
    $bestCount = 0;
    foreach (array("\t", ',', ';', '|') as $d) {
        $n = substr_count($line, $d);
        if ($n > $bestCount) { $bestCount = $n; $best = $d; }
    }
    return $bestCount > 0 ? $best : "\t";
}

/** Normalise a header cell for matching. */
function cc_norm($s) {
    $s = strtolower(trim(cc_strip_bom((string)$s)));
    $s = preg_replace('/[^a-z0-9]+/', ' ', $s);
    return trim($s);
}

/**
 * Pick the source (corrupted) and target (correct) columns.
 *
 * Header names are tried first; if that fails we fall back to content analysis -
 * the most "non-ASCII heavy" single-word column is the source, and the most
 * ASCII single-word column is the target.
 */
function cc_pick_columns($header, $rows) {
    // Columns that are never a word mapping, even if their name matches a hint.
    $blocked = array('codepoint', 'occurrence', 'count', 'pair', 'frequency', 'note', 'comment', 'total');

    $sourceHints = array(
        'word as published', 'word published', 'published', 'cyrillic incorrect word',
        'incorrect word', 'incorrect', 'corrupted', 'original', 'source', 'wrong',
        'bad', 'from', 'cyrillic word', 'cyrillic',
    );
    $targetHints = array(
        'correct latin form', 'correct english word', 'correct latin', 'correct english',
        'correct form', 'correct word', 'correct', 'latin form', 'replacement', 'replace with',
        'english', 'latin', 'target', 'fixed', 'clean',
    );

    $normHeader = array_map('cc_norm', $header);

    $find = function ($hints) use ($normHeader, $blocked) {
        // 1. exact header-name match
        foreach ($hints as $hint) {
            foreach ($normHeader as $i => $h) {
                if ($h !== '' && $h === $hint) return $i;
            }
        }
        // 2. substring match, skipping obviously wrong columns
        foreach ($hints as $hint) {
            foreach ($normHeader as $i => $h) {
                if ($h === '') continue;
                $isBlocked = false;
                foreach ($blocked as $b) {
                    if (strpos($h, $b) !== false) { $isBlocked = true; break; }
                }
                if ($isBlocked) continue;
                if (strpos($h, $hint) !== false) return $i;
            }
        }
        return -1;
    };

    $src = $find($sourceHints);
    $tgt = $find($targetHints);
    if ($tgt === $src) $tgt = -1;

    if ($src === -1 || $tgt === -1) {
        $colCount = count($header);
        $nonAscii = array_fill(0, $colCount, 0);
        $ascii    = array_fill(0, $colCount, 0);
        $spaces   = array_fill(0, $colCount, 0);

        foreach ($rows as $r) {
            for ($i = 0; $i < $colCount; $i++) {
                $v = isset($r[$i]) ? trim((string)$r[$i]) : '';
                if ($v === '') continue;
                if (cc_has_non_ascii($v)) { $nonAscii[$i]++; } else { $ascii[$i]++; }
                if (strpos($v, ' ') !== false) $spaces[$i]++;
            }
        }

        if ($src === -1) {
            $bestScore = -PHP_INT_MAX;
            for ($i = 0; $i < $colCount; $i++) {
                // single corrupted words, not whole sentences
                $score = $nonAscii[$i] - ($spaces[$i] * 2);
                if ($score > $bestScore) { $bestScore = $score; $src = $i; }
            }
        }
        if ($tgt === -1) {
            $bestScore = -PHP_INT_MAX;
            for ($i = 0; $i < $colCount; $i++) {
                if ($i === $src) continue;
                $score = $ascii[$i] - ($spaces[$i] * 2) - ($nonAscii[$i] * 3);
                if ($score > $bestScore) { $bestScore = $score; $tgt = $i; }
            }
        }
    }

    return array($src, $tgt);
}

/* ------------------------------------------------------------------ *
 * Parse
 * ------------------------------------------------------------------ */
$raw = (string)file_get_contents($csvPath);
$raw = cc_strip_bom($raw);
$raw = str_replace(array("\r\n", "\r"), "\n", $raw);
$lines = explode("\n", $raw);

$delimiter = "\t";
foreach ($lines as $l) {
    if (trim($l) !== '') { $delimiter = cc_detect_delimiter($l); break; }
}

$parsed = array();
foreach ($lines as $line) {
    if (trim($line) === '') continue;
    $parsed[] = str_getcsv($line, $delimiter, '"', "\\");
}

$warnings = array();

if (count($parsed) < 2) {
    echo json_encode(array(
        'ok'    => false,
        'error' => 'crilic-wordss.csv has no data rows.',
        'map'   => new stdClass(),
        'count' => 0,
    ), JSON_UNESCAPED_UNICODE);
    exit;
}

$header = array_map(function ($v) { return trim((string)$v); }, $parsed[0]);
$rows   = array_slice($parsed, 1);

// A CSV with no header row at all: the first cell already looks like corrupted data.
if (count($header) >= 2 && cc_has_non_ascii(isset($header[0]) ? $header[0] : '')) {
    array_unshift($rows, $header);
    $generated = array();
    foreach (array_keys($header) as $i) { $generated[] = 'column ' . ($i + 1); }
    $header = $generated;
    $warnings[] = 'No header row detected - columns were chosen by content.';
}

$cols   = cc_pick_columns($header, $rows);
$srcCol = $cols[0];
$tgtCol = $cols[1];

if ($srcCol < 0 || $tgtCol < 0 || $srcCol === $tgtCol) {
    echo json_encode(array(
        'ok'     => false,
        'error'  => 'Could not identify the corrupted-word and correct-word columns in crilic-wordss.csv.',
        'header' => $header,
        'map'    => new stdClass(),
        'count'  => 0,
    ), JSON_UNESCAPED_UNICODE);
    exit;
}

$map         = array();
$skipped     = 0;
$phraseCount = 0;

foreach ($rows as $r) {
    $from = isset($r[$srcCol]) ? trim((string)$r[$srcCol]) : '';
    $to   = isset($r[$tgtCol]) ? trim((string)$r[$tgtCol]) : '';

    if ($from === '' || $to === '')  { $skipped++; continue; }
    if ($from === $to)               { $skipped++; continue; }
    if (!cc_has_non_ascii($from))    { $skipped++; continue; } // nothing to fix
    if (isset($map[$from])) {
        if ($map[$from] !== $to) {
            $warnings[] = 'Duplicate mapping for "' . $from . '" - kept the first one.';
        }
        $skipped++;
        continue;
    }

    if (preg_match('/\s/u', $from)) $phraseCount++;
    $map[$from] = $to;
}

$payload = array(
    'ok'          => true,
    'source'      => basename($csvPath),
    'delimiter'   => $delimiter === "\t" ? 'tab' : $delimiter,
    'columns'     => array(
        'source' => isset($header[$srcCol]) ? $header[$srcCol] : '',
        'target' => isset($header[$tgtCol]) ? $header[$tgtCol] : '',
    ),
    'count'       => count($map),
    'phraseCount' => $phraseCount,
    'skipped'     => $skipped,
    'warnings'    => array_values(array_unique($warnings)),
    'generated'   => gmdate('c'),
    'csvModified' => gmdate('c', $stat['mtime']),
    'map'         => $map,
);

/* ---- Refresh the JSON cache (best effort - never fatal) ---- */
if (!is_dir($cacheDir)) { @mkdir($cacheDir, 0775, true); }
$toCache = $payload;
$toCache['signature'] = $signature;
@file_put_contents($cacheFile, json_encode($toCache, JSON_UNESCAPED_UNICODE), LOCK_EX);

echo json_encode($payload, JSON_UNESCAPED_UNICODE);
