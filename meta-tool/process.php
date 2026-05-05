<?php

declare(strict_types=1);

session_start();
set_time_limit(300);

require_once __DIR__ . '/lib/helper.php';
require_once __DIR__ . '/lib/deepl.php';
require_once __DIR__ . '/lib/gemini.php';
require_once __DIR__ . '/lib/groq.php';
require_once __DIR__ . '/lib/validator.php';
require_once __DIR__ . '/lib/logger.php';

load_env(__DIR__ . '/.env');

// Serve the output CSV as a download
if (isset($_GET['download'])) {
    $output_file = __DIR__ . '/output.csv';
    if (file_exists($output_file)) {
        header('Content-Type: text/csv; charset=UTF-8');
        header('Content-Disposition: attachment; filename="seo-meta-translated.csv"');
        header('Content-Length: ' . filesize($output_file));
        readfile($output_file);
        exit;
    }
    header('Location: index.php');
    exit;
}

// ── Core pipeline functions ────────────────────────────────────────────────

function process_field(
    string $original,
    string $target_lang,
    string $lang_name,
    string $field_type,
    string $log_file
): array {
    $original = clean_text($original);
    $min      = $field_type === 'title' ? 40  : 150;
    $max      = $field_type === 'title' ? 46  : 155;
    $validate = $field_type === 'title' ? 'validate_title' : 'validate_description';

    $attempts = 0;
    $log_data = [
        'original' => $original,
        'deepl'    => '-',
        'gemini'   => '-',
        'groq'     => '-',
    ];

    $best           = null;   // closest to valid length, regardless of meaning
    $best_dist      = PHP_INT_MAX;
    $best_meaning   = null;   // valid length + meaning confirmed
    $deepl_ref      = '';     // passed to AI rewrites as vocabulary reference

    $update_best = function (string $text) use (&$best, &$best_dist, $min, $max): void {
        $dist = distance_to_range(mb_strlen($text, 'UTF-8'), $min, $max);
        if ($dist < $best_dist) {
            $best      = $text;
            $best_dist = $dist;
        }
    };

    // Step 1: DeepL translation
    $translated = deepl_translate($original, $target_lang);
    $attempts++;

    if ($translated) {
        $translated        = clean_text($translated);
        $log_data['deepl'] = $translated;
        $deepl_ref         = $translated;
        $update_best($translated);

        if ($validate($translated)) {
            if (validate_meaning($original, $translated, $lang_name)) {
                return finish($log_data, $translated, $attempts, 'deepl_ok', $log_file, $min, $max, ['deepl_ref' => $deepl_ref]);
            }
            // Length OK but meaning changed — keep as fallback, continue rewriting
            $best_meaning = $best_meaning ?? $translated;
        }
    }

    // Step 2: Gemini rewrite — 2 attempts (always rewrites from original English)
    for ($i = 0; $i < 2; $i++) {
        $text = gemini_rewrite($original, $lang_name, $field_type, $deepl_ref);
        $attempts++;

        if ($text) {
            $text               = clean_text($text);
            $log_data['gemini'] = $text;
            $update_best($text);

            if ($validate($text)) {
                if (validate_meaning($original, $text, $lang_name)) {
                    return finish($log_data, $text, $attempts, 'gemini_ok', $log_file, $min, $max, ['deepl_ref' => $deepl_ref]);
                }
                $best_meaning = $best_meaning ?? $text;
            }
        }
    }

    // Step 3: Groq rewrite — 2 attempts (always rewrites from original English)
    for ($i = 0; $i < 2; $i++) {
        $text = groq_rewrite($original, $lang_name, $field_type, $deepl_ref);
        $attempts++;

        if ($text) {
            $text             = clean_text($text);
            $log_data['groq'] = $text;
            $update_best($text);

            if ($validate($text)) {
                if (validate_meaning($original, $text, $lang_name)) {
                    return finish($log_data, $text, $attempts, 'groq_ok', $log_file, $min, $max, ['deepl_ref' => $deepl_ref]);
                }
                $best_meaning = $best_meaning ?? $text;
            }
        }
    }

    // Step 4: Prefer a length-valid result (even if meaning check failed) over a hard trim
    if ($best_meaning !== null) {
        return finish($log_data, $best_meaning, $attempts, 'meaning_fallback', $log_file, $min, $max, ['deepl_ref' => $deepl_ref]);
    }

    // Step 5: Hard trim of closest-length candidate — absolute last resort
    $fallback = hard_trim($best ?? $original, $max);
    return finish($log_data, $fallback, $attempts, 'fallback_trim', $log_file, $min, $max, ['deepl_ref' => $deepl_ref]);
}

function finish(
    array  $log_data,
    string $text,
    int    $attempts,
    string $status,
    string $log_file,
    int    $min,
    int    $max,
    array  $extra = []
): array {
    $log_data['final']    = $text;
    $log_data['attempts'] = $attempts;
    $log_data['status']   = $status;
    log_row($log_data, $log_file);

    $len = mb_strlen($text, 'UTF-8');
    return array_merge([
        'text'     => $text,
        'status'   => $status,
        'attempts' => $attempts,
        'len'      => $len,
        'valid'    => ($len >= $min && $len <= $max),
    ], $extra);
}

/**
 * Generates one alternative title/description using AI.
 * Called only when the primary result came from AI rewriting (not plain DeepL).
 * Guarantees the result is not identical to $primary_text.
 */
function generate_alt(
    string $original,
    string $lang_name,
    string $field_type,
    string $deepl_ref,
    string $primary_text
): string {
    $validate = $field_type === 'title' ? 'validate_title' : 'validate_description';

    // Try Gemini — pass primary as exclude so the model produces different wording
    $text = gemini_rewrite($original, $lang_name, $field_type, $deepl_ref, $primary_text);
    if ($text) {
        $text = clean_text($text);
        if ($text !== $primary_text && $validate($text) && validate_meaning($original, $text, $lang_name)) {
            return $text;
        }
    }

    // Fallback to Groq
    $text = groq_rewrite($original, $lang_name, $field_type, $deepl_ref, $primary_text);
    if ($text) {
        $text = clean_text($text);
        if ($text !== $primary_text && $validate($text)) {
            return $text;
        }
    }

    return '';
}

function process_csv(
    string $input_path,
    string $target_lang,
    string $output_path,
    string $log_file
): array {
    $lang_name = get_language_name($target_lang);

    $in = @fopen($input_path, 'r');
    if (!$in) return ['error' => 'Cannot open uploaded file.'];

    $headers = fgetcsv($in);
    if (!$headers) {
        fclose($in);
        return ['error' => 'CSV file is empty.'];
    }

    $headers   = array_map('trim', $headers);
    $title_idx = array_search('meta_title', $headers);
    $desc_idx  = array_search('meta_description', $headers);

    $out = @fopen($output_path, 'w');
    if (!$out) {
        fclose($in);
        return ['error' => 'Cannot write output file.'];
    }

    // UTF-8 BOM — keeps Excel from mangling multibyte characters
    fwrite($out, "\xEF\xBB\xBF");
    fputcsv($out, array_merge($headers, ['new_title', 'alt_title', 'new_description']));

    init_log($log_file);

    // Statuses that mean AI rewrote the title (alt title should be generated)
    $regenerated_statuses = ['gemini_ok', 'groq_ok', 'meaning_fallback', 'fallback_trim'];

    $rows = [];
    while (($row = fgetcsv($in)) !== false) {
        $r_title   = ['text' => '', 'status' => 'skipped', 'len' => 0, 'valid' => true, 'attempts' => 0, 'deepl_ref' => ''];
        $r_desc    = ['text' => '', 'status' => 'skipped', 'len' => 0, 'valid' => true, 'attempts' => 0];
        $alt_title = '';

        if ($title_idx !== false && trim($row[$title_idx] ?? '') !== '') {
            $r_title = process_field($row[$title_idx], $target_lang, $lang_name, 'title', $log_file);

            // Generate an alternate title only when AI regeneration was needed
            if (in_array($r_title['status'], $regenerated_statuses, true)) {
                $alt_title = generate_alt(
                    $row[$title_idx],
                    $lang_name,
                    'title',
                    $r_title['deepl_ref'] ?? '',
                    $r_title['text']           // must not duplicate the primary title
                );
            }
        }

        if ($desc_idx !== false && trim($row[$desc_idx] ?? '') !== '') {
            $r_desc = process_field($row[$desc_idx], $target_lang, $lang_name, 'description', $log_file);
        }

        fputcsv($out, array_merge($row, [$r_title['text'], $alt_title, $r_desc['text']]));

        $rows[] = [
            'original_title' => $title_idx !== false ? ($row[$title_idx] ?? '') : '',
            'original_desc'  => $desc_idx  !== false ? ($row[$desc_idx]  ?? '') : '',
            'result_title'   => $r_title,
            'alt_title'      => $alt_title,
            'result_desc'    => $r_desc,
        ];
    }

    fclose($in);
    fclose($out);

    return ['processed' => count($rows), 'rows' => $rows];
}

// ── Web request handler ────────────────────────────────────────────────────

if (!isset($_SESSION['input_file'], $_SESSION['target_lang'])) {
    header('Location: index.php');
    exit;
}

$input_file  = $_SESSION['input_file'];
$target_lang = $_SESSION['target_lang'];

if (!file_exists($input_file)) {
    $_SESSION['error'] = 'Uploaded file not found. Please try again.';
    header('Location: index.php');
    exit;
}

$output_file = __DIR__ . '/output.csv';
$log_file    = __DIR__ . '/logs/process.log';

$result = process_csv($input_file, $target_lang, $output_file, $log_file);

@unlink($input_file);
unset($_SESSION['input_file'], $_SESSION['target_lang']);

$lang_label = get_language_name($target_lang);
$has_error  = isset($result['error']);

$status_labels = [
    'deepl_ok'        => ['label' => 'DeepL',          'class' => 'badge-blue'],
    'gemini_ok'       => ['label' => 'Gemini',         'class' => 'badge-purple'],
    'groq_ok'         => ['label' => 'Groq',           'class' => 'badge-orange'],
    'meaning_fallback'=> ['label' => 'Length OK',      'class' => 'badge-orange'],
    'fallback_trim'   => ['label' => 'Trimmed',        'class' => 'badge-red'],
    'skipped'         => ['label' => 'Skipped',        'class' => 'badge-gray'],
];

?>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Processing Results — SEO Meta Tool</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f5f7fa; color: #1a1a2e; }
  .header { background: #1a1a2e; color: #fff; padding: 1rem 2rem; display: flex; align-items: center; justify-content: space-between; }
  .header h1 { font-size: 1.2rem; font-weight: 600; }
  .container { max-width: 1200px; margin: 2rem auto; padding: 0 1.5rem; }
  .card { background: #fff; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,.08); padding: 1.5rem; margin-bottom: 1.5rem; }
  .summary { display: flex; gap: 1.5rem; flex-wrap: wrap; }
  .stat { flex: 1; min-width: 140px; text-align: center; padding: 1rem; background: #f0f4ff; border-radius: 8px; }
  .stat-num { font-size: 2rem; font-weight: 700; color: #3b5bdb; }
  .stat-label { font-size: .8rem; color: #555; margin-top: .2rem; }
  .btn { display: inline-flex; align-items: center; gap: .4rem; padding: .6rem 1.2rem; border-radius: 6px; font-size: .9rem; font-weight: 500; text-decoration: none; cursor: pointer; border: none; }
  .btn-primary { background: #3b5bdb; color: #fff; }
  .btn-primary:hover { background: #2f4ac7; }
  .btn-secondary { background: #e9ecef; color: #333; }
  .btn-secondary:hover { background: #dee2e6; }
  .actions { display: flex; gap: .8rem; margin-bottom: 1.5rem; }
  table { width: 100%; border-collapse: collapse; font-size: .85rem; }
  th { background: #f8f9fa; text-align: left; padding: .7rem .9rem; font-weight: 600; color: #444; border-bottom: 2px solid #e9ecef; white-space: nowrap; }
  td { padding: .65rem .9rem; border-bottom: 1px solid #f0f0f0; vertical-align: top; }
  tr:hover td { background: #fafbff; }
  .badge { display: inline-block; padding: .15rem .55rem; border-radius: 20px; font-size: .72rem; font-weight: 600; }
  .badge-blue   { background: #dbe4ff; color: #3b5bdb; }
  .badge-purple { background: #f3d9fa; color: #862e9c; }
  .badge-orange { background: #ffe8cc; color: #d9480f; }
  .badge-red    { background: #ffe3e3; color: #c92a2a; }
  .badge-gray   { background: #f1f3f5; color: #666; }
  .badge-green  { background: #d3f9d8; color: #2b8a3e; }
  .len { font-size: .75rem; color: #888; margin-top: .2rem; }
  .len.valid   { color: #2b8a3e; }
  .len.invalid { color: #c92a2a; }
  .text-cell { max-width: 260px; }
  .text-cell span { display: block; line-height: 1.4; }
  .error-box { background: #fff5f5; border: 1px solid #ffc9c9; border-radius: 8px; padding: 1rem 1.2rem; color: #c92a2a; }
  .section-title { font-size: 1rem; font-weight: 600; margin-bottom: 1rem; color: #333; }
  .overflow-x { overflow-x: auto; }
</style>
</head>
<body>

<div class="header">
  <h1>SEO Meta Translation Tool</h1>
  <span style="font-size:.85rem;opacity:.7">Target: <?= htmlspecialchars($lang_label) ?></span>
</div>

<div class="container">

<?php if ($has_error): ?>
  <div class="card">
    <div class="error-box"><?= htmlspecialchars($result['error']) ?></div>
    <div style="margin-top:1rem">
      <a href="index.php" class="btn btn-secondary">← Back</a>
    </div>
  </div>

<?php else:
  $total   = $result['processed'];
  $valid_t = array_filter($result['rows'], fn($r) => $r['result_title']['valid']);
  $valid_d = array_filter($result['rows'], fn($r) => $r['result_desc']['valid']);
  $trimmed = array_filter($result['rows'], fn($r) =>
    $r['result_title']['status'] === 'fallback_trim' || $r['result_desc']['status'] === 'fallback_trim'
  );
?>

  <div class="card">
    <div class="summary">
      <div class="stat">
        <div class="stat-num"><?= $total ?></div>
        <div class="stat-label">Rows processed</div>
      </div>
      <div class="stat">
        <div class="stat-num"><?= count($valid_t) ?></div>
        <div class="stat-label">Valid titles</div>
      </div>
      <div class="stat">
        <div class="stat-num"><?= count($valid_d) ?></div>
        <div class="stat-label">Valid descriptions</div>
      </div>
      <div class="stat">
        <div class="stat-num"><?= count($trimmed) ?></div>
        <div class="stat-label">Hard-trimmed rows</div>
      </div>
    </div>
  </div>

  <div class="actions">
    <a href="process.php?download=1" class="btn btn-primary">⬇ Download output.csv</a>
    <a href="index.php" class="btn btn-secondary">← Process another file</a>
  </div>

  <div class="card">
    <div class="section-title">Results</div>
    <div class="overflow-x">
    <table>
      <thead>
        <tr>
          <th>#</th>
          <th>Original Title</th>
          <th>New Title</th>
          <th>Alt Title</th>
          <th>Title Status</th>
          <th>Original Description</th>
          <th>New Description</th>
          <th>Desc Status</th>
        </tr>
      </thead>
      <tbody>
      <?php foreach ($result['rows'] as $i => $row):
        $rt = $row['result_title'];
        $rd = $row['result_desc'];
        $ts = $status_labels[$rt['status']] ?? ['label' => $rt['status'], 'class' => 'badge-gray'];
        $ds = $status_labels[$rd['status']] ?? ['label' => $rd['status'], 'class' => 'badge-gray'];
      ?>
        <tr>
          <td><?= $i + 1 ?></td>
          <td class="text-cell"><span><?= htmlspecialchars($row['original_title']) ?></span></td>
          <td class="text-cell">
            <span><?= htmlspecialchars($rt['text']) ?></span>
            <?php if ($rt['text'] !== ''): ?>
            <div class="len <?= $rt['valid'] ? 'valid' : 'invalid' ?>">
              <?= $rt['len'] ?> chars <?= $rt['valid'] ? '✓' : '✗' ?>
            </div>
            <?php endif ?>
          </td>
          <td class="text-cell">
            <?php if ($row['alt_title'] !== ''): ?>
              <span><?= htmlspecialchars($row['alt_title']) ?></span>
              <div class="len valid"><?= mb_strlen($row['alt_title'], 'UTF-8') ?> chars ✓</div>
            <?php else: ?>
              <span style="color:#bbb;font-size:.8rem;">—</span>
            <?php endif ?>
          </td>
          <td><span class="badge <?= $ts['class'] ?>"><?= $ts['label'] ?></span><br><small style="color:#aaa"><?= $rt['attempts'] ?>× API</small></td>
          <td class="text-cell"><span><?= htmlspecialchars($row['original_desc']) ?></span></td>
          <td class="text-cell">
            <span><?= htmlspecialchars($rd['text']) ?></span>
            <?php if ($rd['text'] !== ''): ?>
            <div class="len <?= $rd['valid'] ? 'valid' : 'invalid' ?>">
              <?= $rd['len'] ?> chars <?= $rd['valid'] ? '✓' : '✗' ?>
            </div>
            <?php endif ?>
          </td>
          <td><span class="badge <?= $ds['class'] ?>"><?= $ds['label'] ?></span><br><small style="color:#aaa"><?= $rd['attempts'] ?>× API</small></td>
        </tr>
      <?php endforeach ?>
      </tbody>
    </table>
    </div>
  </div>

<?php endif ?>
</div>
</body>
</html>
