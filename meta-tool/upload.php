<?php

declare(strict_types=1);

session_start();

require_once __DIR__ . '/lib/helper.php';

load_env(dirname(__DIR__) . '/.env');

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    header('Location: index.php');
    exit;
}

// ── Validate language ──────────────────────────────────────────────────────
$target_lang = strtoupper(trim($_POST['target_lang'] ?? ''));
if (!$target_lang) {
    $_SESSION['error'] = 'Please select a target language.';
    header('Location: index.php');
    exit;
}

// ── Validate upload ────────────────────────────────────────────────────────
$upload_error = $_FILES['csv_file']['error'] ?? UPLOAD_ERR_NO_FILE;
if ($upload_error !== UPLOAD_ERR_OK) {
    $messages = [
        UPLOAD_ERR_INI_SIZE   => 'File exceeds server upload limit.',
        UPLOAD_ERR_FORM_SIZE  => 'File exceeds form size limit.',
        UPLOAD_ERR_NO_TMP_DIR => 'Server temporary directory missing.',
        UPLOAD_ERR_CANT_WRITE => 'Failed to write file to disk.',
        UPLOAD_ERR_NO_FILE    => 'No file was uploaded.',
    ];
    $_SESSION['error'] = $messages[$upload_error] ?? 'File upload failed.';
    header('Location: index.php');
    exit;
}

$file = $_FILES['csv_file'];

$ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
if ($ext !== 'csv') {
    $_SESSION['error'] = 'Only .csv files are accepted.';
    header('Location: index.php');
    exit;
}

if ($file['size'] > 10 * 1024 * 1024) {
    $_SESSION['error'] = 'File is too large. Maximum allowed size is 10 MB.';
    header('Location: index.php');
    exit;
}

$finfo   = new finfo(FILEINFO_MIME_TYPE);
$mime    = $finfo->file($file['tmp_name']);
$allowed = ['text/plain', 'text/csv', 'application/csv', 'application/octet-stream'];
if (!in_array($mime, $allowed, true)) {
    $_SESSION['error'] = 'Invalid file type. Please upload a plain CSV file.';
    header('Location: index.php');
    exit;
}

// ── Save uploaded file ─────────────────────────────────────────────────────
$tmp_dir = __DIR__ . '/tmp/';
if (!is_dir($tmp_dir)) mkdir($tmp_dir, 0755, true);

$token      = bin2hex(random_bytes(16));
$input_path = $tmp_dir . $token . '.csv';

if (!move_uploaded_file($file['tmp_name'], $input_path)) {
    $_SESSION['error'] = 'Could not save the uploaded file. Please try again.';
    header('Location: index.php');
    exit;
}

normalize_csv_encoding($input_path);

// ── Create job file ────────────────────────────────────────────────────────
$jobs_dir = __DIR__ . '/jobs';
if (!is_dir($jobs_dir)) mkdir($jobs_dir, 0755, true);

$job_id   = bin2hex(random_bytes(16));
$job_file = "{$jobs_dir}/{$job_id}.json";

$job = [
    'status'       => 'processing',   // mark immediately — no race with a daemon
    'input_file'   => $input_path,
    'target_lang'  => $target_lang,
    'output_file'  => null,
    'result'       => null,
    'error'        => null,
    'created_at'   => time(),
    'started_at'   => time(),
    'completed_at' => null,
];
file_put_contents($job_file, json_encode($job), LOCK_EX);

// ── Send redirect to browser NOW, then process in the same PHP process ────
//
// On Hostinger (PHP-FPM) fastcgi_finish_request() closes the client
// connection while PHP continues running — no background daemon needed.
// The connection-close fallback covers other PHP-FPM setups.

header('Location: process.php?job=' . urlencode($job_id));
header('Connection: close');

if (function_exists('fastcgi_finish_request')) {
    fastcgi_finish_request();
} else {
    // Flush & close the socket manually
    if (ob_get_level()) {
        ob_end_clean();
    }
    header('Content-Length: 0');
    flush();
}

// ── Browser is gone — process the job ─────────────────────────────────────
ignore_user_abort(true);
set_time_limit(0);
ini_set('memory_limit', '256M');

require_once __DIR__ . '/lib/deepl.php';
require_once __DIR__ . '/lib/openai.php';
require_once __DIR__ . '/lib/groq.php';
require_once __DIR__ . '/lib/validator.php';
require_once __DIR__ . '/lib/logger.php';
require_once __DIR__ . '/lib/pipeline.php';

$output_dir = __DIR__ . '/output';
if (!is_dir($output_dir)) mkdir($output_dir, 0755, true);

$logs_dir = __DIR__ . '/logs';
if (!is_dir($logs_dir)) mkdir($logs_dir, 0755, true);

$output_file = "{$output_dir}/{$job_id}.csv";
$log_file    = "{$logs_dir}/{$job_id}.log";

$result = process_csv($input_path, $target_lang, $output_file, $log_file);

@unlink($input_path);

// Re-read job file in case anything external modified it
$job = json_decode(file_get_contents($job_file), true) ?: $job;

if (isset($result['error'])) {
    $job['status'] = 'error';
    $job['error']  = $result['error'];
} else {
    $job['status']       = 'done';
    $job['output_file']  = $output_file;
    $job['result']       = $result;
    $job['completed_at'] = time();
}

file_put_contents($job_file, json_encode($job), LOCK_EX);
