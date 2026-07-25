<?php

declare(strict_types=1);

header('Content-Type: application/json');

$job_id = preg_replace('/[^a-f0-9]/', '', $_GET['job'] ?? '');

if (!$job_id) {
    echo json_encode(['status' => 'not_found']);
    exit;
}

$job_file = __DIR__ . '/jobs/' . $job_id . '.json';

if (!file_exists($job_file)) {
    echo json_encode(['status' => 'not_found']);
    exit;
}

$job = json_decode(file_get_contents($job_file), true);

if (!$job) {
    echo json_encode(['status' => 'error', 'error' => 'Invalid job data.']);
    exit;
}

echo json_encode($job);
