<?php

declare(strict_types=1);

function init_log(string $log_file): void
{
    $dir = dirname($log_file);
    if (!is_dir($dir)) mkdir($dir, 0755, true);

    if (!file_exists($log_file)) {
        file_put_contents(
            $log_file,
            "timestamp | original | deepl | gemini | groq | final | attempts | status\n"
        );
    }
}

function log_row(array $data, string $log_file): void
{
    $line = implode(' | ', [
        date('Y-m-d H:i:s'),
        $data['original']  ?? '-',
        $data['deepl']     ?? '-',
        $data['gemini']    ?? '-',
        $data['groq']      ?? '-',
        $data['final']     ?? '-',
        (string) ($data['attempts'] ?? 0),
        $data['status']    ?? '-',
    ]);

    file_put_contents($log_file, $line . PHP_EOL, FILE_APPEND | LOCK_EX);
}
