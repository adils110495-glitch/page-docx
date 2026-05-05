<?php

declare(strict_types=1);

function load_env(string $path): void
{
    if (!file_exists($path)) return;

    foreach (file($path, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES) as $line) {
        $line = trim($line);
        if ($line === '' || $line[0] === '#' || !str_contains($line, '=')) continue;

        [$key, $value] = explode('=', $line, 2);
        $key   = trim($key);
        $value = trim($value, " \t\n\r\0\x0B\"'");

        putenv("{$key}={$value}");
        $_ENV[$key] = $value;
    }
}

function get_language_name(string $code): string
{
    $map = [
        'FR'    => 'French',
        'DE'    => 'German',
        'ES'    => 'Spanish',
        'IT'    => 'Italian',
        'NL'    => 'Dutch',
        'PT'    => 'Portuguese',
        'PT-BR' => 'Portuguese (Brazilian)',
        'PL'    => 'Polish',
        'RU'    => 'Russian',
        'JA'    => 'Japanese',
        'ZH'    => 'Chinese (Simplified)',
        'SV'    => 'Swedish',
        'DA'    => 'Danish',
        'FI'    => 'Finnish',
        'TR'    => 'Turkish',
        'CS'    => 'Czech',
        'RO'    => 'Romanian',
    ];

    return $map[strtoupper($code)] ?? ucfirst(strtolower($code));
}

function clean_text(string $text): string
{
    return trim(preg_replace('/\s+/', ' ', $text));
}

function normalize_csv_encoding(string $path): void
{
    $raw = file_get_contents($path);
    if ($raw === false) return;

    // UTF-16 LE (FF FE) — most common Excel export on Windows
    if (str_starts_with($raw, "\xFF\xFE")) {
        $converted = mb_convert_encoding(substr($raw, 2), 'UTF-8', 'UTF-16LE');
        file_put_contents($path, $converted);
        return;
    }

    // UTF-16 BE (FE FF)
    if (str_starts_with($raw, "\xFE\xFF")) {
        $converted = mb_convert_encoding(substr($raw, 2), 'UTF-8', 'UTF-16BE');
        file_put_contents($path, $converted);
        return;
    }

    // UTF-8 with BOM (EF BB BF) — strip the BOM so fgetcsv doesn't corrupt the first header
    if (str_starts_with($raw, "\xEF\xBB\xBF")) {
        file_put_contents($path, substr($raw, 3));
        return;
    }

    // If it's not valid UTF-8, try to convert from Windows-1252 (Latin-1)
    if (!mb_check_encoding($raw, 'UTF-8')) {
        $converted = mb_convert_encoding($raw, 'UTF-8', 'Windows-1252');
        file_put_contents($path, $converted);
    }
}
