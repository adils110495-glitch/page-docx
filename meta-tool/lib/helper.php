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

function normalize_language_code(string $code): string
{
    $code = strtoupper(trim($code));
    $aliases = ['ZH-CN' => 'ZH', 'ZH-TW' => 'ZH'];
    return $aliases[$code] ?? $code;
}

function get_language_name(string $code): string
{
    $map = [

        // DeepL-supported languages
        'AR' => 'Arabic',
        'BG' => 'Bulgarian',
        'CS' => 'Czech',
        'DA' => 'Danish',
        'DE' => 'German',
        'EL' => 'Greek',
        'EN' => 'English',
        'ES' => 'Spanish',
        'ET' => 'Estonian',
        'FI' => 'Finnish',
        'FR' => 'French',
        'HU' => 'Hungarian',
        'ID' => 'Indonesian',
        'IT' => 'Italian',
        'JA' => 'Japanese',
        'KO' => 'Korean',
        'LT' => 'Lithuanian',
        'LV' => 'Latvian',
        'NB' => 'Norwegian (Bokmål)',
        'NL' => 'Dutch',
        'PL' => 'Polish',
        'PT' => 'Portuguese',
        'PT-BR' => 'Portuguese (Brazilian)',
        'RO' => 'Romanian',
        'RU' => 'Russian',
        'SK' => 'Slovak',
        'SL' => 'Slovenian',
        'SV' => 'Swedish',
        'TR' => 'Turkish',
        'UK' => 'Ukrainian',
        'ZH' => 'Chinese (Simplified)',

        // Additional languages (AI / Extended)
        'ACE' => 'Acehnese',
        'AF' => 'Afrikaans',
        'SQ' => 'Albanian',
        'AN' => 'Aragonese',
        'HY' => 'Armenian',
        'AS' => 'Assamese',
        'AY' => 'Aymara',
        'AZ' => 'Azerbaijani',
        'BA' => 'Bashkir',
        'EU' => 'Basque',
        'BE' => 'Belarusian',
        'BN' => 'Bengali',
        'BHO' => 'Bhojpuri',
        'BS' => 'Bosnian',
        'BR' => 'Breton',
        'BG' => 'Bulgarian',
        'MY' => 'Burmese',
        'YUE' => 'Cantonese',
        'CA' => 'Catalan',
        'CEB' => 'Cebuano',
        'ZH-HANT' => 'Chinese (Traditional)',
        'HR' => 'Croatian',
        'PRS' => 'Dari',
        'NL' => 'Dutch',
        'EN-US' => 'English (American)',
        'EN-GB' => 'English (British)',
        'EO' => 'Esperanto',
        'ET' => 'Estonian',
        'FA' => 'Persian',
        'GL' => 'Galician',
        'KA' => 'Georgian',
        'DE' => 'German',
        'EL' => 'Greek',
        'GN' => 'Guarani',
        'GU' => 'Gujarati',
        'HT' => 'Haitian Creole',
        'HA' => 'Hausa',
        'HE' => 'Hebrew',
        'HI' => 'Hindi',
        'HU' => 'Hungarian',
        'IS' => 'Icelandic',
        'IG' => 'Igbo',
        'ID' => 'Indonesian',
        'GA' => 'Irish',
        'IT' => 'Italian',
        'JA' => 'Japanese',
        'JV' => 'Javanese',
        'PAM' => 'Kapampangan',
        'KK' => 'Kazakh',
        'KO' => 'Korean',
        'GOM' => 'Konkani',
        'KMR' => 'Kurdish (Kurmanji)',
        'CKB' => 'Kurdish (Sorani)',
        'KY' => 'Kyrgyz',
        'LA' => 'Latin',
        'LV' => 'Latvian',
        'LN' => 'Lingala',
        'LT' => 'Lithuanian',
        'LMO' => 'Lombard',
        'LB' => 'Luxembourgish',
        'MK' => 'Macedonian',
        'MAI' => 'Maithili',
        'MG' => 'Malagasy',
        'MS' => 'Malay',
        'ML' => 'Malayalam',
        'MT' => 'Maltese',
        'MI' => 'Maori',
        'MR' => 'Marathi',
        'MN' => 'Mongolian',
        'NE' => 'Nepali',
        'NO' => 'Norwegian',
        'OC' => 'Occitan',
        'OM' => 'Oromo',
        'PAG' => 'Pangasinan',
        'PS' => 'Pashto',
        'PL' => 'Polish',
        'PT-PT' => 'Portuguese (Portugal)',
        'PA' => 'Punjabi',
        'QU' => 'Quechua',
        'RO' => 'Romanian',
        'RU' => 'Russian',
        'SA' => 'Sanskrit',
        'SR' => 'Serbian',
        'ST' => 'Sesotho',
        'SCN' => 'Sicilian',
        'SK' => 'Slovak',
        'SL' => 'Slovenian',
        'ES-419' => 'Spanish (Latin American)',
        'SU' => 'Sundanese',
        'SW' => 'Swahili',
        'SV' => 'Swedish',
        'TL' => 'Tagalog',
        'TG' => 'Tajik',
        'TA' => 'Tamil',
        'TT' => 'Tatar',
        'TE' => 'Telugu',
        'TH' => 'Thai',
        'TS' => 'Tsonga',
        'TN' => 'Tswana',
        'TR' => 'Turkish',
        'UK' => 'Ukrainian',
        'VI' => 'Vietnamese',
        'CY' => 'Welsh',
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
