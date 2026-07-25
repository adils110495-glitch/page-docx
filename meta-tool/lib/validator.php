<?php

declare(strict_types=1);

function validate_title(string $text): bool
{
    $len = mb_strlen($text, 'UTF-8');
    return $len >= 40 && $len <= 46;
}

function validate_description(string $text): bool
{
    $len = mb_strlen($text, 'UTF-8');
    return $len >= 150 && $len <= 155;
}

function distance_to_range(int $len, int $min, int $max): int
{
    if ($len < $min) return $min - $len;
    if ($len > $max) return $len - $max;
    return 0;
}

function hard_trim(string $text, int $max): string
{
    if (mb_strlen($text, 'UTF-8') <= $max) return $text;

    $trimmed    = mb_substr($text, 0, $max, 'UTF-8');
    $last_space = mb_strrpos($trimmed, ' ', 0, 'UTF-8');

    return ($last_space !== false && $last_space > 0)
        ? mb_substr($trimmed, 0, $last_space, 'UTF-8')
        : $trimmed;
}

// Trim to [$min,$max] by finding the last word boundary within that range.
// Falls back to longest word-boundary fit ≤ $max if no in-range boundary exists.
function smart_trim(string $text, int $min, int $max): string
{
    if (mb_strlen($text, 'UTF-8') <= $max) return $text;

    $words  = preg_split('/\s+/u', trim($text));
    $acc    = '';
    $best   = '';   // last position that landed in [$min, $max]
    $last   = '';   // last position ≤ $max (regardless of min)

    foreach ($words as $word) {
        $candidate = $acc === '' ? $word : $acc . ' ' . $word;
        $clen      = mb_strlen($candidate, 'UTF-8');
        if ($clen > $max) break;
        $acc  = $candidate;
        $last = $candidate;
        if ($clen >= $min) $best = $candidate;
    }

    if ($best !== '') return $best;   // exact range hit
    if ($last !== '') return $last;   // best word-boundary fit ≤ max

    return mb_substr($text, 0, $max, 'UTF-8'); // hard cut (single mega-word edge case)
}
