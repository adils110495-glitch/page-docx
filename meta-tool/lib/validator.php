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

    $trimmed = mb_substr($text, 0, $max, 'UTF-8');
    $last_space = mb_strrpos($trimmed, ' ', 0, 'UTF-8');

    if ($last_space !== false && $last_space > 0) {
        return mb_substr($trimmed, 0, $last_space, 'UTF-8');
    }

    return $trimmed;
}
