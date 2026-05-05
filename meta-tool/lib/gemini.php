<?php

declare(strict_types=1);

function gemini_rewrite(
    string $original_english,
    string $target_language,
    string $field_type,
    string $deepl_ref = '',
    string $exclude   = ''    // when set, the output must differ from this text
): ?string {
    $api_key = getenv('GEMINI_API_KEY');
    if (!$api_key) return null;

    [$min, $max] = $field_type === 'title' ? [40, 46] : [150, 155];

    $ref_block     = $deepl_ref !== ''
        ? "\nDeepL translation (vocabulary reference only — do NOT copy if it changes meaning):\n{$deepl_ref}\n"
        : '';
    $exclude_block = $exclude !== ''
        ? "\nDo NOT reproduce this exact phrasing (generate a different wording):\n{$exclude}\n"
        : '';

    $prompt = <<<PROMPT
You are an SEO localisation expert. Write a {$field_type} in {$target_language}.

PRIORITY 1 — EXACT MEANING (NON-NEGOTIABLE):
- Express the IDENTICAL meaning as the Original English below
- Same topic, intent, key facts, numbers, and entities (brand names, prices, locations)
- Do NOT add information not present in the original
- Do NOT remove any key concept from the original
- Do NOT change informational text to promotional or vice versa

PRIORITY 2 — CHARACTER COUNT (MANDATORY):
- The result MUST be between {$min} and {$max} characters — count every character carefully
- Expand or condense phrasing to hit the range; never sacrifice meaning to do so

PRIORITY 3 — SEO QUALITY:
- Natural, fluent, and search-optimised in {$target_language}

Original English:
{$original_english}
{$ref_block}{$exclude_block}
OUTPUT: Return ONLY the final text. No quotes, labels, or explanation.
PROMPT;

    $url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' . $api_key;

    // Use higher temperature when an alternate (non-duplicate) phrasing is required
    $temperature = $exclude !== '' ? 0.7 : 0.3;

    $payload = json_encode([
        'contents'         => [['parts' => [['text' => $prompt]]]],
        'generationConfig' => ['temperature' => $temperature, 'maxOutputTokens' => 300],
    ]);

    $ch = curl_init($url);
    curl_setopt_array($ch, [
        CURLOPT_POST           => true,
        CURLOPT_POSTFIELDS     => $payload,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HTTPHEADER     => ['Content-Type: application/json'],
        CURLOPT_TIMEOUT        => 30,
    ]);

    $response = curl_exec($ch);
    curl_close($ch);

    if (!$response) return null;

    $data = json_decode($response, true);
    $text = $data['candidates'][0]['content']['parts'][0]['text'] ?? null;

    return $text ? trim(trim($text, "\"'\n\r")) : null;
}

/**
 * Asks Gemini to compare the original English and the rewritten text.
 * Returns true (meaning preserved) or false (meaning changed).
 * On API failure returns true so the pipeline keeps moving.
 */
function validate_meaning(string $original_english, string $rewritten, string $target_language): bool
{
    $api_key = getenv('GEMINI_API_KEY');
    if (!$api_key) return true;

    $prompt = <<<PROMPT
Original (English):
{$original_english}

Rewritten ({$target_language}):
{$rewritten}

Does the rewritten text express the EXACT same meaning, intent, and key information as the original English?
Answer ONLY: YES or NO
PROMPT;

    $url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' . $api_key;

    $payload = json_encode([
        'contents'         => [['parts' => [['text' => $prompt]]]],
        'generationConfig' => ['temperature' => 0.1, 'maxOutputTokens' => 5],
    ]);

    $ch = curl_init($url);
    curl_setopt_array($ch, [
        CURLOPT_POST           => true,
        CURLOPT_POSTFIELDS     => $payload,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HTTPHEADER     => ['Content-Type: application/json'],
        CURLOPT_TIMEOUT        => 15,
    ]);

    $response = curl_exec($ch);
    curl_close($ch);

    if (!$response) return true;

    $data   = json_decode($response, true);
    $answer = strtoupper(trim($data['candidates'][0]['content']['parts'][0]['text'] ?? 'YES'));

    return str_starts_with($answer, 'YES');
}
