<?php

declare(strict_types=1);

function groq_rewrite(
    string $original_english,
    string $target_language,
    string $field_type,
    string $deepl_ref = '',
    string $exclude   = ''    // when set, the output must differ from this text
): ?string {
    $api_key = getenv('GROQ_API_KEY');
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

    // Use higher temperature when an alternate (non-duplicate) phrasing is required
    $temperature = $exclude !== '' ? 0.7 : 0.3;

    $payload = json_encode([
        'model'    => 'llama-3.3-70b-versatile',
        'messages' => [
            [
                'role'    => 'system',
                'content' => 'You are an SEO localisation expert. Output ONLY the requested text. No quotes, no labels, no explanation.',
            ],
            ['role' => 'user', 'content' => $prompt],
        ],
        'temperature' => $temperature,
        'max_tokens'  => 300,
    ]);

    $ch = curl_init('https://api.groq.com/openai/v1/chat/completions');
    curl_setopt_array($ch, [
        CURLOPT_POST           => true,
        CURLOPT_POSTFIELDS     => $payload,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HTTPHEADER     => [
            'Content-Type: application/json',
            'Authorization: Bearer ' . $api_key,
        ],
        CURLOPT_TIMEOUT => 30,
    ]);

    $response = curl_exec($ch);
    curl_close($ch);

    if (!$response) return null;

    $data = json_decode($response, true);
    $text = $data['choices'][0]['message']['content'] ?? null;

    return $text ? trim(trim($text, "\"'\n\r")) : null;
}
