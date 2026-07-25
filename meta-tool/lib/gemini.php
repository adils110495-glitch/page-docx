<?php

declare(strict_types=1);

function gemini_rewrite(
    string $original_english,
    string $target_language,
    string $field_type,
    string $deepl_ref = '',
    string $exclude   = '',   // when set, the output must differ from this text
    string $feedback  = ''    // e.g. "previous was 53 chars, too long by 7"
): ?string {
    $api_key = getenv('GEMINI_API_KEY');
    if (!$api_key) return null;

    [$min, $max] = $field_type === 'title' ? [40, 46] : [150, 155];

    $ref_block      = $deepl_ref !== ''
        ? "\nReference translation (vocabulary only):\n{$deepl_ref}\n"
        : '';
    $exclude_block  = $exclude !== ''
        ? "\nDo NOT reproduce this exact phrasing (generate different wording):\n{$exclude}\n"
        : '';
    $feedback_block = $feedback !== ''
        ? "\nPREVIOUS ATTEMPT FAILED: {$feedback} — adjust your output accordingly.\n"
        : '';

    $prompt = <<<PROMPT
You are an SEO localisation expert. You MUST follow the checklist below without exception.

═══════════════════════════════════════
SEO META {$field_type} CHECKLIST
═══════════════════════════════════════

CHECKLIST ITEM 1 — LANGUAGE (MANDATORY)
✔ Output must be written entirely in {$target_language}.
✔ Any English in your output = automatic failure. Rewrite in {$target_language}.

CHECKLIST ITEM 2 — CHARACTER COUNT (MANDATORY, HIGHEST PRIORITY)
✔ Total character count MUST be between {$min} and {$max} (inclusive).
✔ Count every character: letters, spaces, punctuation, numbers, symbols.
✔ If your draft has more than {$max} chars → remove words until it fits.
✔ If your draft has fewer than {$min} chars → add words until it fits.
✔ Do NOT output until character count is confirmed within {$min}–{$max}.
{$feedback_block}
CHECKLIST ITEM 3 — MEANING (REQUIRED)
✔ Cover the same main topic and intent as the Original English below.
✔ Keep key entities (brand name, numbers, locations) where length permits.
✔ Rephrase, condense, or adapt freely — word-for-word translation is NOT required.
✔ The character count limit overrides preserving minor details.

CHECKLIST ITEM 4 — SEO QUALITY (REQUIRED)
✔ Natural, fluent language that reads well in {$target_language}.
✔ Search-optimised: use keywords a user would type to find this content.

═══════════════════════════════════════
Original English:
{$original_english}
{$ref_block}{$exclude_block}
SELF-CHECK BEFORE OUTPUT:
→ Is it in {$target_language}? (yes/no — if no, rewrite)
→ Character count between {$min} and {$max}? (yes/no — if no, rewrite)
→ Covers the main topic? (yes/no — if no, rewrite)
Only output when all three answers are YES.

OUTPUT: Return ONLY the final {$target_language} text. No counts, no labels, no explanation.
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

Does the rewritten text cover the same main topic and intent as the original English (even if phrased differently or condensed for SEO)?
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

    if (!$response) return true;

    $data   = json_decode($response, true);
    $answer = strtoupper(trim($data['candidates'][0]['content']['parts'][0]['text'] ?? 'YES'));

    return str_starts_with($answer, 'YES');
}
