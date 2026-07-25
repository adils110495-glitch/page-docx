<?php

declare(strict_types=1);

function groq_rewrite(
    string $original_english,
    string $target_language,
    string $field_type,
    string $deepl_ref     = '',
    string $exclude       = '',
    string $feedback      = '',
    string $current_text  = ''  // if set, use correction mode (edit this text instead of generating)
): ?string {
    $api_key = getenv('GROQ_API_KEY');
    if (!$api_key) return null;

    [$min, $max] = $field_type === 'title' ? [40, 46] : [150, 155];
    $mid = (int)(($min + $max) / 2);

    $ref_block     = $deepl_ref !== ''
        ? "\nReference translation (vocabulary only):\n{$deepl_ref}\n"
        : '';
    $exclude_block = $exclude !== ''
        ? "\nDo NOT reuse this phrasing — generate different wording:\n{$exclude}\n"
        : '';

    if ($current_text !== '') {
        // ── Correction mode: model edits its own invalid output ──────────────
        $curr_len = mb_strlen($current_text, 'UTF-8');
        if ($curr_len > $max) {
            $delta  = $curr_len - $max;
            $action = "Remove at least {$delta} characters worth of words";
        } else {
            $delta  = $min - $curr_len;
            $action = "Add at least {$delta} characters worth of words";
        }
        $prompt = <<<PROMPT
You are an SEO editor. Fix the character count of this {$target_language} {$field_type}.

Current text ({$curr_len} chars — outside {$min}–{$max}):
"{$current_text}"

Action: {$action}. Target: {$mid} chars (valid: {$min}–{$max}).

Rules:
- EDIT the current text — do NOT retranslate from English
- Change only what is needed to hit {$min}–{$max} chars
- Keep {$target_language}, same meaning, SEO quality
- Count your result before outputting. Must be {$min}–{$max} chars.

OUTPUT: Return ONLY the edited {$target_language} text. No counts, no labels, no explanation.
PROMPT;
    } else {
        // ── Generation mode ────────────────────────────────────────────────
        $feedback_block = $feedback !== ''
            ? "\n⚠ PREVIOUS ATTEMPT FAILED: {$feedback}. You MUST fix this.\n"
            : '';
        $prompt = <<<PROMPT
You are an SEO localisation expert. Write a {$field_type} in {$target_language}.

══════════════════════════════════════════════
RULE 1 — LENGTH (HIGHEST PRIORITY, NON-NEGOTIABLE)
  Target: {$mid} characters. Valid range: {$min}–{$max} characters.
  Count EVERY character including spaces and punctuation.
  Outside {$min}–{$max} = WRONG. Rewrite until it fits.
  Too long → remove words. Too short → add words.
{$feedback_block}
RULE 2 — LANGUAGE
  Every word must be in {$target_language}. No English words at all.

RULE 3 — TOPIC
  Same subject and intent as the English original. Rephrase freely to meet length.
  Key entities (brand, numbers) kept where length allows. Length overrides details.

RULE 4 — SEO
  Natural, fluent, search-optimised {$target_language}.
══════════════════════════════════════════════

Original English:
{$original_english}
{$ref_block}{$exclude_block}
BEFORE OUTPUTTING:
  Count every character in your draft.
  If count < {$min}: add words until ≥ {$min}.
  If count > {$max}: remove words until ≤ {$max}.
  Target {$mid} chars. Only output when {$min} ≤ count ≤ {$max}.

OUTPUT: Return ONLY the final {$target_language} text. No counts, no labels, no explanation.
PROMPT;
    }

    // Use higher temperature when an alternate (non-duplicate) phrasing is required
    $temperature = $exclude !== '' ? 0.7 : 0.3;

    $payload = json_encode([
        'model'    => 'llama-3.3-70b-versatile',
        'messages' => [
            [
                'role'    => 'system',
                'content' => 'You are an SEO localisation expert. You follow a strict checklist: (1) correct language, (2) character count within the specified range — this is the highest priority, (3) same topic as the original, (4) SEO-optimised. Output ONLY the final text. No counts, no labels, no explanation.',
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
        CURLOPT_TIMEOUT => 12,
    ]);

    $response = curl_exec($ch);

    if (!$response) return null;

    $data = json_decode($response, true);
    $text = $data['choices'][0]['message']['content'] ?? null;

    return $text ? trim(trim($text, "\"'\n\r")) : null;
}
