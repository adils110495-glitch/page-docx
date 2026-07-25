<?php

declare(strict_types=1);

const OPENAI_MODEL       = 'gpt-4.1-mini';
const OPENAI_BATCH_SIZE  = 5;   // items per prompt
const OPENAI_CONCURRENCY = 3;   // parallel curl handles
const OPENAI_TIMEOUT     = 45;  // seconds per request

// ── Single rewrite (used for alt-title generation) ─────────────────────────

function openai_rewrite(
    string $original_english,
    string $target_language,
    string $field_type,
    string $deepl_ref = '',
    string $exclude   = '',
    string $feedback  = ''
): ?string {
    $results = openai_run_batches_concurrent([[
        'lang_name'  => $target_language,
        'field_type' => $field_type,
        'items'      => [[
            'original'  => $original_english,
            'deepl_ref' => $deepl_ref,
            'exclude'   => $exclude,
            'feedback'  => $feedback,
        ]],
    ]]);

    return $results[0][0] ?? null;
}

// ── Build a correction prompt: show the model its own invalid output ─────────
// Items must have 'current_text' field (the invalid AI output to be edited).

function openai_build_correction_prompt(array $batch): string
{
    $lang_name  = $batch['lang_name'];
    $field_type = $batch['field_type'];
    $items      = $batch['items'];
    [$min, $max] = $field_type === 'title' ? [40, 46] : [150, 155];
    $mid        = (int)(($min + $max) / 2);
    $count      = count($items);

    $lines = '';
    foreach ($items as $i => $item) {
        $n        = $i + 1;
        $curr     = $item['current_text'];
        $curr_len = mb_strlen($curr, 'UTF-8');
        if ($curr_len > $max) {
            $delta  = $curr_len - $max;
            $action = "Remove at least {$delta} chars of words (e.g. drop one phrase)";
        } else {
            $delta  = $min - $curr_len;
            $action = "Add at least {$delta} chars of words (e.g. expand a detail)";
        }
        $lines .= "[{$n}] Current ({$curr_len} chars — outside {$min}–{$max}): \"{$curr}\"\n     → {$action}\n";
    }

    return <<<PROMPT
You are an SEO editor. Fix the character count of ALL {$count} {$lang_name} {$field_type}(s) below.

Valid range: {$min}–{$max} characters (target {$mid}). Count every char including spaces.
DO NOT retranslate from English. EDIT the current text only.

{$lines}
Rules:
- Change only what is needed to hit {$min}–{$max} chars
- Keep {$lang_name}, same meaning, SEO quality
- Count your result — if still outside {$min}–{$max}, edit again before outputting

OUTPUT: Return ONLY a JSON array of exactly {$count} strings. No markdown, no explanation.
PROMPT;
}

// ── Build prompt for a batch of items (same language + field_type) ──────────

function openai_build_batch_prompt(array $batch): string
{
    $lang_name  = $batch['lang_name'];
    $field_type = $batch['field_type'];
    $items      = $batch['items'];
    [$min, $max] = $field_type === 'title' ? [40, 46] : [150, 155];
    $mid        = (int)(($min + $max) / 2);
    $count      = count($items);

    $lines = '';
    foreach ($items as $i => $item) {
        $n   = $i + 1;
        $ref = ($item['deepl_ref'] ?? '') !== '' ? "\n   Reference: \"{$item['deepl_ref']}\"" : '';
        $fb  = ($item['feedback']  ?? '') !== '' ? "\n   ⚠ RETRY: {$item['feedback']}." : '';
        $ex  = ($item['exclude']   ?? '') !== '' ? "\n   Do NOT reuse: \"{$item['exclude']}\"" : '';
        $lines .= "[{$n}] \"{$item['original']}\"{$ref}{$fb}{$ex}\n";
    }

    return <<<PROMPT
You are an SEO localisation expert. Translate/rewrite ALL {$count} items into {$lang_name}.

══════════════════════════════════════════════
RULE 1 — LENGTH (HIGHEST PRIORITY, NON-NEGOTIABLE)
  Target: {$mid} characters. Valid range: {$min}–{$max} characters.
  Count EVERY character including spaces and punctuation.
  An output outside {$min}–{$max} chars is WRONG — rewrite until it fits.
  Shorten by removing words. Lengthen by adding words. Do not pad with filler.

RULE 2 — LANGUAGE
  Every word must be in {$lang_name}. No English.

RULE 3 — TOPIC
  Same subject and intent as the English original. Rephrase freely to meet length.

RULE 4 — SEO
  Natural, fluent, search-optimised {$lang_name}.
══════════════════════════════════════════════

Items ({$field_type}s):
{$lines}
SELF-CHECK (do this for EACH item before writing):
  1. Count characters in your output.
  2. Is the count between {$min} and {$max}? If NO → rewrite.
  3. Only proceed to JSON once ALL items pass.

OUTPUT: Return ONLY a JSON array of exactly {$count} strings. No markdown, no explanation.
PROMPT;
}

// ── Build a curl handle for one batch (does NOT add to multi) ───────────────

function openai_make_handle(array $batch, string $api_key): \CurlHandle|false
{
    $count   = count($batch['items']);
    $prompt  = ($batch['correction'] ?? false)
        ? openai_build_correction_prompt($batch)
        : openai_build_batch_prompt($batch);

    $payload = json_encode([
        'model'       => OPENAI_MODEL,
        'messages'    => [
            [
                'role'    => 'system',
                'content' => 'Output ONLY a JSON array of strings. No explanation, no markdown.',
            ],
            [
                'role'    => 'user',
                'content' => $prompt,
            ],
        ],
        'temperature' => 0.3,
        'max_tokens'  => max(300, 120 * $count),
    ]);

    $ch = curl_init('https://api.openai.com/v1/chat/completions');
    curl_setopt_array($ch, [
        CURLOPT_POST           => true,
        CURLOPT_POSTFIELDS     => $payload,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HTTPHEADER     => [
            'Content-Type: application/json',
            'Authorization: Bearer ' . $api_key,
        ],
        CURLOPT_TIMEOUT => OPENAI_TIMEOUT,
    ]);
    return $ch;
}

// ── Parse a raw API response into an array of strings ──────────────────────

function openai_parse_response(string $response, int $expected_count): array
{
    $data    = json_decode($response, true);
    $content = $data['choices'][0]['message']['content'] ?? null;
    if (!$content) return array_fill(0, $expected_count, null);

    // Strip optional markdown fences
    $content = trim(preg_replace('/^```(?:json)?\s*|\s*```$/s', '', trim($content)));
    $arr     = json_decode($content, true);

    if (!is_array($arr) || count($arr) !== $expected_count) {
        return array_fill(0, $expected_count, null);
    }
    return array_map(fn($r) => is_string($r) ? trim($r) : null, $arr);
}

// ── Run batches concurrently via curl_multi ─────────────────────────────────
// $batches : array of batch descriptors (lang_name, field_type, items[])
// Returns  : array indexed same as $batches, each value = array of ?string

function openai_run_batches_concurrent(array $batches): array
{
    $api_key = getenv('OPENAI_API_KEY');
    if (!$api_key) {
        return array_map(fn($b) => array_fill(0, count($b['items']), null), $batches);
    }

    $results   = array_fill(0, count($batches), []);
    $remaining = array_keys($batches);
    $active    = [];   // curl_handle_id => batch_index
    $mh        = curl_multi_init();

    $seed = function () use (&$remaining, $batches, $api_key, $mh, &$active): void {
        while (!empty($remaining) && count($active) < OPENAI_CONCURRENCY) {
            $idx = array_shift($remaining);
            $ch  = openai_make_handle($batches[$idx], $api_key);
            curl_multi_add_handle($mh, $ch);
            $active[(int)$ch] = $idx;
        }
    };

    $seed();

    do {
        curl_multi_exec($mh, $still_running);
        if ($still_running > 0) curl_multi_select($mh, 0.5);

        while ($info = curl_multi_info_read($mh)) {
            if ($info['msg'] !== CURLMSG_DONE) continue;

            $ch        = $info['handle'];
            $batch_idx = $active[(int)$ch] ?? null;

            if ($batch_idx !== null) {
                $raw   = curl_multi_getcontent($ch);
                $count = count($batches[$batch_idx]['items']);
                $results[$batch_idx] = $raw
                    ? openai_parse_response($raw, $count)
                    : array_fill(0, $count, null);
                unset($active[(int)$ch]);
            }

            curl_multi_remove_handle($mh, $ch);
            unset($ch); // PHP 8: curl_close is a no-op; unset releases the handle
            $seed(); // may add new handles — $active updated, loop must continue
        }
    } while ($still_running > 0 || !empty($remaining) || !empty($active));

    curl_multi_close($mh);
    return $results;
}

// ── validate_meaning kept for compatibility (Gemini-based) ─────────────────

function validate_meaning(string $original_english, string $rewritten, string $target_language): bool
{
    $api_key = getenv('GEMINI_API_KEY');
    if (!$api_key) return true;

    $prompt = <<<PROMPT
Original (English): {$original_english}
Rewritten ({$target_language}): {$rewritten}
Does the rewritten text cover the same main topic and intent as the original (even if phrased differently for SEO)?
Answer ONLY: YES or NO
PROMPT;

    $url     = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' . $api_key;
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
        CURLOPT_TIMEOUT        => 10,
    ]);
    $response = curl_exec($ch);
    unset($ch);

    if (!$response) return true;
    $data   = json_decode($response, true);
    $answer = strtoupper(trim($data['candidates'][0]['content']['parts'][0]['text'] ?? 'YES'));
    return str_starts_with($answer, 'YES');
}
