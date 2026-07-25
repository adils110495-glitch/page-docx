<?php

declare(strict_types=1);

// ── Helpers ────────────────────────────────────────────────────────────────

function make_feedback(string $text, int $min, int $max): string
{
    $len = mb_strlen($text, 'UTF-8');
    $mid = (int)(($min + $max) / 2);
    if ($len > $max) return "was {$len} chars ({$len - $max} too long) — remove words, target {$mid} chars";
    if ($len < $min) return "was {$len} chars ({$min - $len} too short) — add words, target {$mid} chars";
    return '';
}

function field_range(string $field_type): array
{
    return $field_type === 'title' ? [40, 46] : [150, 155];
}

function is_valid_field(string $text, string $field_type): bool
{
    [$min, $max] = field_range($field_type);
    $len = mb_strlen($text, 'UTF-8');
    return $len >= $min && $len <= $max;
}

// ── Batch all items through OpenAI, then Groq for failures ─────────────────
// $jobs : array of ['original'=>str, 'lang'=>str, 'lang_name'=>str,
//                   'deepl_ref'=>str, 'field_type'=>str]
// Returns: array indexed same as $jobs — ['text'=>str, 'status'=>str,
//           'attempts'=>int, 'len'=>int, 'valid'=>bool, 'deepl_ref'=>str]

function run_batch_pipeline(array $jobs): array
{
    if (empty($jobs)) return [];

    $results  = array_fill(0, count($jobs), null);
    $pending  = array_keys($jobs);  // job indices still needing valid output
    $attempts = array_fill(0, count($jobs), 0);
    $feedback = array_fill(0, count($jobs), '');
    $best     = array_fill(0, count($jobs), null); // closest result per job

    $track_best = function (int $ji, string $text) use (&$best, $jobs): void {
        [$min, $max] = field_range($jobs[$ji]['field_type']);
        $dist = distance_to_range(mb_strlen($text, 'UTF-8'), $min, $max);
        if ($best[$ji] === null || distance_to_range(mb_strlen($best[$ji], 'UTF-8'), $min, $max) > $dist) {
            $best[$ji] = $text;
        }
    };

    // Helper: run one OpenAI pass and return still-failing indices.
    // $correction=true uses the correction prompt (edit invalid output) instead of generation.
    $openai_pass = function (array $indices, bool $single_item, bool $correction = false) use (
        $jobs, &$results, &$attempts, &$feedback, &$best, $track_best
    ): array {
        if (empty($indices)) return [];

        $batches   = [];
        $batch_map = [];

        if ($single_item) {
            foreach ($indices as $ji) {
                $bidx = count($batches);
                $batch_map[$bidx] = [$ji];
                $item = [
                    'original'  => $jobs[$ji]['original'],
                    'deepl_ref' => $jobs[$ji]['deepl_ref'],
                    'feedback'  => $feedback[$ji],
                    'exclude'   => '',
                ];
                $use_correction = $correction && $best[$ji] !== null;
                if ($use_correction) {
                    $item['current_text'] = $best[$ji];
                }
                $batches[] = [
                    'lang_name'  => $jobs[$ji]['lang_name'],
                    'field_type' => $jobs[$ji]['field_type'],
                    'correction' => $use_correction,
                    'items'      => [$item],
                ];
            }
        } else {
            $groups = [];
            foreach ($indices as $ji) {
                $key = $jobs[$ji]['lang_name'] . '||' . $jobs[$ji]['field_type'];
                $groups[$key][] = $ji;
            }
            foreach ($groups as $group_jobs) {
                foreach (array_chunk($group_jobs, OPENAI_BATCH_SIZE) as $chunk) {
                    $first = $jobs[$chunk[0]];
                    $bidx  = count($batches);
                    $batch_map[$bidx] = $chunk;
                    $batches[] = [
                        'lang_name'  => $first['lang_name'],
                        'field_type' => $first['field_type'],
                        'correction' => false,
                        'items'      => array_map(fn($ji) => [
                            'original'  => $jobs[$ji]['original'],
                            'deepl_ref' => $jobs[$ji]['deepl_ref'],
                            'feedback'  => $feedback[$ji],
                            'exclude'   => '',
                        ], $chunk),
                    ];
                }
            }
        }

        $batch_results = openai_run_batches_concurrent($batches);
        $still_pending = [];

        foreach ($batch_map as $bidx => $chunk) {
            $texts = $batch_results[$bidx] ?? [];
            foreach ($chunk as $pos => $ji) {
                $attempts[$ji]++;
                $text = $texts[$pos] ?? null;
                if ($text) {
                    $text = clean_text($text);
                    $track_best($ji, $text);
                    $feedback[$ji] = make_feedback($text, ...field_range($jobs[$ji]['field_type']));
                    if (is_valid_field($text, $jobs[$ji]['field_type'])) {
                        $results[$ji] = $text;
                        continue;
                    }
                }
                $still_pending[] = $ji;
            }
        }

        return array_unique($still_pending);
    };

    // ── Pass 1 & 2: OpenAI batch generation ───────────────────────────────────
    $pending = $openai_pass($pending, false);
    $pending = $openai_pass($pending, false);

    // ── Pass 3: OpenAI correction — show model its own invalid output to edit ─
    $pending = $openai_pass($pending, true, true);

    // ── Pass 4 & 5: Groq correction (uses best text seen so far as input) ─────
    foreach ($pending as $ji) {
        $job = $jobs[$ji];
        [$min, $max] = field_range($job['field_type']);

        for ($i = 0; $i < 3; $i++) {
            $text = groq_rewrite(
                $job['original'], $job['lang_name'], $job['field_type'],
                $job['deepl_ref'], '', $feedback[$ji], $best[$ji] ?? ''
            );
            $attempts[$ji]++;

            if ($text) {
                $text = clean_text($text);
                $track_best($ji, $text);
                $feedback[$ji] = make_feedback($text, $min, $max);

                if (is_valid_field($text, $job['field_type'])) {
                    $results[$ji] = $text;
                    break;
                }
            }
        }
    }

    // ── Build final result objects ─────────────────────────────────────────
    $output = [];
    foreach ($jobs as $ji => $job) {
        [$min, $max] = field_range($job['field_type']);

        if ($results[$ji] !== null) {
            $final  = $results[$ji];
            $status = 'openai_ok';
        } else {
            // All passes failed to produce valid output — leave blank
            $final  = '';
            $status = 'failed';
        }

        $len = mb_strlen($final, 'UTF-8');
        $output[$ji] = [
            'text'      => $final,
            'status'    => $status,
            'attempts'  => $attempts[$ji],
            'len'       => $len,
            'valid'     => ($len >= $min && $len <= $max),
            'deepl_ref' => $job['deepl_ref'],
        ];
    }

    return $output;
}

// ── Alt title generation (batched, one call per language group) ────────────

function generate_alts_batch(array $alt_jobs): array
{
    // $alt_jobs: [row_idx => ['original', 'lang_name', 'field_type',
    //                         'deepl_ref', 'primary_text']]
    if (empty($alt_jobs)) return [];

    $groups = [];
    foreach ($alt_jobs as $row_idx => $job) {
        $key = $job['lang_name'] . '||' . $job['field_type'];
        $groups[$key][] = $row_idx;
    }

    $batches   = [];
    $batch_map = [];

    foreach ($groups as $chunk) {
        foreach (array_chunk($chunk, OPENAI_BATCH_SIZE) as $sub) {
            $first     = $alt_jobs[$sub[0]];
            $bidx      = count($batches);
            $batch_map[$bidx] = $sub;
            $batches[] = [
                'lang_name'  => $first['lang_name'],
                'field_type' => $first['field_type'],
                'items'      => array_map(fn($ri) => [
                    'original'  => $alt_jobs[$ri]['original'],
                    'deepl_ref' => $alt_jobs[$ri]['deepl_ref'],
                    'feedback'  => '',
                    'exclude'   => $alt_jobs[$ri]['primary_text'],
                ], $sub),
            ];
        }
    }

    $batch_results = openai_run_batches_concurrent($batches);

    $alts = [];
    foreach ($batch_map as $bidx => $row_indices) {
        $texts = $batch_results[$bidx] ?? [];
        foreach ($row_indices as $pos => $row_idx) {
            $text    = $texts[$pos] ?? null;
            $primary = $alt_jobs[$row_idx]['primary_text'];
            $ft      = $alt_jobs[$row_idx]['field_type'];

            if ($text && clean_text($text) !== $primary && is_valid_field(clean_text($text), $ft)) {
                $alts[$row_idx] = clean_text($text);
            } else {
                $alts[$row_idx] = '';
            }
        }
    }
    return $alts;
}

// ── Main entry point ───────────────────────────────────────────────────────

function process_csv(
    string $input_path,
    string $target_lang,
    string $output_path,
    string $log_file
): array {
    $in = @fopen($input_path, 'r');
    if (!$in) return ['error' => 'Cannot open uploaded file.'];

    $headers = fgetcsv($in);
    if (!$headers) { fclose($in); return ['error' => 'CSV file is empty.']; }

    $headers    = array_map('trim', $headers);
    $headers_ci = array_map('strtolower', $headers);
    $title_idx  = array_search('meta_title', $headers_ci);
    $desc_idx   = array_search('meta_description', $headers_ci);
    $lang_idx   = array_search('language', $headers_ci);

    // ── Read all rows ──────────────────────────────────────────────────────
    $all_rows = [];
    while (($row = fgetcsv($in)) !== false) $all_rows[] = $row;
    fclose($in);

    init_log($log_file);

    // ── Resolve per-row language ───────────────────────────────────────────
    $row_langs = [];
    foreach ($all_rows as $i => $row) {
        $raw = ($lang_idx !== false) ? trim($row[$lang_idx] ?? '') : '';
        $code = !empty($raw) ? normalize_language_code($raw) : $target_lang;
        $row_langs[$i] = ['code' => $code, 'name' => get_language_name($code)];
    }

    // ── DeepL pass: fast translation for supported languages ───────────────
    $deepl_title = [];
    $deepl_desc  = [];

    foreach ($all_rows as $i => $row) {
        $lang = $row_langs[$i];

        if ($title_idx !== false && trim($row[$title_idx] ?? '') !== '') {
            $t = deepl_translate(clean_text($row[$title_idx]), $lang['code']);
            $deepl_title[$i] = $t ? clean_text($t) : '';
        }
        if ($desc_idx !== false && trim($row[$desc_idx] ?? '') !== '') {
            $d = deepl_translate(clean_text($row[$desc_idx]), $lang['code']);
            $deepl_desc[$i] = $d ? clean_text($d) : '';
        }
    }

    // ── Build AI jobs for fields where DeepL was insufficient ─────────────
    $title_jobs = [];
    $desc_jobs  = [];

    foreach ($all_rows as $i => $row) {
        $lang = $row_langs[$i];

        if ($title_idx !== false && trim($row[$title_idx] ?? '') !== '') {
            $deepl = $deepl_title[$i] ?? '';
            if ($deepl !== '' && is_valid_field($deepl, 'title')) continue; // DeepL perfect
            $title_jobs[$i] = [
                'original'   => clean_text($row[$title_idx]),
                'lang'       => $lang['code'],
                'lang_name'  => $lang['name'],
                'deepl_ref'  => $deepl,
                'field_type' => 'title',
            ];
        }

        if ($desc_idx !== false && trim($row[$desc_idx] ?? '') !== '') {
            $deepl = $deepl_desc[$i] ?? '';
            if ($deepl !== '' && is_valid_field($deepl, 'description')) continue;
            $desc_jobs[$i] = [
                'original'   => clean_text($row[$desc_idx]),
                'lang'       => $lang['code'],
                'lang_name'  => $lang['name'],
                'deepl_ref'  => $deepl,
                'field_type' => 'description',
            ];
        }
    }

    // ── Run batch AI pipeline ──────────────────────────────────────────────
    $title_ai = run_batch_pipeline($title_jobs);
    $desc_ai  = run_batch_pipeline($desc_jobs);

    // ── Assemble per-row results ───────────────────────────────────────────
    $regenerated_statuses = ['openai_ok', 'groq_ok', 'meaning_fallback', 'best_available'];
    $rows    = [];
    $alt_jobs = [];

    foreach ($all_rows as $i => $row) {
        $lang = $row_langs[$i];

        // Title result
        if ($title_idx !== false && trim($row[$title_idx] ?? '') !== '') {
            $deepl = $deepl_title[$i] ?? '';
            if ($deepl !== '' && is_valid_field($deepl, 'title') && !isset($title_jobs[$i])) {
                $r_title = [
                    'text' => $deepl, 'status' => 'deepl_ok',
                    'len'  => mb_strlen($deepl, 'UTF-8'), 'valid' => true,
                    'attempts' => 1, 'deepl_ref' => $deepl,
                ];
            } else {
                $r_title = $title_ai[$i] ?? [
                    'text' => clean_text($row[$title_idx]), 'status' => 'skipped',
                    'len'  => mb_strlen(clean_text($row[$title_idx]), 'UTF-8'),
                    'valid' => false, 'attempts' => 0, 'deepl_ref' => '',
                ];
            }
        } else {
            $r_title = ['text' => '', 'status' => 'skipped', 'len' => 0, 'valid' => true, 'attempts' => 0, 'deepl_ref' => ''];
        }

        // Description result
        if ($desc_idx !== false && trim($row[$desc_idx] ?? '') !== '') {
            $deepl = $deepl_desc[$i] ?? '';
            if ($deepl !== '' && is_valid_field($deepl, 'description') && !isset($desc_jobs[$i])) {
                $r_desc = [
                    'text' => $deepl, 'status' => 'deepl_ok',
                    'len'  => mb_strlen($deepl, 'UTF-8'), 'valid' => true,
                    'attempts' => 1, 'deepl_ref' => $deepl,
                ];
            } else {
                $r_desc = $desc_ai[$i] ?? [
                    'text' => clean_text($row[$desc_idx]), 'status' => 'skipped',
                    'len'  => mb_strlen(clean_text($row[$desc_idx]), 'UTF-8'),
                    'valid' => false, 'attempts' => 0, 'deepl_ref' => '',
                ];
            }
        } else {
            $r_desc = ['text' => '', 'status' => 'skipped', 'len' => 0, 'valid' => true, 'attempts' => 0, 'deepl_ref' => ''];
        }

        // Queue alt title generation if title was AI-rewritten
        if ($r_title['text'] !== '' && in_array($r_title['status'], $regenerated_statuses, true)) {
            $alt_jobs[$i] = [
                'original'     => $title_idx !== false ? clean_text($row[$title_idx]) : '',
                'lang_name'    => $lang['name'],
                'field_type'   => 'title',
                'deepl_ref'    => $r_title['deepl_ref'] ?? '',
                'primary_text' => $r_title['text'],
            ];
        }

        // Queue alt description generation if description was AI-rewritten
        $alt_desc_jobs = $alt_desc_jobs ?? [];
        if ($r_desc['text'] !== '' && in_array($r_desc['status'], $regenerated_statuses, true)) {
            $alt_desc_jobs[$i] = [
                'original'     => $desc_idx !== false ? clean_text($row[$desc_idx]) : '',
                'lang_name'    => $lang['name'],
                'field_type'   => 'description',
                'deepl_ref'    => $r_desc['deepl_ref'] ?? '',
                'primary_text' => $r_desc['text'],
            ];
        }

        $rows[$i] = [
            'row'            => $row,
            'original_title' => $title_idx !== false ? ($row[$title_idx] ?? '') : '',
            'original_desc'  => $desc_idx  !== false ? ($row[$desc_idx]  ?? '') : '',
            'result_title'   => $r_title,
            'result_desc'    => $r_desc,
            'language'       => $lang['name'],
            'alt_title'      => '',
            'alt_desc'       => '',
        ];
    }

    // ── Batch alt title generation ─────────────────────────────────────────
    $alts = generate_alts_batch($alt_jobs);
    foreach ($alts as $i => $alt) {
        $rows[$i]['alt_title'] = $alt;
    }

    // ── Batch alt description generation ──────────────────────────────────
    $alt_descs = generate_alts_batch($alt_desc_jobs ?? []);
    foreach ($alt_descs as $i => $alt) {
        $rows[$i]['alt_desc'] = $alt;
    }

    // ── Write output CSV ───────────────────────────────────────────────────
    $out = @fopen($output_path, 'w');
    if (!$out) return ['error' => 'Cannot write output file.'];

    fwrite($out, "\xEF\xBB\xBF");
    fputcsv($out, array_merge($headers, ['new_title', 'alt_title', 'new_description', 'alt_description']));

    $result_rows = [];
    foreach ($rows as $i => $r) {
        fputcsv($out, array_merge($r['row'], [
            $r['result_title']['text'],
            $r['alt_title'],
            $r['result_desc']['text'],
            $r['alt_desc'],
        ]));
        $result_rows[] = [
            'original_title' => $r['original_title'],
            'original_desc'  => $r['original_desc'],
            'result_title'   => $r['result_title'],
            'alt_title'      => $r['alt_title'],
            'result_desc'    => $r['result_desc'],
            'alt_desc'       => $r['alt_desc'],
            'language'       => $r['language'],
        ];

        // Log each row
        log_row([
            'original' => $r['original_title'] ?: $r['original_desc'],
            'deepl'    => $r['result_title']['deepl_ref'] ?? '-',
            'openai'   => $r['result_title']['text'],
            'gemini'   => '-',
            'groq'     => '-',
            'final'    => $r['result_title']['text'],
            'attempts' => $r['result_title']['attempts'],
            'status'   => $r['result_title']['status'],
        ], $log_file);
    }

    fclose($out);
    return ['processed' => count($result_rows), 'rows' => $result_rows];
}
