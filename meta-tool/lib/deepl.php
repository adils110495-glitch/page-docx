<?php

declare(strict_types=1);

function deepl_translate(string $text, string $target_lang): ?string
{
    $api_key = getenv('DEEPL_API_KEY');
    if (!$api_key) return null;

    // Free tier keys end with ':fx'
    $url = str_ends_with($api_key, ':fx')
        ? 'https://api-free.deepl.com/v2/translate'
        : 'https://api.deepl.com/v2/translate';

    $payload = json_encode([
        'text'        => [$text],
        'target_lang' => strtoupper($target_lang),
    ]);

    $ch = curl_init($url);
    curl_setopt_array($ch, [
        CURLOPT_POST          => true,
        CURLOPT_POSTFIELDS    => $payload,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_HTTPHEADER    => [
            'Authorization: DeepL-Auth-Key ' . $api_key,
            'Content-Type: application/json',
        ],
        CURLOPT_TIMEOUT => 30,
    ]);

    $response = curl_exec($ch);

    if (!$response) return null;

    $data = json_decode($response, true);
    return $data['translations'][0]['text'] ?? null;
}
