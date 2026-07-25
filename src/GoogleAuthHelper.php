<?php

declare(strict_types=1);

namespace App;

use Google\Client;

/**
 * Manages Google OAuth 2.0 credentials and token lifecycle.
 *
 * Google Cloud setup:
 *   1. Go to https://console.cloud.google.com/
 *   2. Create or select a project
 *   3. Enable "Google Sheets API" and "Google Drive API"
 *   4. Navigate to APIs & Services → Credentials
 *   5. Create OAuth 2.0 Client ID (type: Web Application)
 *   6. Add your redirect URI: http://localhost:8085/google-callback.php
 *   7. Download the JSON file and save it as google-credentials.json in the app root
 */
class GoogleAuthHelper
{
    private const SCOPES = [
        'https://www.googleapis.com/auth/documents',
        'https://www.googleapis.com/auth/drive.file',
    ];

    private string $credentialsPath;
    private string $tokenPath;

    public function __construct(string $baseDir = '')
    {
        if ($baseDir === '') {
            $baseDir = dirname(__DIR__);
        }
        $this->credentialsPath = $baseDir . '/google-credentials.json';
        $this->tokenPath       = $baseDir . '/output/.google-token.json';
    }

    public function hasCredentials(): bool
    {
        return file_exists($this->credentialsPath);
    }

    public function hasToken(): bool
    {
        return file_exists($this->tokenPath);
    }

    public function isAuthenticated(): bool
    {
        if (!$this->hasCredentials() || !$this->hasToken()) {
            return false;
        }
        try {
            $client = $this->createClient();
            return !$client->isAccessTokenExpired();
        } catch (\Exception) {
            return false;
        }
    }

    /**
     * Returns a configured Google Client with a valid access token (if available).
     */
    public function createClient(): Client
    {
        $client = new Client();
        $client->setApplicationName('Doc Generator');
        $client->setScopes(self::SCOPES);
        $client->setAuthConfig($this->credentialsPath);
        $client->setAccessType('offline');
        $client->setPrompt('select_account consent');

        if ($this->hasToken()) {
            $token = json_decode(file_get_contents($this->tokenPath), true);
            $client->setAccessToken($token);

            if ($client->isAccessTokenExpired() && $client->getRefreshToken()) {
                $newToken = $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
                if (!isset($newToken['error'])) {
                    $this->saveToken($client->getAccessToken());
                }
            }
        }

        return $client;
    }

    /**
     * Builds the Google OAuth authorization URL.
     */
    public function getAuthUrl(string $redirectUri): string
    {
        $client = new Client();
        $client->setAuthConfig($this->credentialsPath);
        $client->setScopes(self::SCOPES);
        $client->setRedirectUri($redirectUri);
        $client->setAccessType('offline');
        $client->setPrompt('select_account consent');
        return $client->createAuthUrl();
    }

    /**
     * Exchanges an authorization code for an access+refresh token and persists it.
     *
     * @throws \RuntimeException on API error
     */
    public function exchangeCode(string $code, string $redirectUri): void
    {
        $client = new Client();
        $client->setAuthConfig($this->credentialsPath);
        $client->setScopes(self::SCOPES);
        $client->setRedirectUri($redirectUri);

        $token = $client->fetchAccessTokenWithAuthCode($code);

        if (isset($token['error'])) {
            throw new \RuntimeException(
                'OAuth token exchange failed: ' . ($token['error_description'] ?? $token['error'])
            );
        }

        $this->saveToken($token);
    }

    /**
     * Revokes the stored token and removes it from disk.
     */
    public function revokeToken(): void
    {
        if ($this->hasToken()) {
            try {
                $client = $this->createClient();
                $client->revokeToken();
            } catch (\Exception) {
                // best-effort revocation
            }
            @unlink($this->tokenPath);
        }
    }

    private function saveToken(array $token): void
    {
        $dir = dirname($this->tokenPath);
        if (!is_dir($dir)) {
            mkdir($dir, 0777, true);
        }
        file_put_contents($this->tokenPath, json_encode($token));
    }
}
