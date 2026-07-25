<?php

/**
 * Initiates the Google OAuth 2.0 authorization flow.
 *
 * Usage:
 *   Visit /google-auth.php                  → redirects to Google consent screen
 *   Visit /google-auth.php?action=revoke    → revokes the stored token
 *
 * After the user grants consent, Google redirects to google-callback.php.
 */

require_once __DIR__ . '/vendor/autoload.php';

use App\GoogleAuthHelper;

$auth = new GoogleAuthHelper();

// ── Revoke action ──────────────────────────────────────────────────────────
if (isset($_GET['action']) && $_GET['action'] === 'revoke') {
    $auth->revokeToken();
    header('Location: lang-generator.php?google_status=revoked');
    exit;
}

// ── Guard: credentials file must exist ────────────────────────────────────
if (!$auth->hasCredentials()) {
    header('Location: lang-generator.php?google_status=no_credentials');
    exit;
}

// ── Redirect to Google consent screen ─────────────────────────────────────
$scheme      = (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] === 'on') ? 'https' : 'http';
$callbackUrl = $scheme . '://' . $_SERVER['HTTP_HOST'] . '/google-callback.php';

$authUrl = $auth->getAuthUrl($callbackUrl);
header('Location: ' . $authUrl);
exit;
