<?php

/**
 * Google OAuth 2.0 callback handler.
 *
 * Google redirects here after the user grants (or denies) consent.
 * On success the token is stored and the user is sent back to lang-generator.php.
 *
 * This URL must be listed as an authorized redirect URI in your Google Cloud
 * OAuth 2.0 client credentials (e.g. http://localhost:8085/google-callback.php).
 */

session_start();
require_once __DIR__ . '/vendor/autoload.php';

use App\GoogleAuthHelper;

// ── Denied / error ────────────────────────────────────────────────────────
if (isset($_GET['error'])) {
    $error = htmlspecialchars($_GET['error'], ENT_QUOTES, 'UTF-8');
    header('Location: lang-generator.php?google_status=auth_failed&error=' . urlencode($error));
    exit;
}

if (!isset($_GET['code'])) {
    header('Location: lang-generator.php?google_status=auth_failed&error=' . urlencode('No authorization code received'));
    exit;
}

// ── Exchange code for token ───────────────────────────────────────────────
$auth        = new GoogleAuthHelper();
$scheme      = (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] === 'on') ? 'https' : 'http';
$callbackUrl = $scheme . '://' . $_SERVER['HTTP_HOST'] . '/google-callback.php';

try {
    $auth->exchangeCode($_GET['code'], $callbackUrl);
    header('Location: lang-generator.php?google_status=connected');
} catch (\Exception $e) {
    header('Location: lang-generator.php?google_status=auth_failed&error=' . urlencode($e->getMessage()));
}
exit;
