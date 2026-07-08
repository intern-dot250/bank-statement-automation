"""Password-only login backed by Supabase Auth.

This app has a single shared team login: one fixed Supabase Auth user
(email set via APP_LOGIN_EMAIL) that the UI never shows. The login page
only asks for a password; the email is filled in internally before
calling Supabase's password sign-in.

Required environment variables:
    SUPABASE_URL        e.g. https://oyaifetwchwyohzfbndu.supabase.co
    SUPABASE_ANON_KEY    the "Publishable" / anon API key
    APP_LOGIN_EMAIL      the email of the one shared Supabase Auth user
"""

from __future__ import annotations

import logging
import os

logger = logging.getLogger(__name__)

_client = None


def _get_client():
    """Return a cached Supabase client, or None if not configured."""
    global _client
    if _client is not None:
        return _client

    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_ANON_KEY")
    if not url or not key:
        logger.error("SUPABASE_URL/SUPABASE_ANON_KEY not configured.")
        return None

    from supabase import create_client  # imported lazily; optional dependency

    _client = create_client(url, key)
    return _client


def login(password: str) -> dict | None:
    """Attempt to sign in with the fixed team email and the given password.

    Returns:
        A dict with "access_token" and "refresh_token" on success, or
        None if login failed (wrong password, misconfiguration, etc.).
    """
    email = os.environ.get("APP_LOGIN_EMAIL")
    if not email:
        logger.error("APP_LOGIN_EMAIL not configured.")
        return None

    client = _get_client()
    if client is None:
        return None

    try:
        response = client.auth.sign_in_with_password({"email": email, "password": password})
    except Exception as exc:
        logger.info("Login failed: %s", exc)
        return None

    if not response or not response.session:
        return None

    return {
        "access_token": response.session.access_token,
        "refresh_token": response.session.refresh_token,
    }


def logout(refresh_token: str | None) -> None:
    """Best-effort sign-out; failures are non-fatal since the Flask
    session is cleared by the caller regardless."""
    client = _get_client()
    if client is None:
        return
    try:
        client.auth.sign_out()
    except Exception as exc:
        logger.info("Sign-out call failed (session is still cleared locally): %s", exc)
