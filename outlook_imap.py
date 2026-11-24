#!/usr/bin/env python3
"""Outlook.com over IMAP with OAuth (XOAUTH2): authorization, refresh, and connection.

Usage:
    outlook_imap.py authorize <client-id>   device code flow (microsoft.com/link)
    outlook_imap.py check                   connects and lists folders, changes nothing

``<client-id>`` is a Microsoft Entra app registration accepting personal accounts with
public client flows enabled (no client secret). The refresh token is stored in
``$OUTLOOK_SORTER_TOKEN`` (default ``~/.config/outlook-spam-sorter/token.json``, mode 600),
is never printed, and is updated whenever Microsoft rotates it. User-facing messages stay
in French.
"""

import base64
import imaplib
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

AUTHORITY = "https://login.microsoftonline.com/consumers/oauth2/v2.0"
SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All"
HOST = "outlook.office365.com"
TOKEN = Path(os.environ.get("OUTLOOK_SORTER_TOKEN") or Path.home() / ".config/outlook-spam-sorter/token.json")
# Matched by the Dashboard ("Session Outlook") to show the reconnect tile.
EXPIRED = "Session Outlook expirée : relancer outlook_imap.py authorize"
ROLES = {
    "junk": ("\\Junk", ("junk", "junk email", "courrier indésirable")),
    "trash": ("\\Trash", ("deleted", "deleted items", "éléments supprimés")),
    "archive": ("\\Archive", ("archive", "archives", "archivage")),
}


class AuthError(Exception):
    """Raised when the OAuth grant is missing, expired, or rejected by Outlook."""


def post(url: str, fields: dict[str, str]) -> dict[str, Any]:
    """Send a form-encoded POST request and return the parsed JSON body."""
    req = urllib.request.Request(
        url,
        urllib.parse.urlencode(fields).encode(),
        {"Content-Type": "application/x-www-form-urlencoded"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as response:
            return json.load(response)
    except urllib.error.HTTPError as error:
        try:
            return json.load(error)
        except ValueError:
            return {"error": f"http_{error.code}"}


def save(state: dict[str, Any]) -> None:
    """Persist the OAuth state atomically with mode 600 permissions."""
    TOKEN.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    tmp = TOKEN.with_suffix(".tmp")
    fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w") as handle:
        json.dump(state, handle)
    os.replace(tmp, TOKEN)


def _extract_account_address(id_token: str) -> str:
    """Decode the account email address from the JWT payload of an OIDC id_token."""
    part = id_token.split(".")[1]
    claims = json.loads(base64.urlsafe_b64decode(part + "=" * (-len(part) % 4)))
    return (claims.get("email") or claims.get("preferred_username") or "").lower()


def _poll_device_token(client_id: str, device_code: str, interval: int, deadline: float) -> dict[str, Any]:
    """Poll Microsoft's token endpoint until authorization succeeds, fails, or expires."""
    while time.time() < deadline:
        time.sleep(interval)
        token_reply = post(
            f"{AUTHORITY}/token",
            {
                "client_id": client_id,
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "device_code": device_code,
            },
        )
        if token_reply.get("error") == "authorization_pending":
            continue
        if token_reply.get("error") == "slow_down":
            interval += 5
            continue
        if "refresh_token" not in token_reply or "id_token" not in token_reply:
            desc = str(token_reply.get("error_description"))[:300]
            sys.exit(f"Refus de Microsoft : {token_reply.get('error')} — {desc}")
        return token_reply
    sys.exit("Code expiré : relance la commande.")


def authorize(client_id: str) -> None:
    """Run the interactive OAuth device-code flow and store the resulting refresh token."""
    reply = post(f"{AUTHORITY}/devicecode", {"client_id": client_id, "scope": f"openid email offline_access {SCOPE}"})
    if "device_code" not in reply:
        desc = str(reply.get("error_description"))[:300]
        sys.exit(f"Refus de Microsoft : {reply.get('error')} — {desc}")
    minutes = reply["expires_in"] // 60
    print(
        f"Sur n'importe quel appareil, ouvre {reply['verification_uri']} et saisis le code "
        f"{reply['user_code']} (valable {minutes} min).",
        flush=True,
    )
    interval = int(reply.get("interval", 5))
    deadline = time.time() + int(reply["expires_in"])
    token_reply = _poll_device_token(client_id, reply["device_code"], interval, deadline)
    address = _extract_account_address(token_reply["id_token"])
    if "@" not in address:
        sys.exit("Microsoft n'a pas donné l'adresse du compte : accord non enregistré.")
    save({"client_id": client_id, "address": address, "refresh_token": token_reply["refresh_token"]})
    print(f"Accord enregistré pour {address}.")


def access_token() -> tuple[str, str]:
    """Return ``(address, access_token)``, raising AuthError if the grant cannot be refreshed."""
    try:
        state = json.loads(TOKEN.read_text())
    except OSError, ValueError:
        raise AuthError(EXPIRED) from None
    reply = post(
        f"{AUTHORITY}/token",
        {
            "client_id": state["client_id"],
            "grant_type": "refresh_token",
            "refresh_token": state["refresh_token"],
            "scope": f"offline_access {SCOPE}",
        },
    )
    if "access_token" not in reply:
        if reply.get("error") in ("invalid_grant", "interaction_required", "unauthorized_client", "invalid_client"):
            raise AuthError(EXPIRED)
        raise RuntimeError(f"Microsoft : {reply.get('error')}")
    if reply.get("refresh_token") and reply["refresh_token"] != state["refresh_token"]:
        save({**state, "refresh_token": reply["refresh_token"]})
    return state["address"], reply["access_token"]


def connect() -> imaplib.IMAP4_SSL:
    """Open an authenticated XOAUTH2 IMAP4_SSL connection to Outlook.com."""
    address, token = access_token()
    client = imaplib.IMAP4_SSL(HOST, 993, timeout=120)
    try:
        client.authenticate("XOAUTH2", lambda _: f"user={address}\x01auth=Bearer {token}\x01\x01".encode())
    except imaplib.IMAP4.error as error:
        raise AuthError(EXPIRED) from error
    return client


def mutf7(text: str) -> str:
    """Decode IMAP modified UTF-7 (RFC 3501) folder names to Unicode."""
    return re.sub(
        r"&([^-]*)-",
        lambda match: (
            "&"
            if not match[1]
            else base64.b64decode(match[1].replace(",", "/") + "=" * (-len(match[1]) % 4)).decode(
                "utf-16-be", "replace"
            )
        ),
        text,
    )


def folders(client: Any) -> dict[str, str]:
    """Resolve ``{"junk", "trash", "archive": imap_name}`` by special-use flag or decoded name."""
    typ, rows = client.list()
    if typ != "OK":
        raise RuntimeError("liste des dossiers refusée")
    found: dict[str, str] = {}
    for row in rows or []:
        match = re.match(r'\((?P<flags>[^)]*)\) "[^"]*" (?P<name>.+)$', row.decode(errors="replace"))
        if not match:
            continue
        name = match["name"].strip('"')
        flags = match["flags"].split()
        decoded_lower = mutf7(name).lower()
        for role, (flag, names) in ROLES.items():
            if flag in flags or (role not in found and decoded_lower in names):
                found[role] = name
    missing = [role for role in ROLES if role not in found]
    if missing:
        raise RuntimeError(f"dossier Outlook introuvable : {', '.join(missing)}")
    return found


def main(argv: list[str] | None = None) -> None:
    """CLI entry point for OAuth authorization and IMAP connectivity check."""
    args = sys.argv[1:] if argv is None else argv
    if args[:1] == ["authorize"] and len(args) == 2:
        authorize(args[1])
        return
    if args == ["check"]:
        try:
            client = connect()
        except AuthError as error:
            sys.exit(str(error))
        summary = ", ".join(f"{key} = {mutf7(val)}" for key, val in folders(client).items())
        print("Connexion IMAP : OK. Dossiers :", summary)
        client.logout()
        return
    sys.exit(__doc__)


if __name__ == "__main__":
    main()
