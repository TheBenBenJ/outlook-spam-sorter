#!/usr/bin/env python3
"""Outlook.com over IMAP with OAuth (XOAUTH2): authorization, token refresh and connection, for imap-pass.py. No browser, no password.

outlook_imap.py authorize <client-id>   device code flow: prints a code to enter at microsoft.com/link, from any device
outlook_imap.py check                   connects and lists the folders, changes nothing

<client-id> is an app registration in Microsoft Entra that accepts personal Microsoft accounts and allows public client flows
(no client secret). The refresh token goes to $OUTLOOK_SORTER_TOKEN (default ~/.config/outlook-spam-sorter/token.json, mode 600),
is never printed, and is replaced whenever Microsoft rotates it. Messages that reach a person stay in French, like the rest of the pass.
"""
import base64, imaplib, json, os, re, sys, time, urllib.error, urllib.parse, urllib.request
from pathlib import Path

AUTHORITY = "https://login.microsoftonline.com/consumers/oauth2/v2.0"
SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All"
HOST = "outlook.office365.com"
TOKEN = Path(os.environ.get("OUTLOOK_SORTER_TOKEN") or Path.home() / ".config/outlook-spam-sorter/token.json")
# Read by the Dashboard, which looks for "Session Outlook" to show its "reconnect" tile.
EXPIRED = "Session Outlook expirée : relancer outlook_imap.py authorize"


class AuthError(Exception):
    pass


def post(url, fields):
    req = urllib.request.Request(url, urllib.parse.urlencode(fields).encode(), {"Content-Type": "application/x-www-form-urlencoded"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        try:
            return json.load(e)
        except ValueError:
            return {"error": f"http_{e.code}"}


def save(state):
    TOKEN.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    tmp = TOKEN.with_suffix(".tmp")
    fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w") as f:
        json.dump(state, f)
    os.replace(tmp, TOKEN)


def authorize(client_id):
    r = post(f"{AUTHORITY}/devicecode", {"client_id": client_id, "scope": f"openid email offline_access {SCOPE}"})
    if "device_code" not in r:
        sys.exit(f"Refus de Microsoft : {r.get('error')} — {str(r.get('error_description'))[:300]}")
    print(f"Sur n'importe quel appareil, ouvre {r['verification_uri']} et saisis le code {r['user_code']} (valable {r['expires_in'] // 60} min).", flush=True)
    interval, deadline = int(r.get("interval", 5)), time.time() + int(r["expires_in"])
    while time.time() < deadline:
        time.sleep(interval)
        t = post(f"{AUTHORITY}/token", {"client_id": client_id, "grant_type": "urn:ietf:params:oauth:grant-type:device_code", "device_code": r["device_code"]})
        if t.get("error") == "authorization_pending":
            continue
        if t.get("error") == "slow_down":
            interval += 5
            continue
        if "refresh_token" not in t or "id_token" not in t:
            sys.exit(f"Refus de Microsoft : {t.get('error')} — {str(t.get('error_description'))[:300]}")
        part = t["id_token"].split(".")[1]
        claims = json.loads(base64.urlsafe_b64decode(part + "=" * (-len(part) % 4)))
        address = (claims.get("email") or claims.get("preferred_username") or "").lower()
        if "@" not in address:
            sys.exit("Microsoft n'a pas donné l'adresse du compte : accord non enregistré.")
        save({"client_id": client_id, "address": address, "refresh_token": t["refresh_token"]})
        print(f"Accord enregistré pour {address}.")
        return
    sys.exit("Code expiré : relance la commande.")


def access_token():
    """-> (address, access token). AuthError when the grant is missing, revoked or expired: only a person can renew it."""
    try:
        state = json.loads(TOKEN.read_text())
    except (OSError, ValueError):
        raise AuthError(EXPIRED)
    t = post(f"{AUTHORITY}/token", {"client_id": state["client_id"], "grant_type": "refresh_token", "refresh_token": state["refresh_token"], "scope": f"offline_access {SCOPE}"})
    if "access_token" not in t:
        if t.get("error") in ("invalid_grant", "interaction_required", "unauthorized_client", "invalid_client"):
            raise AuthError(EXPIRED)
        raise RuntimeError(f"Microsoft : {t.get('error')}")
    if t.get("refresh_token") and t["refresh_token"] != state["refresh_token"]:
        save({**state, "refresh_token": t["refresh_token"]})
    return state["address"], t["access_token"]


def connect():
    address, token = access_token()
    m = imaplib.IMAP4_SSL(HOST, 993, timeout=120)
    try:
        m.authenticate("XOAUTH2", lambda _: f"user={address}\x01auth=Bearer {token}\x01\x01".encode())
    except imaplib.IMAP4.error as e:
        raise AuthError(EXPIRED) from e
    return m


def mutf7(s):
    """Folder names are in modified UTF-7 (RFC 3501): "Courrier ind&AOk-sirable" reads "Courrier indésirable"."""
    return re.sub(r"&([^-]*)-", lambda x: "&" if not x[1] else base64.b64decode(x[1].replace(",", "/") + "=" * (-len(x[1]) % 4)).decode("utf-16-be", "replace"), s)


ROLES = {"junk": ("\\Junk", ("junk", "junk email", "courrier indésirable")), "trash": ("\\Trash", ("deleted", "deleted items", "éléments supprimés")),
         "archive": ("\\Archive", ("archive", "archives", "archivage"))}


def folders(m):
    """-> {"junk", "trash", "archive": IMAP name}, found by special-use flag, else by name (Outlook.com flags Junk and Trash, not its archive)."""
    typ, rows = m.list()
    if typ != "OK":
        raise RuntimeError("liste des dossiers refusée")
    found = {}
    for row in rows or []:
        x = re.match(r'\((?P<flags>[^)]*)\) "[^"]*" (?P<name>.+)$', row.decode(errors="replace"))
        if not x:
            continue
        name = x["name"].strip('"')
        for role, (flag, names) in ROLES.items():
            if flag in x["flags"].split() or (role not in found and mutf7(name).lower() in names):
                found[role] = name
    missing = [r for r in ROLES if r not in found]
    if missing:
        raise RuntimeError(f"dossier Outlook introuvable : {', '.join(missing)}")
    return found


if __name__ == "__main__":
    a = sys.argv[1:]
    if a[:1] == ["authorize"] and len(a) == 2:
        authorize(a[1])
    elif a == ["check"]:
        try:
            m = connect()
        except AuthError as e:
            sys.exit(str(e))
        print("Connexion IMAP : OK. Dossiers :", ", ".join(f"{k} = {mutf7(v)}" for k, v in folders(m).items()))
        m.logout()
    else:
        sys.exit(__doc__)
