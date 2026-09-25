"""Forward unread Inbox mail and mark it read."""
import argparse, imaplib, json, re, smtplib, sqlite3, ssl
from base64 import b64decode, b64encode
from contextlib import closing
from copy import deepcopy
from datetime import datetime, timedelta, timezone
from email import policy
from email.encoders import encode_base64
from email.parser import BytesParser
from email.utils import formataddr, formatdate, getaddresses, make_msgid, parsedate_to_datetime
from pathlib import Path
from sys import stderr

import msal
from msal_extensions import PersistedTokenCache, build_encrypted_persistence
from msal_extensions.persistence import PersistenceDecryptionError

CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"  # Thunderbird ID
ROOT = Path(__file__).resolve().parent
STATE = ROOT / "MailFWD.json"  # Save encrypted authentication info within a json file
SCOPE = ["https://outlook.office.com/IMAP.AccessAsUser.All", "https://outlook.office.com/SMTP.Send"]


def checked(command, *args):
    status, data = command(*args)
    if status != "OK":
        raise RuntimeError(f"{args[0] if args else command.__name__} failed, check account access.")
    return data


def credentials(account, state, login=False):
    path = ROOT / ".login.bin"
    if not path.exists() and "login" in state:
        temporary = path.with_suffix(".tmp")
        temporary.write_bytes(b64decode(state["login"], validate=True))
        temporary.replace(path)
    cache = PersistedTokenCache(build_encrypted_persistence(str(path)))
    app = msal.PublicClientApplication(CLIENT_ID, token_cache=cache,
                                      authority=state.get("authority", "https://login.microsoftonline.com/organizations"))
    accounts = app.get_accounts(username=account)
    result = app.acquire_token_silent(SCOPE, account=accounts[0]) if accounts and not login else None
    if not result or "access_token" not in result:
        flow = app.initiate_device_flow(scopes=SCOPE)
        if "user_code" not in flow:
            raise RuntimeError("Device sign-in is unavailable for this app.")
        print(flow["message"], file=stderr, flush=True)
        result = app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise RuntimeError("Sign-in failed, check account permissions.")
        if result.get("id_token_claims", {}).get("preferred_username", "").lower() != account.lower():
            raise RuntimeError("Sign in with the requested login name.")
    if path.exists():
        state["login"] = b64encode(path.read_bytes()).decode("ascii")
        state.pop("flows", None)
        state.pop("sent", None)
        temporary = STATE.with_suffix(".tmp")
        temporary.write_text(json.dumps(state), encoding="utf-8")
        temporary.replace(STATE)
        path.unlink()
    return result["access_token"]


def prepare_forward(message, account, recipient):
    message = deepcopy(message)

    def addresses(*names):
        values = [value for name in names for value in message.get_all(name, [])]
        return [(name, address) for name, address in getaddresses(values) if address]

    authors = addresses("From")
    author = next((name or address for name, address in authors), "Source mail")
    reply = addresses("Reply-To") or authors
    excluded = {account.lower(), recipient.lower()} | {address.lower() for _, address in reply}
    cc = {}
    for name, address in addresses("To", "Cc"):
        if address.lower() not in excluded:
            cc.setdefault(address.lower(), formataddr((name, address)))
    message_id = message.get("Message-ID", "")
    references = f"{message.get('References', message.get('In-Reply-To', ''))} {message_id}"
    headers = {
        "From": formataddr((" ".join(author.splitlines()) + " via MailFWD", account)),
        "To": recipient, "Subject": " ".join(str(message.get("Subject", "")).splitlines()),
        "Date": formatdate(localtime=True), "Message-ID": make_msgid(),
        "Reply-To": ", ".join(map(formataddr, reply)), "Cc": ", ".join(cc.values()),
        "In-Reply-To": message_id, "References": " ".join(references.splitlines()).strip(),
    }
    # Keep the original MIME payload intact, replacing only its outer headers.
    for name in message.keys():
        if not (name.lower().startswith("content-") or name.lower() == "mime-version"):
            del message[name]
    for name, value in headers.items():
        if value or name == "Subject":
            message[name] = value
    for part in message.walk():
        if part.is_multipart() or part.get("Content-Transfer-Encoding", "").lower() in ("base64", "quoted-printable"):
            continue
        body = part.get_payload(decode=True)
        if body and (not body.isascii() or b"\x00" in body):
            del part["Content-Transfer-Encoding"]
            part.set_payload(body)
            encode_base64(part)
    message.policy = policy.SMTP.clone(cte_type="7bit")
    return message


def parse_time(value):
    try:
        if not re.fullmatch(r"[0-9]{12}", value):
            raise ValueError
        return datetime.strptime(value, "%Y%m%d%H%M").astimezone()
    except ValueError:
        raise argparse.ArgumentTypeError("Use a valid local time in YYYYMMDDHHMM format.")


def forward_mail(account, recipients, token, after=None, verbose=False):
    count = 0
    tls = ssl.create_default_context()
    with imaplib.IMAP4_SSL("outlook.office365.com", ssl_context=tls, timeout=60) as source:
        payload = f"user={account}\x01auth=Bearer {token}\x01\x01"
        answers = iter((payload.encode(), b""))
        source.authenticate("XOAUTH2", lambda _: next(answers, b""))
        checked(source.select, "INBOX")
        search = ["UNSEEN"]
        if after is not None:
            day = after.astimezone(timezone.utc) - timedelta(days=1)
            search = ["SINCE", imaplib.Time2Internaldate(day).split()[0].strip('"')]
        uids = sorted(checked(source.uid, "SEARCH", None, "UNDELETED", *search)[0].decode().split(), key=int)
        if not uids:
            return 0
        with smtplib.SMTP("smtp.office365.com", 587, timeout=60) as smtp:
            smtp.starttls(context=tls)
            smtp.ehlo()
            smtp.auth("XOAUTH2", lambda challenge=None: payload if challenge is None else "")
            for uid in uids:
                parts = checked(source.uid, "FETCH", uid, "(INTERNALDATE BODY.PEEK[])")
                raw = next((part[1] for part in parts if isinstance(part, tuple)), None)
                if raw is None:  # Moved/deleted during this run.
                    continue
                metadata = b" ".join(part[0] if isinstance(part, tuple) else part for part in parts if part)
                date = re.search(rb'INTERNALDATE "([^"]+)"', metadata)
                if not date:
                    raise RuntimeError("Source did not return the message's received time.")
                received = parsedate_to_datetime(date[1].decode("ascii"))
                if after is not None and received <= after:
                    continue
                original = BytesParser(policy=policy.default).parsebytes(raw)
                for recipient in recipients:
                    forward = prepare_forward(original, account, recipient)
                    smtp.send_message(forward, from_addr=account, to_addrs=[recipient])
                    count += 1
                    if verbose:
                        sender = " ".join(str(original.get("From", "(unknown sender)")).split())
                        print(f"Forwarded {uid}\n  Flow: {account} -> {recipient}"
                              f"\n  Received: {received.astimezone().isoformat(' ', timespec='seconds')}"
                              f"\n  From: {sender!r}\n  Subject: {str(forward['Subject'])!r}\n", flush=True)
                checked(source.uid, "STORE", uid, "+FLAGS.SILENT", r"(\Seen)")
    return count


def main():
    parser = argparse.ArgumentParser(description="Forward unread Inbox mail, or all Inbox mail after a local time.")
    parser.add_argument("--from", dest="account", required=True, metavar="ADDRESS", help="Source login address")
    parser.add_argument("--to", nargs="+", required=True, metavar="ADDRESS",
                        type=lambda value: value if "@" in value else parse_time(value),
                        help="recipients, optionally followed by YYYYMMDDHHMM (local time)")
    parser.add_argument("--login", action="store_true", help="sign in again")
    parser.add_argument("--verbose", action="store_true", help="print each forward's sender, subject and received time")
    args = parser.parse_args()
    after = args.to.pop() if isinstance(args.to[-1], datetime) else None
    if not args.to:
        parser.error("--to requires at least one email address")
    for address in [args.account, *args.to]:
        if not isinstance(address, str) or not re.fullmatch(r"[^\s@<>,;]+@[^\s@<>,;]+", address):
            raise ValueError(f"Invalid email address: {address}")
    # SQLite provides a crash-safe process lock on Windows, macOS and Linux.
    with closing(sqlite3.connect(ROOT / ".sync.lock", timeout=0)) as lock:
        lock.execute("BEGIN EXCLUSIVE")
        state = json.loads(STATE.read_text(encoding="utf-8")) if STATE.exists() else {}
        token = credentials(args.account, state, login=args.login)
        recipients = list({address.lower(): address for address in args.to}.values())
        print(forward_mail(args.account, recipients, token, after, verbose=args.verbose), flush=True)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("Stopped.", file=stderr)
    except (RuntimeError, ValueError) as error:
        raise SystemExit(str(error))
    except smtplib.SMTPAuthenticationError as error:
        raise SystemExit(f"Source rejected SMTP sign-in ({error.smtp_code}), check whether it permits SMTP AUTH for your account.")
    except PersistenceDecryptionError:
        raise SystemExit("Cannot decrypt saved sign-in. Run Dagu under the same Windows account and profile "
                         "used to sign in, on the same computer. --login cannot unlock this cache.")
    except Exception as error:
        raise SystemExit(f"Stopped ({type(error).__name__}). Check settings, credentials and account access.")
