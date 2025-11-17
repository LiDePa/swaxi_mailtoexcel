import argparse
from datetime import datetime
import sys
import pywintypes
import win32com.client

from mail import finde_ordner

FILTER_FORMATS = {
    "us_date": "%m/%d/%Y",
    "us_datetime": "%m/%d/%Y %I:%M %p",
    "de_date": "%d.%m.%Y",
    "de_datetime": "%d.%m.%Y %H:%M",
    "iso_date": "%Y-%m-%d",
    "iso_datetime": "%Y-%m-%d %H:%M",
}


def parse_args():
    parser = argparse.ArgumentParser(description="Probe different Outlook Restrict filters")
    parser.add_argument("start", help="Start timestamp, e.g. '12.11.2025 18:00'")
    parser.add_argument("end", help="End timestamp, e.g. '17.11.2025 13:00'")
    parser.add_argument(
        "--account",
        default="swaxi dispo",
        help="Shared mailbox display name to inspect",
    )
    return parser.parse_args()


def parse_dt(value):
    templates = ["%d.%m.%Y %H:%M", "%Y-%m-%d %H:%M", "%d.%m.%Y"]
    for tmpl in templates:
        try:
            return datetime.strptime(value, tmpl)
        except ValueError:
            continue
    raise SystemExit(f"Could not parse datetime '{value}' with known formats.")


def connect_to_outlook(account_name):
    try:
        outlook_app = win32com.client.Dispatch("Outlook.Application")
        mapi = outlook_app.GetNamespace("MAPI")
    except pywintypes.com_error as exc:
        raise SystemExit(f"Failed to attach to Outlook: {exc}")

    for acc in mapi.Folders:
        if acc.Name.lower() == account_name.lower():
            return acc
    raise SystemExit(f"Account '{account_name}' not found in Outlook profile.")


def run_probe(folder, filter_string):
    try:
        items = folder.Items.Restrict(filter_string)
        return list(items)
    except pywintypes.com_error as exc:
        print(f"  Restrict failed: {exc}")
        return []


def main():
    global args
    args = parse_args()
    start = parse_dt(args.start)
    end = parse_dt(args.end)
    mailbox = connect_to_outlook(args.account)

    archive = finde_ordner(mailbox, "Archive")
    dienstwuensche = finde_ordner(archive, "Dienstwünsche / FZ") if archive else None

    targets = {
        "Inbox": finde_ordner(mailbox, "Inbox") or finde_ordner(mailbox, "Posteingang"),
        "Dienstwünsche / FZ": dienstwuensche,
    }

    for name, folder in targets.items():
        if folder is None:
            print(f"Folder '{name}' not found, skipping.")
            continue
        print(f"\n=== {name} ===")
        for fmt_name, fmt in FILTER_FORMATS.items():
            filter_str = (
                f"[ReceivedTime] >= '{start.strftime(fmt)}' AND "
                f"[ReceivedTime] <= '{end.strftime(fmt)}'"
            )
            try:
                items = folder.Items.Restrict(filter_str)
                count = items.Count
                sample = items.Item(1) if count else None
            except pywintypes.com_error as exc:
                print(f"{fmt_name:>12}: Restrict failed -> {exc}")
                continue

            first_ts = None
            if sample and hasattr(sample, "ReceivedTime"):
                first_ts = sample.ReceivedTime
            print(f"{fmt_name:>12}: {count} items (first: {first_ts})")

    print("\nDone.")


if __name__ == "__main__":
    main()
