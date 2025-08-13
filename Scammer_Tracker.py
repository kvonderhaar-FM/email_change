# -*- coding: utf-8 -*-
import os, sys, time, csv, re, datetime, traceback, requests
from msal import PublicClientApplication, SerializableTokenCache

try:
    from tqdm import tqdm
except Exception:
    tqdm = None  # fallback to simple scan counters

# ====== CONFIG ======
TENANT_ID = "b58880c5-a46b-406e-a7df-53bdd1ffeb40"
CLIENT_ID = "5e0bd539-2974-4c90-87f4-366d2a699619"
MAILBOX   = "gmalawi@fabricatedmetals.com"

SCOPES      = ["Mail.Read"]
PAGE_SIZE   = 50
DAYS_BACK   = int(os.environ.get("DAYS_BACK", 5))      # window to scan
MAX_SCAN    = int(os.environ.get("MAX_SCAN", 100))     # cap on messages scanned
MAX_RESULTS = int(os.environ.get("MAX_RESULTS", 100))   # cap on matches saved/printed
CACHE_PATH  = os.path.expanduser("~/.msal_mailbox_cache.bin")

# ====== TERMS (word boundaries for single words; literal for phrases) ======
CHANGE_SINGLE = ["change","modify","switch","update","alter","revise","amend","edit"]

ADDR_SINGLE   = ["address"]
ADDR_PHRASE   = ["mailing address","billing address","remit to","remittance address","ship to"]

BANK_SINGLE   = ["ach"]
BANK_PHRASE   = [
    "bank account","bank information","bank details",
    "routing information","routing number","account number",
    "wire instructions","direct deposit","payment details",
    "routing #","acct #","account #","bank acct"
]

def _compile_terms(single_terms, phrase_terms):
    parts = [r"\b{}\b".format(re.escape(t)) for t in single_terms]
    parts += [re.escape(p) for p in phrase_terms]
    return re.compile("|".join(parts), re.IGNORECASE) if parts else None

PAT_CHANGE   = _compile_terms(CHANGE_SINGLE, CHANGE_PHRASE)
PAT_ADDRBANK = _compile_terms(ADDR_SINGLE + BANK_SINGLE, ADDR_PHRASE + BANK_PHRASE)

def qualifies(subject, preview):
    text = "{}\n{}".format(subject or "", preview or "")
    return bool(PAT_CHANGE.search(text) and PAT_ADDRBANK.search(text))

def iso_days_ago(n):
    # timezone-aware UTC (Python 3.11+ has datetime.UTC)
    dt = datetime.datetime.now(datetime.UTC) - datetime.timedelta(days=n)
    return dt.replace(microsecond=0).isoformat().replace("+00:00", "Z")

def base_url():
    return ("https://graph.microsoft.com/v1.0/me/messages"
            if MAILBOX.strip().lower() == "me"
            else "https://graph.microsoft.com/v1.0/users/{}/messages".format(MAILBOX))

def get_token():
    cache = SerializableTokenCache()
    if os.path.exists(CACHE_PATH):
        try:
            cache.deserialize(open(CACHE_PATH, "r").read())
        except Exception:
            pass
    app = PublicClientApplication(
        client_id=CLIENT_ID,
        authority="https://login.microsoftonline.com/{}".format(TENANT_ID),
        token_cache=cache
    )
    result = app.acquire_token_silent(SCOPES, account=None)
    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError("Device flow failed: {}".format(flow))
        print("[ACTION] Open {} and enter code: {}".format(flow["verification_uri"], flow["user_code"]))
        result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError("Auth failed: {}".format(result))
    open(CACHE_PATH, "w").write(cache.serialize())
    return result["access_token"]

def fetch_pages(token):
    """Yield (messages_page, total_estimate) where total_estimate may be None."""
    url = base_url()
    since = iso_days_ago(DAYS_BACK)
    params = {
        "$filter": "receivedDateTime ge {}".format(since),
        "$orderby": "receivedDateTime desc",
        "$top": str(PAGE_SIZE),
        "$select": "id,subject,from,receivedDateTime,webLink,internetMessageId,bodyPreview",
        "$count": "true",
    }
    headers = {
        "Authorization": "Bearer {}".format(token),
        "ConsistencyLevel": "eventual",
        "Prefer": 'outlook.body-content-type="text"',
    }
    first = True
    total_est = None
    while True:
        r = requests.get(url, headers=headers, params=params)
        if r.status_code == 429:
            time.sleep(int(r.headers.get("Retry-After", "5"))); continue
        r.raise_for_status()
        data = r.json()
        if first:
            total_est = data.get("@odata.count")
            first = False
        yield data.get("value", []), total_est
        next_link = data.get("@odata.nextLink")
        if not next_link:
            break
        url = next_link
        params = None

def main():
    try:
        print("[INFO] Strict local matching (no $search). Window: last {} days.".format(DAYS_BACK))
        print("[INFO] Caps: MAX_SCAN={}, MAX_RESULTS={}\n".format(MAX_SCAN, MAX_RESULTS))
        token = get_token()

        print("email_address\treceivedDateTime\tsubject")
        print("-" * 90)

        matches = []
        scanned = 0

        # progress bar tracks emails scanned
        bar = None
        first_page = True
        total_est = None

        for page, total in fetch_pages(token):
            if first_page:
                total_est = total
                if tqdm:
                    bar = tqdm(total=total_est, unit="msg", dynamic_ncols=True) if total_est else tqdm(unit="msg", dynamic_ncols=True)
                first_page = False

            for msg in page:
                scanned += 1
                if bar:
                    bar.update(1)
                elif scanned % 200 == 0:
                    if total_est:
                        pct = scanned / total_est * 100
                        print("[SCAN] {}/{} ({:.1f}%)".format(scanned, total_est, pct))
                    else:
                        print("[SCAN] {} scanned...".format(scanned))

                sender = (msg.get("from") or {}).get("emailAddress", {}).get("address")
                rdt    = msg.get("receivedDateTime")
                subj   = (msg.get("subject") or "").replace("\r"," ").replace("\n"," ").strip()
                prev   = msg.get("bodyPreview") or ""

                if qualifies(subj, prev):
                    print("{}\t{}\t{}".format(sender or "", rdt or "", subj))
                    matches.append({
                        "id": msg.get("id"),
                        "internetMessageId": msg.get("internetMessageId"),
                        "from": sender,
                        "receivedDateTime": rdt,
                        "subject": subj,
                        "webLink": msg.get("webLink"),
                    })

                if scanned >= MAX_SCAN or len(matches) >= MAX_RESULTS:
                    break

            if scanned >= MAX_SCAN or len(matches) >= MAX_RESULTS:
                break

        if bar:
            bar.close()

        out = "hits.csv"
        if matches:
            with open(out, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=list(matches[0].keys()))
                w.writeheader(); w.writerows(matches)
            print("\n[DONE] Printed {} matches (scanned {} msgs) and saved to {}".format(len(matches), scanned, out))
        else:
            print("\n[DONE] No matches found (scanned {} msgs).".format(scanned))

        print("\nNotes:")
        print("  - Progress bar shows ALL messages scanned.")
        print("  - Only qualifying emails are printed/saved.")
        print("  - Tune speed/scope with DAYS_BACK, MAX_SCAN, MAX_RESULTS env vars.")
        print("  - Optional: pip install tqdm for a nicer progress bar.")
    except Exception:
        print("[FATAL]"); traceback.print_exc(); sys.exit(1)

if __name__ == "__main__":
    main()