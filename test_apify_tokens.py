#!/usr/bin/env python3
"""
test_apify_tokens.py - check every Apify API token before relying on it.

For each token in APIFY_API_TOKENS (comma / newline separated, same variable
script.py uses) this:
  1. reads the account behind the token (/users/me) and its plan
  2. reads the monthly credit limit and how much is used (/users/me/limits)
  3. starts a tiny real run (1 product URL) of the Amazon actor and of the
     BestBuy actor that script.py uses, and reports the price that came back
     or the exact error Apify returned (e.g. a permission / credit problem).

Cost: about $0.003 per actor per token (~$0.03 for 5 tokens).

Usage:
  APIFY_API_TOKENS="tok1,tok2,..." python test_apify_tokens.py
  python test_apify_tokens.py --no-runs     # only account + credit checks, free
"""
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

API = "https://api.apify.com/v2"
AMAZON_ACTOR = "U3DyJ7kdhQlYyeQKd"    # delicious_zebu/amazon-product-details-scraper
BESTBUY_ACTOR = "pbUZ4z2ORsyKhZshL"   # benthepythondev/bestbuy-scraper

AMAZON_TEST_URL = "https://www.amazon.com/dp/B0H1DG7XRW"   # Watch9 40mm Cream Bluetooth
BESTBUY_TEST_URL = ("https://www.bestbuy.com/product/"
                    "samsung-galaxy-z-fold8-ultra-256gb-unlocked-cream/JJGRF3TZF2")


def call(method, path, token, body=None, params=None, timeout=90):
    """Return (http_status, parsed_json_or_text)."""
    url = API + path + ("?" + urllib.parse.urlencode(params) if params else "")
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        "Authorization": f"Bearer {token}", "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            raw = r.read().decode("utf-8")
            return r.status, (json.loads(raw) if raw else None)
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            return e.code, json.loads(raw)
        except Exception:
            return e.code, raw


def err_text(payload):
    if isinstance(payload, dict) and isinstance(payload.get("error"), dict):
        e = payload["error"]
        return f"{e.get('type')}: {e.get('message')}"
    return str(payload)[:300]


def run_actor(token, actor_id, actor_input):
    """Start a run, wait for it, return (ok, detail, items)."""
    st, body = call("POST", f"/acts/{actor_id}/runs", token, body=actor_input,
                    params={"timeout": 300})
    if st != 201:
        return False, f"could not start run -> HTTP {st} {err_text(body)}", None
    run = body["data"]
    while run.get("status") not in ("SUCCEEDED", "FAILED", "ABORTED", "TIMED-OUT"):
        st, body = call("GET", f"/actor-runs/{run['id']}", token,
                        params={"waitForFinish": 60})
        if st != 200:
            return False, f"polling failed -> HTTP {st} {err_text(body)}", None
        run = body["data"]
    st, items = call("GET", f"/datasets/{run['defaultDatasetId']}/items", token,
                     params={"clean": "true", "format": "json"})
    items = items if isinstance(items, list) else []
    if run["status"] != "SUCCEEDED":
        return False, f"run {run['status']}: {run.get('statusMessage')}", items
    return True, f"run SUCCEEDED ({len(items)} item(s))", items


def main():
    raw = os.environ.get("APIFY_API_TOKENS") or os.environ.get("APIFY_API_TOKEN") or ""
    tokens = [t for t in re.split(r"[\s,;]+", raw) if t]
    if not tokens:
        sys.exit("Set APIFY_API_TOKENS (comma separated) first.")
    do_runs = "--no-runs" not in sys.argv
    dump = {}

    for i, tok in enumerate(tokens, start=1):
        print(f"\n===== token #{i} (...{tok[-4:]}) =====")
        st, me = call("GET", "/users/me", tok)
        if st != 200:
            print(f"  account : FAILED HTTP {st} {err_text(me)}")
            continue
        d = me["data"]
        plan = (d.get("plan") or {}).get("id") or "?"
        print(f"  account : {d.get('username')}  plan={plan}")

        st, lim = call("GET", "/users/me/limits", tok)
        if st == 200:
            ld = lim["data"]
            mx = (ld.get("limits") or {}).get("maxMonthlyUsageUsd")
            used = (ld.get("current") or {}).get("monthlyUsageUsd")
            left = (mx - used) if isinstance(mx, (int, float)) and isinstance(used, (int, float)) else None
            print(f"  credit  : used ${used} of ${mx} this cycle"
                  + (f" -> ${left:.2f} left" if left is not None else ""))
        else:
            print(f"  credit  : limits not readable HTTP {st} {err_text(lim)}")

        if not do_runs:
            continue

        ok, detail, items = run_actor(tok, AMAZON_ACTOR, {
            "Params": [AMAZON_TEST_URL], "deliverTo": "US", "zipCode": "10001"})
        price = items[0].get("price_value") if items else None
        print(f"  amazon  : {'OK ' if ok else 'ERR'} {detail}; price={price}")
        dump[f"token{i}_amazon"] = items

        ok, detail, items = run_actor(tok, BESTBUY_ACTOR, {
            "mode": "direct_urls", "productUrls": [BESTBUY_TEST_URL], "maxProducts": 1})
        price = items[0].get("price") if items else None
        print(f"  bestbuy : {'OK ' if ok else 'ERR'} {detail}; price={price}")
        dump[f"token{i}_bestbuy"] = items
        time.sleep(1)

    if dump and os.environ.get("APIFY_TEST_DUMP"):
        with open(os.environ["APIFY_TEST_DUMP"], "w", encoding="utf-8") as f:
            json.dump(dump, f, indent=2, ensure_ascii=False)


if __name__ == "__main__":
    main()
