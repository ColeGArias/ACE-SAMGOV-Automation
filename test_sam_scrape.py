"""Throwaway diagnostic script - not part of the production automation.
Tests whether a bootstrapped anonymous session against sam.gov's internal
UI API (sgs/v1/search + opps/v2/opportunities) works without real login."""

import requests

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"

session = requests.Session()
session.headers.update({
    "User-Agent": UA,
    "Accept": "application/json, text/plain, */*",
})

print("--- Step 1a: load sam.gov search page (may set some cookies) ---")
resp = session.get("https://sam.gov/search", timeout=30)
print("status:", resp.status_code)
print("cookies after page load:", dict(session.cookies))

print("\n--- Step 1b: call the session bootstrap endpoint directly ---")
resp = session.get(
    "https://sam.gov/api/prod/iam/rms/v4/session",
    params={"random": "123456789", "api_key": "null"},
    timeout=30,
)
print("status:", resp.status_code)
print("response headers:", dict(resp.headers))
print("cookies after session call:", dict(session.cookies))
print("body (first 1500 chars):", resp.text[:1500])

token = (
    session.cookies.get("XSRF-TOKEN")
    or session.cookies.get("SESSION")
)
if not token:
    try:
        body = resp.json()
        token = body.get("sessionId") or body.get("token") or body.get("xsrfToken")
        print("token found in JSON body:", token)
    except ValueError:
        pass

if not token:
    print("\nNo usable token found from cookies or response body. Stopping.")
    raise SystemExit(1)

print("using token:", token)
session.headers.update({
    "x-auth-token": token,
    "x-xsrf-token": token,
})

print("\n--- Step 2: call the search endpoint ---")
search_params = {
    "index": "opp",
    "page": 0,
    "sort": "-modifiedDate",
    "size": 5,
    "mode": "search",
    "responseType": "json",
    "q": "",
    "qMode": "ALL",
    "modified_date.to": "2026-07-08-04:00",
    "modified_date.from": "2026-06-08-04:00",
    "naics": "23,237,2362,2373,2379",
    "notice_type": "a",
}
resp = session.get("https://sam.gov/api/prod/sgs/v1/search/", params=search_params, timeout=30)
print("status:", resp.status_code)
print("body (first 1000 chars):", resp.text[:1000])

if resp.status_code == 200:
    data = resp.json()
    results = data.get("_embedded", {}).get("results", [])
    print(f"\ngot {len(results)} results, totalElements={data.get('page', {}).get('totalElements')}")
    if results:
        first_id = results[0]["_id"]
        print("\n--- Step 3: call the detail endpoint for the first result ---")
        detail_resp = session.get(
            f"https://sam.gov/api/prod/opps/v2/opportunities/{first_id}",
            params={"api_key": "null"},
            timeout=30,
        )
        print("status:", detail_resp.status_code)
        print("body (first 1500 chars):", detail_resp.text[:1500])
