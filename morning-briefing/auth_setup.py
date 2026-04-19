"""
One-time device-code auth to obtain a Microsoft Graph refresh token.

Run locally (NOT in CI):
    pip install -r requirements.txt
    export AZURE_CLIENT_ID=...
    export AZURE_TENANT_ID=...
    python auth_setup.py

Follow the printed URL + device code, sign in with your work account, then
copy the printed refresh token into the GitHub Actions secret MS_REFRESH_TOKEN.
"""
import os
import sys

import msal

SCOPES = [
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Tasks.ReadWrite",
    "User.Read",
]


def main() -> int:
    client_id = os.environ.get("AZURE_CLIENT_ID")
    tenant_id = os.environ.get("AZURE_TENANT_ID")
    if not client_id or not tenant_id:
        print("ERROR: set AZURE_CLIENT_ID and AZURE_TENANT_ID env vars first.", file=sys.stderr)
        return 1

    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
    )

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print(f"ERROR: failed to start device flow: {flow}", file=sys.stderr)
        return 1

    print(flow["message"])
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        print(f"ERROR: auth failed: {result.get('error_description', result)}", file=sys.stderr)
        return 1

    cache = app.token_cache.serialize()
    import json
    parsed = json.loads(cache)
    refresh_tokens = parsed.get("RefreshToken", {})
    if not refresh_tokens:
        print("ERROR: no refresh token returned. Did you grant offline_access?", file=sys.stderr)
        return 1

    rt_entry = next(iter(refresh_tokens.values()))
    refresh_token = rt_entry["secret"]

    print()
    print("=" * 60)
    print("SUCCESS")
    print("=" * 60)
    print(f"Signed in as: {result.get('id_token_claims', {}).get('preferred_username', 'unknown')}")
    print()
    print("Refresh token (copy this into GitHub secret MS_REFRESH_TOKEN):")
    print()
    print(refresh_token)
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
