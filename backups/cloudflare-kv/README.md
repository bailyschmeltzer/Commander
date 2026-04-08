# Cloudflare KV Backups

This folder stores automated snapshots of the synced pod state from Cloudflare KV.

## What gets backed up

- KV namespace entry: `pod:default:state`
- Output files:
  - `latest.json`: most recent snapshot
  - `history/<timestamp>.json`: timestamped snapshots only when the value changes

## GitHub setup

Add these repository secrets before running the workflow:

- `CLOUDFLARE_API_TOKEN`
  - Minimum scope: Workers KV Storage Read for the target account
- `CLOUDFLARE_ACCOUNT_ID`
- `CLOUDFLARE_KV_NAMESPACE_ID`

The workflow file is in `.github/workflows/backup-cloudflare-kv.yml`.

## Notes

- Keep this repository private if the synced game data should stay private.
- This backs up the cloud-synced state only. Anything still unsynced in a browser's local storage is not included.
- To restore, copy a snapshot back into the same Cloudflare KV key.