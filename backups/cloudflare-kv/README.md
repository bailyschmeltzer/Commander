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
- `CLOUDFLARE_API_TOKEN_RESTORE`
  - Minimum scope: Workers KV Storage Edit for the target account
- `CLOUDFLARE_ACCOUNT_ID`
- `CLOUDFLARE_KV_NAMESPACE_ID`

The workflow files are:

- `.github/workflows/backup-cloudflare-kv.yml`
- `.github/workflows/restore-cloudflare-kv.yml`

## Notes

- Keep this repository private if the synced game data should stay private.
- This backs up the cloud-synced state only. Anything still unsynced in a browser's local storage is not included.

## Restore

Use the `Restore Cloudflare KV` workflow from GitHub Actions when you need to write a saved snapshot back into Cloudflare KV.

Required inputs:

- `snapshot_path`: backup file to restore, such as `backups/cloudflare-kv/latest.json`
- `kv_key`: defaults to `pod:default:state`
- `confirm_restore`: must be exactly `RESTORE`

Recommended restore process:

1. Run the normal backup workflow first so you have a fresh snapshot of current live data.
2. Run the restore workflow with the snapshot you want.
3. Confirm the app data matches expectations.
4. Run the backup workflow again so the repo reflects the restored state.