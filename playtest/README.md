# Playtest Lab Isolation Notes

This feature is intentionally isolated from the main app runtime.

## Files
- playtest.html
- playtest/playtest.css
- playtest/playtest.js

## Separation Rules
- Does not load `app.js`.
- Does not depend on shared page scripts.
- Uses its own local storage key: `commanderPlaytestSessionV2`.
- Only reads saved deck data from `commanderDeckRecords`.
- Includes isolated debug panel, action log, smoke checklist, and session import/export.

## Safe Removal
1. Delete `playtest.html`.
2. Delete the `playtest/` folder.
3. Remove the optional link added in `decklists.html`.
