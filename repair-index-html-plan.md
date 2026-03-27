# Repair Plan: index.html stabilization

## Scope
- Fix structural issues in the HTML document.
- Harden startup and Apps Script-dependent flows against missing runtime or malformed local state.
- Refresh the active view reliably after data reloads so edits are visible immediately.

## Checkpoints
1. Remove premature `</body>` / `</html>` so all scripts live inside the document.
2. Add safe helpers for Apps Script runtime detection and localStorage JSON parsing.
3. Improve `handleLogin`, `loadInitialData`, `loadAdminData`, `processSync`, and documentation loading error handling.
4. Re-run a syntax check and verify the document has only one closing `</body>` / `</html>` pair.
