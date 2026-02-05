# Project Lessons & Rules

This file tracks mistakes, their root causes, and rules to prevent recurrence.
Updated automatically by the AI Assistant.

## 2026-02-04: Subject Dropdown Data Mismatch
- **Issue**: The Subject Dropdown in the Approval screen showed "undefined".
- **Cause**: Inconsistency between backend functions. `getAdminDashboardData` returned objects (`[{name, ...}]`), while `getInitData` returned strings (`["Name", ...]`). The frontend code was updated to expect objects, causing `undefined` when `getInitData` was used.
- **Root Cause**: Logic duplication. Master data retrieval logic existed in both `admin_server.js` and `user_api.js` with different implementations.
- **New Rule**: 
    1. **Single Source of Truth**: Common logic (like retrieving Master Data) MUST be centralized in a shared function (e.g., in `app_server.js` or `utils.js`). Do not duplicate logic across `admin_` and `user_` files.
    2. **Interface Consistency**: Always verify that the data structure returned by `google.script.run` exactly matches what the frontend `withSuccessHandler` expects, especially when multiple API endpoints feed the same UI component.

## 2026-02-04: Refactoring & Architecture Lessons
- **Rule**: **Avoid Monoliths**. A single file (e.g., `index.html`) over 2000 lines is unmaintainable. Even in GAS, strict modularization using `HtmlService` (`<?!= include('js'); ?>`) is mandatory for stability and developer sanity.
- **Rule**: **No Magic Strings**. Never hardcode status strings or keys (e.g., `'支払済'`, `'APPROVED'`) in Frontend or Backend logic.
    - **Solution**: Centralize all constants in `app_constants.js`, expose them via a `getAllConstants()` function, and inject them into the Global Frontend Scope (`APP_CONST`).
- **Rule**: **Regex Robustness**. When using Regex for status checks, do NOT hardcode literals inside the pattern (e.g., `/^(APPROVED|承認済み)$/`). If the Constant changes, the Regex will break silently. Construct Regex dynamically from Constants or use `Array.includes()`.
- **Rule**: **Scalability First**. Any report or calculation that iterates from "the beginning of time" (O(N)) is a ticking time bomb in GAS (6 min timeout).
    - **Solution**: Implement "Snapshotting" or "Checkpointing" for aggregated data.
- **Rule**: **Cache Invalidation Trade-offs**. When implementing cache (snapshots), ensuring consistency is harder than the caching itself.
    - **Issue**: Deleting "All Snapshots after X Date" is safe but inefficient (`O(M)` deletions). Deleting "Only Relevant Branch Snapshots" is efficient but complex (`O(1)` deletion but risk of missing dependencies).
    - **Decision**: Prioritize **Consistency** over Efficiency initially. Optimize granularity (e.g., filter by Branch) only when measurement proves it's a bottleneck.

## 2026-02-04: Concurrency & Data Integrity
- **Issue**: Standard `LockService` only protects the backend execution. It does NOT protect against "Stale Reads" (User A opens a form, User B changes it, User A saves and overwrites User B).
- **Rule**: **Optimistic Locking Required for Strict Integrity**. For critical financial records, implementing a `Version` or `UpdatedAt` check is necessary.
    - **Current State**: The app currently relies on "Last Write Wins".
    - **Mitigation**: Critical status checks (e.g., "Is Paid?") are correctly re-checked inside the `LockService` block before writing. This works for simple guards but not for complex content merges.

## 2026-02-04: Date Handling & Timezones
- **Risk**: Google Apps Script `Date` objects default to the Script Timezone (usually Pacific or owner's zone) unless configured. Mixing `new Date()` and `Utilities.formatDate` can lead to "off-by-one-day" errors.
- **Rule**: **Explicit Timezone Handling**.
    - Always define `var TIMEZONE = 'Asia/Tokyo';` as a top-level constant.
    - When parsing YYYY-MM-DD strings to Date objects for comparison, always set time explicitly (e.g., `d.setHours(0,0,0,0)`) or use `Utilities.formatDate(d, TIMEZONE, ...)` immediately.
    - **Never** rely on default string parsing (`new Date('2024-01-01')`) without verifying the script runtime timezone.

## 2026-02-04: Security & Access Control
- **Risk**: In GAS, any global function can technically be called from the frontend via `google.script.run` unless specifically properly scoped or named.
- **Rule**: **Explicit Role Checks & Naming Convention**.
    - All Admin logic MUST start with `if (user.role !== 'ADMIN') throw ...`.
    - Private helper functions should end with `_` (underscore). While GAS V8 doesn't strictly hide them from `run`, it's a strong convention, and critical logic should be wrapped in the `api_` functions that HAVE the role check.

## 2026-02-04: Frontend Error Handling & Feedback
- **Risk**: Backend errors (Script Timeout, Lock Error, Permission Denied) often result in silent failures on the UI if not handled.
- **Rule**: **Mandatory Failure Handler**.
    - Every `google.script.run` call MUST have a `.withFailureHandler(onFailure)`.
    - The `onFailure` handler MUST:
        1. Call `hideLoad()` (if a loader is active).
        2. Call `showToast('Error: ' + e.message, 'error')` or equivalent to notify the user.
        3. Log to `console.error` for debugging.

## 2026-02-04: Adversarial Review Findings (Summary)
- **Rule**: **LockService is Mandatory**. Any function that writes to the Spreadsheet (Approve, Reject, Update) MUST use `LockService` to prevent race conditions.
- **Rule**: **Avoid N+1 Queries**. Do not call `google.script.run` in a loop (e.g., for fetching images per row). Fetch all necessary data in a single batch call.
- **Rule**: **Always Push**. Run `clasp push` immediately after verifying changes to keep the GAS deployment in sync.
- **Rule**: **Always Deploy**. After pushing changes, ALWAYS update the deployment using `clasp deploy` to ensure the Web App reflects the latest code.
- **Rule**: **Adversarial Review**. Perform an Adversarial Code Review after major refactoring to identify security and reliability gaps before final delivery.

## 2026-02-04: Communication & Knowledge Policy
- **Rule**: **Japanese Output**. All user notifications, summaries, and status reports MUST be written in Japanese (日本語).
- **Rule**: **Always Update Lessons (Skill Update)**. Whenever a task or fix is successfully implemented, any new insights, successful patterns, or preventive rules MUST be added to this `project_lessons.md` file. This ensures that "skills" learned during the project are continuously codified and never forgotten.

## 2026-02-04: HTML/CSS Syntax & UI Layering
- **Issue**: Raw HTML tags displayed on screen, CSS `SyntaxError` on `querySelectorAll`, and confirmation dialogs hidden behind modals.
- **Cause**: Accidental spaces introduced during manual file repair (e.g., `< div`, `[data - file - id]`) and insufficient `z-index` coordination.
- **New Rule**:
    1. **Strict HTML Integrity**: Never allow spaces within tag brackets (e.g., `<div` NOT `< div`).
    2. **Valid Selectors**: Always double-check dynamic CSS selectors for unintended whitespace.
    3. **Hierarchical Z-Index**: Define a clear layering strategy:
        - Main Modals: 900-999
        - Confirmation Dialogs: 1100+
        - Toast Notifications: 2000+
        - Global Loading Overlay: 3000+

## 2026-02-04: Date Range Filtering Pattern
- **Pattern**: When transitioning from single-month to date range filtering:
    - **Frontend**: Use two `input type="date"` and initialize to a sensible default (e.g., 1st of current month to today).
    - **Backend**: In GAS, handle date strings by creating `Date` objects and normalizing time to `00:00:00` for the start and `23:59:59` for the end to ensure inclusive search.
    - **Summary UI**: Always reflect the selected range in the result summary label to prevent user confusion.

## 2026-02-04: Modal UI Event Flow
- **Pattern**: When a critical action (Approve, Update, Delete) is performed within a modal:
    - **Close on Success**: The modal MUST be closed automatically upon successful backend response.
    - **Refresh Context**: If the action affects the data displayed in the background screen (e.g., History list), trigger a data refresh (e.g., `loadPaymentList()`) as part of the modal closure or success handler.

## 2026-02-05: Adversarial Code Review Protocol
- **Trigger**: Before finalizing any `implementation_plan.md` or executing complex code changes.
- **Action**: The AI MUST adopt the persona of a critical "Senior Engineer (Reviewer)" and critique the proposed changes.
- **Checklist**:
    1.  **Security**: Are permissions (ADMIN vs User) strictly enforced? Is input sanitized?
    2.  **Reliability**: Are there race conditions? Is `LockService` used? Are errors handled (toast)?
    3.  **Performance**: Are there N+1 queries? Is snapshotting used for large datasets?
    4.  **Edge Cases**: What happens if the network fails? What if data is empty?
- **Goal**: Identify bugs and design flaws *before* writing code, not after.

## 2026-02-05: Deployment Persistence
- **Issue**: Running `clasp deploy` without arguments creates a NEW Deployment ID, changing the Web App URL and breaking user bookmarks.
- **Rule**: **Update Existing Deployment**.
    1.  **Check**: Identify the active Deployment ID from `deployments.txt` (or `clasp deployments`).
    2.  **Command**: Use `clasp deploy -i <EXISTING_ID> --description "..."` to overwrite the existing active deployment.
    3.  **Verify**: Confirm the Deployment ID in the output matches the expected ID.

## 2026-02-05: Debugging Skills
- **Skill**: **Detection of Duplicate Script Declarations**
    - **Symptom**: `SyntaxError: Identifier 'xyz' has already been declared`.
    - **Cause**: This often happens when `index.html` contains both an inline `<script>` block defining variables AND an external `include('js')` that defines the same variables. This usually occurs during refactoring when moving inline scripts to separate files but forgetting to remove the original inline block.
    - **Fix**: Check `index.html` for large inline script blocks and cross-reference with included `.html` or `.js` files. Remove the redundant inline block.
    - **Verification**: Ensure the error disappears in the console and the application loads.
