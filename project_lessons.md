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
    - **User Feedback**: Always show a success toast *after* triggering the closure to provide clear feedback.
