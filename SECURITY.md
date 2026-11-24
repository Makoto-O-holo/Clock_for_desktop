# Security policy

## Reporting a vulnerability

Please do not report security issues in public issues. Use GitHub's private vulnerability reporting for this repository (Security → Advisories → Report a vulnerability). Include the affected version, impact, and steps to reproduce. Do not include personal data in reports.

## Scope

The app is a local Windows clock. It does not intentionally make network requests. Settings are stored in the current user's `%APPDATA%\Clock` directory. Startup registration creates a shortcut only after the user selects the corresponding context menu item.
