# LPR Registration System (Google Apps Script)

A production-ready registration and check-in system built with Google Apps Script and Google Sheets.

## Overview

This project manages:

- Student registrations
- Trimester subscriptions
- Class cards (6 / 12)
- Multi-class drop-ins
- Door check-in lists
- Payment logging
- Owed / credit balance tracking
- Context-based backdating for historical entries

Designed for high-speed real-world use at class entrances.

---

## Key Architecture Concepts

### Context Engine
All actions use a session context date:
- ISSUE (card / trimester)
- SPEND / ATTEND
- DROP-IN
- DOORLIST validation

This allows full historical backdating while preserving audit timestamps.

### Idempotency
Each submission includes a `submission_id` to prevent duplicate writes.

### Multi-Class Drop-ins
Drop-ins support multiple classes per session.
Expected price auto-calculates as:

12€ × number_of_classes

Balance calculations reflect real expected totals.

### Door List Engine
- Zebra rows
- Partial validation indicators
- Explicit class-level validation
- Smart enable/disable logic

---

## Tech Stack

- Google Apps Script (backend)
- HTMLService frontend
- Google Sheets datastore
- Vanilla JS (no external framework)

---

## Deployment

1. Create a Google Apps Script project bound to a Google Sheet.
2. Paste:
   - `apps-script/Code.js`
   - `apps-script/page.html`
3. Deploy as Web App.
4. Set your own deployment URL in configuration.

> This public repository excludes private deployment URLs and operational credentials.

---

## Version

Current benchmark release: **v4.5.0**

---

## Roadmap

- Reduce doorlist POST overhead (batching)
- Financial report: users who owe / are owed
- UI refinements for mobile
- Email template improvements

---

## License

MIT
