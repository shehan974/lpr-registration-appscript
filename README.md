# LPR Registration System (Google Apps Script)

Production-ready registration and operational check-in system built with Google Apps Script and Google Sheets.

## Overview

This system manages:

- Student registrations
- Trimester subscriptions
- Class cards (6 / 12 sessions)
- Multi-class drop-ins
- Real-time door check-in lists
- Payment logging
- Owed / credit balance tracking
- Controlled historical backdating

Designed for high-speed real-world usage in live class environments, where operational reliability and fast data validation are critical.

---

## System Design & Architecture

### Context-Based Processing Engine

All transactions operate using a controlled session context date to allow:

- ISSUE (card / trimester)
- ATTEND / SPEND
- DROP-IN
- Doorlist validation

This enables historical backdating while preserving immutable audit timestamps for traceability.

### Idempotent Submissions

Each transaction includes a unique `submission_id` to prevent duplicate writes and ensure data integrity during concurrent usage.

### Multi-Class Drop-In Logic

Supports multiple classes per session.

Dynamic pricing calculation based on session count.
Balance tracking reflects real-time expected totals.

### Doorlist Validation Engine

- Zebra-row rendering for visual clarity
- Partial validation indicators
- Explicit class-level validation logic
- Conditional enable/disable controls
- Optimised for fast entrance throughput

---

## Technical Stack

- Google Apps Script (server-side logic)
- HTMLService frontend
- Google Sheets datastore (structured transactional model)
- Vanilla JavaScript (no external frameworks)

---

## Operational Characteristics

- Designed for production usage
- Supports concurrent form submissions
- Minimises manual intervention
- Reduces administrative overhead
- Ensures auditability of financial records
- Improves entrance flow efficiency

---

## Deployment

1. Create a Google Apps Script project bound to a Google Sheet.
2. Add:
   - `apps-script/Code.js`
   - `apps-script/page.html`
3. Deploy as Web App.
4. Configure deployment URL in script properties.

> This public repository excludes private deployment URLs, sheet IDs and operational credentials.

---

## Version

Current benchmark release: **v4.5.0**

---

## Roadmap

- Batch optimisation for doorlist POST operations
- Financial reporting dashboard (owed / credit reconciliation)
- Mobile UI refinement
- Notification and email automation improvements

---

## License

MIT
