# docker/ - how the office server works

Last changed: 10/07/2026 - **payments-test is OFF** (`PAYMENTS_TEST=1` turns it on) until the base AP/AR run is proven on the server. When on, test mode adds it after AR: `invoice-sync/qbo_payment_notify.py --test` posts "TEST - QuickBooks payment received" cards to TEAMS_WEBHOOK_PAYMENTS_TEST (its own posted record), or logs what it would post while that key is blank - the team's Payments channel is proven here before go-live. Earlier: 10/06/2026 - the server's Bill Tracker is no longer locked to its own account (no chmod 600 in server mode); the Mac's sync-all stands down on AP + AR while writer.json says server. Earlier: 10/05/2026 - STATUS only (test week on the Synology; two open side notes). Earlier, 10/02/2026: the mirror reads MIRROR_KEY from secrets.env (qbo_vault fix). Same day: compose.yml builds as uid 1027 (the Synology's svc-automation user). Earlier, 10/01/2026: first build: the Dockerfile reuses GID 100 (Synology's users group); the server's QBO
login check needs only the four QBO keys. The flow itself is unchanged.

Update this chart - and the line above - in the same commit as any change to `docker/`
(`.github/flow_guard.sh` fails the push otherwise).

```mermaid
flowchart TD
    classDef src fill:#eef4ff,stroke:#5b7fc7,color:#111
    classDef tool fill:#f4f4f4,stroke:#777,color:#111
    classDef gate fill:#fff3e0,stroke:#d9822b,color:#111
    classDef out fill:#eaf7ea,stroke:#3c8d40,color:#111
    classDef stop fill:#fdecec,stroke:#c0392b,color:#111

    QBO[("QuickBooks Online<br/>the server's OWN Intuit app")]:::src
    KEYS[("automation-keys<br/>encrypted shared folder")]:::src
    MAC[("Owner's Mac<br/>push_registers.sh -> the rule files")]:::src
    FILES[("Synology shares<br/>General List · PO tracker (Cloud Sync copy) · Accounting")]:::src

    CLOCK["scheduler.py<br/>the one process"]:::tool
    MIR["mirror refresh · every 3 min<br/>nightly count check"]:::tool
    STALE{"mirror refreshed in the<br/>last 30 min?"}:::gate
    MODE{"ACB_SERVER_MODE"}:::gate
    WRITER{"writer file says<br/>server?"}:::gate
    AP["AP · bill-tracker"]:::tool
    AR["AR · invoice-sync"]:::tool
    TEST["test: trackers to _server-test/<br/>AR dry run · Invoice Tracker rebuilt read-only"]:::out
    PTEST["payments-test - OFF until PAYMENTS_TEST=1<br/>TEST QuickBooks payment cards -> Payments channel<br/>(dry run log while no test webhook)"]:::out
    LIVE["live: Bill Tracker + Invoice Tracker<br/>Notion + Teams"]:::out
    DOWN["stand down<br/>(the Mac's sync-all writes)"]:::stop
    STATUS["status.json + Teams alerts<br/>(a failed run · an hour with no good refresh)"]:::out

    KEYS --> CLOCK
    QBO --> MIR
    CLOCK --> MIR --> STALE
    STALE -- no --> STATUS
    STALE -- yes --> MODE
    MODE -- test --> AP
    MODE -- live --> WRITER
    WRITER -- no / missing --> DOWN
    WRITER -- yes --> AP
    MAC --> AP
    FILES --> AP
    AP --> AR
    AR --> TEST
    AR --> PTEST
    AR --> LIVE
    CLOCK --> STATUS
```
