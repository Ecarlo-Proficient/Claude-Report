# docker/ - how the office server works

Last changed: 09/30/2026 - the README is plain setup instructions for the developer (no approval stops). Earlier
the same day: NEW, the office server package (replaced the retired invoice-sync-only container).

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
    AR --> LIVE
    CLOCK --> STATUS
```
