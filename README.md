# Exchange Message Tracking GUI

A Windows Forms frontend for `Get-MessageTrackingLog` with proper CSV flattening,
multi-server iteration, and distribution-list expansion forensics.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![Exchange](https://img.shields.io/badge/Exchange-2016%2F2019%2FSE-red)
![License](https://img.shields.io/badge/License-MIT-green)

## Why

The classic FrankysWeb Exchange Message Tracking GUI from 2019 has a subtle but
painful CSV-export bug that silently drops every bit of diagnostic detail from
`Recipients`, `RecipientStatus`, and `EventData`. This refactor fixes that and
adds the features that matter for incident response:

- **Proper CSV output** — recipients and event-data survive the export instead
  of serializing as `System.String[]` or empty cells.
- **Multi-server iteration** — queries every Transport Service in parallel with
  per-server error handling. One dead server no longer blocks the search.
- **Multi-sender / multi-recipient** filters — comma or semicolon separated.
- **MessageId auto-expansion** — when the initial result set is small, each
  unique MessageId is re-queried across all servers to reveal the full
  transport journey: `RECEIVE → EXPAND → TRANSFER → REDIRECT → DELIVER/FAIL`.
- **Forensics meta-export** — classifies the expanded events into six
  categories (group expansions, redirects, external deliveries with cloud
  provider detection, nested group chains, auth failures, other delivery
  failures) and writes a narrative Markdown report plus per-category CSVs.

## Quick start

Run from an Exchange Management Shell on any server with mail-tracking RBAC:

```powershell
PS> .\Exchange-MessageTracking-GUI.ps1
```

Or dot-source to get just the helper functions in your current shell:

```powershell
PS> . .\Exchange-MessageTracking-GUI.ps1
PS> Get-TransportService | ForEach-Object {
        Get-MessageTrackingLog -Server $_.Name -Start (Get-Date).AddHours(-1) -ResultSize unlimited
    } | ConvertTo-FlatMessageTrackingLog |
        Export-Csv .\tracking.csv -Delimiter ';' -NoTypeInformation -Encoding UTF8
```

## What the GUI does

1. Enter filters (sender, recipient, subject, etc. — each optional via checkbox).
2. Click **Suchen**. Results appear in an `Out-GridView` window, CSV export is optional.
3. If the result contains ≤ N unique MessageIds (default 50), each one is
   automatically re-queried across all servers — giving you the full transport
   journey, not just the events where your recipient filter matched.
4. Click **Forensik-Export** to get a timestamped folder with a Markdown
   narrative and six category CSVs.

## Requirements

| Component          | Requirement                                                     |
|--------------------|-----------------------------------------------------------------|
| PowerShell         | 5.1+ (typically the Exchange Management Shell)                  |
| Exchange Server    | 2013, 2016, 2019, or Subscription Edition                       |
| RBAC roles         | `View-Only Recipients` + `Message Tracking` (or `Transport Hygiene`) |
| Runtime            | Local session on an Exchange server, or EMS with implicit remoting |

No external modules, no internet access required.

## Public API (dot-sourced)

Four helpers become available when you dot-source the script:

| Function                             | Purpose                                                                 |
|--------------------------------------|-------------------------------------------------------------------------|
| `ConvertTo-FlatMessageTrackingLog`   | Flattens `Recipients` / `RecipientStatus` / `EventData` for CSV export. |
| `Invoke-MessageTrackingSearch`       | Runs `Get-MessageTrackingLog` across a server list with error handling. |
| `Expand-MessageTrackingByMessageId`  | Re-queries each unique MessageId across all servers.                    |
| `ConvertTo-MessageTrackingForensics` | Classifies expanded logs into six forensic categories.                  |

See [`MANUAL.md`](MANUAL.md) for the full German operator's manual covering
workflows, the CSV-fix detail, forensic categories, output structure, and an
SMTP DSN-code reference.

## Credits

Inspired by the original [FrankysWeb Exchange Message Tracking GUI](https://www.frankysweb.de/)
(PowerShell Studio, 2019). This is a full rewrite — no SAPIEN-generated code,
proper parameter splatting instead of `Invoke-Expression`, and the forensic
features are new.

## License

MIT — see [`LICENSE`](LICENSE).
