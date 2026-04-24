# Betriebshandbuch: Exchange Message Tracking GUI

Dieses Handbuch beschreibt die refaktorierte Windows-Forms-Oberfläche für
`Get-MessageTrackingLog` inklusive Multi-Server-Iteration, MessageId-Expansion
und Forensik-Export.

---

## Inhalt

1. [Zusammenfassung](#1-zusammenfassung)
2. [Voraussetzungen](#2-voraussetzungen)
3. [Installation und Aufruf](#3-installation-und-aufruf)
4. [Oberfläche im Überblick](#4-oberfläche-im-überblick)
5. [Arbeitsabläufe](#5-arbeitsabläufe)
6. [Der CSV-Fix im Detail](#6-der-csv-fix-im-detail)
7. [Forensik-Kategorien](#7-forensik-kategorien)
8. [Ausgabestruktur](#8-ausgabestruktur)
9. [Erweiterte Nutzung — Helper in der Shell](#9-erweiterte-nutzung)
10. [Troubleshooting](#10-troubleshooting)
11. [Anhang — Häufige SMTP DSN-Codes](#11-anhang--häufige-smtp-dsn-codes)

---

## 1. Zusammenfassung

Das Werkzeug ist eine Windows-Forms-basierte Oberfläche für das
Exchange-Cmdlet **Get-MessageTrackingLog**. Es sucht parallel über alle
Transport-Services eines DAG-Verbundes, flacht die Mehrwert-Eigenschaften
*Recipients*, *RecipientStatus* und *EventData* korrekt für CSV-Export ab
und ermöglicht eine tiefe forensische Analyse einzelner Nachrichten —
einschließlich Distributionslisten-Expansion, Nested-Group-Erkennung,
Cloud-Routing und Auth-Fehlern.

Gegenüber der FrankysWeb-Vorlage wurden drei wesentliche Mängel behoben:

- Der CSV-Export für Mehrwert-Eigenschaften liefert jetzt lesbare Daten statt `System.String[]`.
- Der Suchbefehl wird über Parameter-Splatting aufgebaut statt per `Invoke-Expression` (keine String-Injektion).
- Jeder Transport-Server wird einzeln mit eigener Fehlerbehandlung abgefragt — ein ausgefallener Server blockiert nicht die gesamte Suche.

Neu hinzugekommen: automatische MessageId-Expansion bei kleinen
Ergebnismengen und ein Forensik-Export, der die gesamte Transport-Reise
einer Nachricht in sechs Kategorien zerlegt und als Markdown-Report plus
Einzel-CSVs in einen Zeitstempel-Ordner schreibt.

---

## 2. Voraussetzungen

| Komponente          | Anforderung                                                                        |
|---------------------|------------------------------------------------------------------------------------|
| PowerShell          | 5.1 oder neuer (typisch: Exchange Management Shell)                                |
| Exchange-Version    | Exchange Server 2013, 2016, 2019, Subscription Edition                             |
| RBAC-Rollen         | **View-Only Recipients** plus **Message Tracking** (oder **Transport Hygiene**)    |
| Ausführungsort      | Lokale Session auf einem Exchange-Server oder EMS mit implizitem Remoting          |
| Rechte Dateisystem  | Schreibrechte im gewählten Export-Ordner                                           |

> **Hinweis:** Das Skript verwendet ausschließlich Exchange-Cmdlets und
> .NET Windows Forms. Es werden keine externen Module geladen — die
> Ausführung ist auch auf gehärteten Systemen ohne Internet-Zugang möglich.

---

## 3. Installation und Aufruf

Das Skript ist eine einzelne Datei `Exchange-MessageTracking-GUI.ps1`. Es
kann direkt aufgerufen werden oder per Dot-Source eingebunden werden, um
nur die Helper-Funktionen in der aktuellen Shell verfügbar zu machen.

### 3.1 Start der Oberfläche

```powershell
# In der Exchange Management Shell
PS> cd C:\Tools\MessageTracking
PS> .\Exchange-MessageTracking-GUI.ps1
```

### 3.2 Dot-Source für Shell-Nutzung

Beim Dot-Source (führender Punkt mit Leerzeichen) lädt das Skript die vier
Helper-Funktionen in die aktuelle Shell, **ohne** die GUI zu starten.

```powershell
# Helper laden, GUI bleibt zu
PS> . .\Exchange-MessageTracking-GUI.ps1

# Beispiel: fehlgeschlagene Mails der letzten 30 Minuten als CSV
PS> $logs = Get-TransportService | ForEach-Object {
        Get-MessageTrackingLog -Server $_.Name -Start (Get-Date).AddMinutes(-30) -ResultSize unlimited
    } | Where-Object { $_.EventId -match "fail" }

PS> $logs | ConvertTo-FlatMessageTrackingLog |
        Export-Csv .\fails.csv -Delimiter ";" -NoTypeInformation -Encoding UTF8
```

---

## 4. Oberfläche im Überblick

Die Oberfläche gliedert sich in vier Bereiche: Filter-Reihen, Zeitraum,
Options-Gruppe und Aktions-Buttons. Alle Filter sind per Checkbox
aktivierbar — leere Checkboxen werden ignoriert.

### 4.1 Filter-Eingaben

| Feld          | Parameter            | Beispiel                                                         |
|---------------|----------------------|------------------------------------------------------------------|
| Sender        | `-Sender`            | `user@example.com`  oder  `user1@example.com; user2@example.com` |
| Empfänger     | `-Recipients`        | `group@example.com`  oder  `a@example.com, b@example.com`        |
| EventID       | `-EventId`           | FAIL, DELIVER, EXPAND, REDIRECT                                  |
| MessageID     | `-MessageId`         | `<20260424120000.ABC123@mail.example.com>`                       |
| InternalMsgID | `-InternalMessageId` | 71244 (Integer, server-lokal)                                    |
| Subject       | `-MessageSubject`    | "Q2 Planung" — exakte Zeichenkette                               |
| Reference     | `-Reference`         | Parent-MessageId bei DSN/EXPAND/REDIRECT                         |
| Server(s)     | Zielserver-Liste     | leer = alle, oder: `MAIL-01, MAIL-02`                            |

> **Mehrfachwerte:** Die Felder **Sender**, **Empfänger** und **Server(s)**
> akzeptieren mehrere Werte, getrennt durch Komma oder Semikolon.
> Whitespace wird automatisch entfernt. Bei mehreren Sendern wird pro Sender
> einmal abgefragt (das Exchange-Cmdlet akzeptiert Sender nur einzeln). Bei
> mehreren Empfängern wird ein einziger Aufruf mit Array-Parameter gemacht
> (Exchange filtert dann auf "enthält einen der Empfänger").

### 4.2 Zeitraum

Start- und End-Zeit werden über native **DateTimePicker**-Controls erfasst
(Format `dd.MM.yyyy HH:mm`). Beide lassen sich per Checkbox deaktivieren —
in diesem Fall verwendet `Get-MessageTrackingLog` seine Standardwerte
(die letzten 30 Tage).

### 4.3 Optionen

| Option                                  | Wirkung                                                                                                                     |
|-----------------------------------------|-----------------------------------------------------------------------------------------------------------------------------|
| HealthMailbox-Nachrichten ausblenden    | Filtert Sender/Empfänger mit Muster "HealthMailbox" nach der Suche heraus.                                                  |
| Nur Fehler (`EventId -match "fail"`)    | Post-Filter auf FAIL-Events. Äquivalent zum Ad-hoc-Pattern in Konsolen-Skripten.                                            |
| Gruppierung nach Sender                 | Öffnet ein zweites Gridview mit Count/Name-Aggregation — nützlich bei Spam-Wellen.                                          |
| CSV-Export (flach)                      | Exportiert das Roh-Ergebnis als Semikolon-CSV in UTF-8 in den gewählten Ordner.                                             |
| Detailansicht                           | Zeigt alle Spalten im Haupt-Gridview (andernfalls nur Timestamp/EventId/Sender/Rcpts/Subject).                              |
| Empfangsbericht (DELIVER/STOREDRIVER)   | Zusätzliches Gridview mit der benutzerfreundlichen Ansicht "was wurde wann zugestellt".                                     |
| Auto-Expand, wenn Treffer ≤ N           | Wenn das Suchergebnis maximal N eindeutige MessageIds enthält, wird jede davon erneut über alle Server abgefragt.           |
| Forensik-Export                         | Erzeugt nach der Expansion einen Markdown-Report plus Detail-CSVs in einem Zeitstempel-Unterordner.                         |

### 4.4 Aktions-Buttons

| Button                  | Funktion                               | Aktiv wenn                      |
|-------------------------|----------------------------------------|---------------------------------|
| Suchen                  | Führt Suche mit aktuellen Filtern aus  | jederzeit                       |
| MessageIds expandieren  | Expandiert letztes Ergebnis manuell    | Suche lieferte Ergebnisse       |
| Forensik-Export         | Generiert Meta-Report aus Expansion    | Expansion wurde durchgeführt    |
| Schließen               | Beendet die Oberfläche                 | jederzeit                       |

---

## 5. Arbeitsabläufe

### 5.1 Standard-Suche

Der einfachste Fall — ein Benutzer fragt, ob eine bestimmte Nachricht
angekommen ist:

- Sender und/oder Empfänger setzen, Zeitraum auf ±2 Stunden um den erwarteten Zeitpunkt
- "HealthMailbox ausblenden" bleibt aktiv (Standard)
- "Suchen" — Haupt-Gridview zeigt die gefundenen Events; bei Bedarf Detailansicht aktivieren

### 5.2 Suche mit Auto-Expand (Haupt-Flow bei Tickets)

Der empfohlene Flow für Ticket-Bearbeitung. Die initiale Suche findet
Nachrichten, die den Filtern entsprechen. Bei wenigen Treffern startet die
Expansion automatisch und zeigt die vollständige Reise jeder einzelnen
Nachricht über alle Server — inklusive EXPAND, REDIRECT, TRANSFER,
DELIVER/SEND/FAIL-Events, die im Initial-Filter nicht auftauchten.

> **Warum ist das wichtig?** Eine recipient-gefilterte Suche findet nur die
> Events, in denen der gesuchte Empfänger direkt auftaucht. Wurde die
> Nachricht über eine Verteilerliste zugestellt, fehlen die EXPAND-Events
> und die Zustellung an andere Listen-Mitglieder. Die Expansion per
> MessageId zeigt den kompletten Fluss.

**Ablauf:**

- Sender oder Empfänger oder Subject-Fragment eingeben; Zeitfenster setzen
- Auto-Expand-Haken ist aktiv (Standard), Schwelle steht auf 50
- "Suchen" — bei ≤ 50 eindeutigen MessageIds startet die Expansion sofort
- Zweites Gridview zeigt die expandierten Events — spaltenorientierte Reiseansicht
- Bei gesetztem "Forensik-Export"-Haken öffnet sich danach der Markdown-Report

### 5.3 Manuelle MessageId-Expansion

Wenn die initiale Suche mehr als N Treffer liefert, wird die Auto-Expansion
übersprungen (Schutz vor versehentlich teurer Re-Query-Flut). In diesem
Fall:

- Filter enger ziehen und erneut suchen, bis die Schwelle unterschritten wird, **oder**
- Expand-Limit im Zahlenfeld hochziehen und "MessageIds expandieren" drücken

### 5.4 Forensik-Export

Nach erfolgter Expansion schreibt der Forensik-Export einen
Zeitstempel-Ordner mit einem Markdown-Report und sechs Kategorie-CSVs in
den gewählten Export-Ordner.

---

## 6. Der CSV-Fix im Detail

Die FrankysWeb-Vorlage enthielt einen doppelten Fehler beim CSV-Export.

### 6.1 Problem 1 — Scriptblock als Spalten-Literal

```powershell
# FrankysWeb Zeile 489 (fehlerhaft):
$logs | select { $_.Recipients }, { $_.RecipientStatus }, { $_.EventData }, * `
      -ExcludeProperty recipients, RecipientStatus, EventData | Export-Csv ...
```

In PowerShell erzeugt ein Scriptblock-Literal an dieser Stelle eine Spalte,
deren **Name die Scriptblock-Quelltext-Darstellung** ist. Die CSV bekommt
drei Spalten mit Überschriften `$_.Recipients`, `$_.RecipientStatus` und
`$_.EventData` statt der erwarteten Namen.

### 6.2 Problem 2 — keine Abflachung

Auch die auf den ersten Blick korrekte Fassung

```powershell
$logs | select *, @{N="Recipients"; E={ $_.Recipients }},
                  @{N="EventData";  E={ $_.EventData  }} -ExcludeProperty ...
```

erzeugt zwar korrekte Spaltennamen, übergibt aber das Array bzw. die
KeyValuePair-Liste unverändert an `Export-Csv`. Das Ergebnis:

- **Recipients**: wird als leerzeichen-getrennte Liste serialisiert — doppeldeutig wenn einzelne Werte Leerzeichen enthalten
- **RecipientStatus**: dito, und die Indizes korrespondieren nicht mehr sichtbar zu Recipients
- **EventData**: hat keine sinnvolle `ToString`-Methode — die Spalte bleibt **leer**, alle diagnostisch wertvollen Key=Value-Paare gehen verloren

### 6.3 Lösung

Die Helper-Funktion `ConvertTo-FlatMessageTrackingLog` flacht jede
Mehrwert-Eigenschaft explizit ab:

```powershell
$logs | Select-Object -ExcludeProperty Recipients,RecipientStatus,EventData -Property *,
    @{N="Recipients";      E={ $_.Recipients      -join "; " }},
    @{N="RecipientStatus"; E={ $_.RecipientStatus -join "; " }},
    @{N="EventData";       E={ ($_.EventData | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; " }}
```

Das Ergebnis ist ein lesbares, maschinenverarbeitbares CSV — alle
diagnostischen Details aus EventData bleiben erhalten.

---

## 7. Forensik-Kategorien

Der Forensik-Report zerlegt die expandierten Events in sechs Kategorien.
Jede Kategorie wird sowohl im Markdown-Report erzählt als auch als eigene
CSV abgelegt.

| Kategorie           | Erkannt über                                                   | Was steht drin                                                                                 |
|---------------------|----------------------------------------------------------------|------------------------------------------------------------------------------------------------|
| GroupExpansions     | `EventId = EXPAND`                                             | Gruppen-Adresse (`RelatedRecipientAddress`), expandierte Mitglieder, Zeitstempel, Server       |
| Redirects           | `EventId = REDIRECT`                                           | Original-Empfänger vs. Ziel-Empfänger (z.B. Forwards, Mail-Kontakte)                           |
| ExternalDeliveries  | `EventId = SEND/DELIVER`, Ziel-Domain nicht in InternalDomains | Ziel-Domains, erkannter Cloud-Provider (M365/Google/Proofpoint/Mimecast/Barracuda/Cisco)       |
| NestedChains        | Mehrere EXPANDs zur selben MessageId                           | Verschachtelung: Gruppe A → Gruppe B → Mitglieder; Tiefe der Verschachtelung                   |
| AuthFailures        | `FAIL` mit DSN `5.7.0/5.7.1/5.7.8`, `AuthRequired`, `Relay denied` | Fehlgeschlagener Empfänger, SMTP-Status, Reason, FailureCategory                            |
| DeliveryFailures    | Alle anderen FAIL-Events                                       | Fehlgeschlagener Empfänger, DSN-Code, Reason, kompletter EventData-Kontext                     |

### 7.1 Cloud-Provider-Erkennung

Die Cloud-Provider-Erkennung erfolgt über MX-Muster in den
Empfänger-Adressen:

| Provider                         | Muster                                    |
|----------------------------------|-------------------------------------------|
| Microsoft 365 / Exchange Online  | `*.mail.protection.outlook.com`           |
| Google Workspace                 | `aspmx.*google`, `googlemail`, `gmail.com`|
| Proofpoint                       | `*.pphosted.com`, `proofpoint`            |
| Mimecast                         | `mimecast`                                |
| Barracuda                        | `barracuda`                               |
| Cisco IronPort                   | `cisco.iphmx`                             |
| External SMTP                    | alle anderen externen Domains             |

### 7.2 Interne Domains

Als "intern" gelten standardmäßig die Domains aller Sender im Datensatz.
Wird der forensische Report programmatisch über
`ConvertTo-MessageTrackingForensics` erzeugt, kann die Liste explizit
übergeben werden:

```powershell
$expanded | ConvertTo-MessageTrackingForensics `
    -InternalDomains "example.com","subsidiary.example.com","example.net"
```

---

## 8. Ausgabestruktur

Alle Exporte landen im Ordner, der im Feld "Export-Ordner" ausgewählt ist.
Unterordner werden automatisch erzeugt.

### 8.1 CSV-Export (Rohsuche)

```
C:\Temp\
└── 2026.04.24_14.23.15_sender-user1_example.com.csv
    └── geflachte Tracking-Logs der initialen Suche
```

Der Dateiname enthält einen Zeitstempel und einen Filter-Hinweis
(Sender/Recipient/msgid/fail). Sonderzeichen werden durch Unterstrich
ersetzt.

### 8.2 Forensik-Export

```
C:\Temp\
└── 2026.04.24_14.23.15_tracking_forensics\
    ├── 00_forensics_report.md        # narrativer Markdown-Report
    ├── 01_group_expansions.csv        # EXPAND-Events
    ├── 02_redirects.csv               # REDIRECT-Events
    ├── 03_external_deliveries.csv     # SEND/DELIVER an externe Domains
    ├── 04_nested_chains.csv           # Verschachtelte Gruppen
    ├── 05_auth_failures.csv           # 5.7.x / AuthRequired
    ├── 06_delivery_failures.csv       # alle anderen FAILs
    └── 99_expanded_tracking_flat.csv  # Roh-Expansion, geflacht
```

Die Datei `99_expanded_tracking_flat.csv` ist die vollständige
Roh-Expansion und eignet sich für Excel-Pivots oder weitergehende Analysen.

---

## 9. Erweiterte Nutzung

Nach Dot-Source des Skripts stehen vier öffentliche Funktionen zur
Verfügung, die auch außerhalb der GUI nutzbar sind.

### 9.1 ConvertTo-FlatMessageTrackingLog

Flacht Recipients, RecipientStatus und EventData ab. Pipeline-fähig.

```powershell
PS> $logs | ConvertTo-FlatMessageTrackingLog |
        Export-Csv .\tracking.csv -Delimiter ";" -NoTypeInformation -Encoding UTF8
```

### 9.2 Invoke-MessageTrackingSearch

Führt `Get-MessageTrackingLog` über eine Server-Liste aus. Bei leerer Liste
werden alle Transport-Services automatisch ermittelt.

```powershell
PS> $logs = Invoke-MessageTrackingSearch -Parameters @{
        Sender     = "monitor@example.com"
        Start      = (Get-Date).AddHours(-2)
        End        = (Get-Date)
        ResultSize = "unlimited"
    } -ExcludeHealth -FailuresOnly

# Multi-Sender: einmal pro Sender iterieren
PS> $logs = Invoke-MessageTrackingSearch -Parameters @{
        Sender     = @("a@example.com","b@example.com","c@example.com")
        Start      = (Get-Date).AddHours(-2)
        ResultSize = "unlimited"
    }
```

### 9.3 Expand-MessageTrackingByMessageId

Nimmt ein Ergebnis-Set und fragt für jede eindeutige MessageId die volle
Transport-Reise ab. Default-Limit: 50 MessageIds (Schutz vor versehentlich
teuren Re-Queries).

```powershell
PS> $expanded = Expand-MessageTrackingByMessageId `
        -Logs  $logs `
        -Limit 100 `
        -Start (Get-Date).AddHours(-3) `
        -End   (Get-Date)
```

### 9.4 ConvertTo-MessageTrackingForensics

Klassifiziert expandierte Logs in die sechs Forensik-Kategorien.

```powershell
PS> $forensics = $expanded | ConvertTo-MessageTrackingForensics `
        -InternalDomains "example.com","subsidiary.example.com"

PS> $forensics.Summary.Counts
PS> $forensics.AuthFailures | Format-Table

# Report als Markdown speichern
PS> Format-MessageTrackingForensicsReport -Forensics $forensics |
        Set-Content .\report.md -Encoding UTF8
```

---

## 10. Troubleshooting

| Symptom                                                            | Ursache und Behebung                                                                                             |
|--------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------|
| "Get-MessageTrackingLog: Die Benennung ... wurde nicht erkannt"    | Die Exchange Management Shell ist nicht geladen. Skript muss in der EMS laufen.                                  |
| "Get-TransportService schlägt fehl: Zugriff verweigert"            | RBAC-Rollen fehlen. Admin-Konto benötigt `View-Only Recipients` plus `Message Tracking`.                         |
| `Server <Name> schlägt fehl: Kennwort abgelaufen / WinRM-Fehler`    | Warnung wird pro Server ausgegeben, andere Server werden weiterhin abgefragt. Betroffenen Server prüfen.         |
| Auto-Expand übersprungen — über Schwelle                            | Initiale Suche lieferte mehr als N eindeutige MessageIds. Filter verengen oder Schwelle hochziehen.              |
| Forensik-Button deaktiviert                                         | Keine Expansion vorhanden. Zuerst "MessageIds expandieren" ausführen oder Auto-Expand aktivieren.                |
| CSV enthält leere EventData-Spalte                                  | Das Rohobjekt hatte kein EventData (typisch bei RECEIVE ohne Transport-Agents). Kein Fehler.                     |
| `Out-GridView` öffnet sich nicht                                    | PowerShell ISE oder Constrained Language Mode. Skript in normaler PowerShell-Konsole ausführen.                  |

---

## 11. Anhang — Häufige SMTP DSN-Codes

Zur Interpretation der Spalte **Status** im Forensik-Report. Die DSN-Klasse
`5.x.x` bedeutet dauerhafter Fehler, `4.x.x` vorläufiger Fehler (Retry).

| DSN    | Bedeutung                              | Typischer Kontext                                             |
|--------|----------------------------------------|---------------------------------------------------------------|
| 5.1.1  | Empfänger unbekannt                    | Bad mailbox address — Empfänger existiert nicht               |
| 5.1.2  | Bad hostname                           | Fehlgeleitet oder DNS-Problem                                 |
| 5.1.10 | Empfänger-Adresse verweigert           | EOP/Defender: Adresse geblockt oder ungültig                  |
| 5.2.2  | Mailbox voll                           | Quota überschritten — empfängerseitig                         |
| 5.3.4  | Nachricht zu groß                      | Size-Limit des Connectors oder Empfängers                     |
| 5.4.4  | Routing-Fehler                         | Kein Pfad zum Ziel (MX/Connector fehlt)                       |
| 5.4.6  | Routing-Schleife                       | Mail-Loop — Konfigurationsfehler                              |
| 5.4.7  | Verzögerungs-Timeout                   | Nach Defers endgültig aufgegeben                              |
| 5.7.0  | Allgemeiner Sicherheitsfehler          | Transport-Rule, Connector-TLS                                 |
| 5.7.1  | Zustellung nicht autorisiert           | **Auth-Fehler** — Relay abgelehnt, nicht authentifiziert      |
| 5.7.5  | Kryptografischer Fehler                | TLS-Zertifikat oder -Protokoll inkompatibel                   |
| 5.7.8  | Authentifizierungsdaten ungültig       | **Auth-Fehler** — falsches Kennwort / AuthMech                |
| 5.7.9  | Authentifizierung erforderlich         | **Auth-Fehler** — Pre-Auth Submission                         |
| 4.4.7  | Nachricht abgelaufen                   | In Queue, wiederholte Retries, jetzt DSN                      |
| 4.7.26 | Strenge Authentifizierung erforderlich | SPF/DKIM/DMARC Enforcement (EOP)                              |

Die Kategorie **AuthFailures** im Forensik-Report umfasst `5.7.0`, `5.7.1`,
`5.7.8` sowie Events mit `Reason=AuthRequired`, `NotAuthenticated`,
`Client was not authenticated` oder `RelayDenied` im EventData. Alle
anderen 5.x.x- und 4.x.x-Fehler landen in **DeliveryFailures** mit
vollständigem EventData-Kontext für die tiefere Analyse.
