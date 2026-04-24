#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange Message Tracking GUI (refaktorierte Fassung, Windows Forms)

.DESCRIPTION
    Oberflaeche fuer Get-MessageTrackingLog. Sucht ueber alle
    Get-TransportService-Server (oder eine explizit angegebene, kommagetrennte
    Liste). Exportiert Recipients, RecipientStatus und EventData korrekt
    abgeflacht als CSV - das ist der Kern-Fix gegenueber der FrankysWeb-Vorlage,
    in der die drei Mehrwert-Properties als "System.String[]" oder leerer
    Blob in der CSV landeten.

    Zusaetzlich gegenueber dem Original:
      - Splatting statt Invoke-Expression (keine String-Injektion)
      - Multi-Server-Iteration mit Per-Server-Fehlerbehandlung
      - Schnellfilter "nur Fehler" (EventId -match 'fail')
      - Gruppierung nach Sender als Schnellauswertung
      - Automatische, mit Zeitstempel versehene Dateinamen
      - Semikolon-Delimiter + UTF8 (deutsche Excel-Defaults)

.NOTES
    Voraussetzung: Exchange Management Shell (lokal), passende RBAC-Rollen
    fuer Get-MessageTrackingLog und Get-TransportService.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =============================================================================
# Helpers
# =============================================================================

function ConvertFrom-EventDataList {
    <#
    .SYNOPSIS
        Wandelt eine EventData-Name/Value-Liste in ein PSCustomObject um.
    .EXAMPLE
        $log.EventData | ConvertFrom-EventDataList
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)] $List
    )
    begin { $h = [ordered]@{} }
    process {
        foreach ($nv in $List) { $h[$nv.Key] = $nv.Value }
    }
    end { [pscustomobject]$h }
}

function ConvertTo-FlatMessageTrackingLog {
    <#
    .SYNOPSIS
        Der Fix. Abflachen von Recipients / RecipientStatus / EventData
        fuer verlustfreien CSV- und Gridview-Export.

    .DESCRIPTION
        Die FrankysWeb-Vorlage nutzt `select { $_.Recipients }, ... *` mit
        Scriptblock-Literalen - die CSV bekommt dann Spalten wie
        "$_.Recipients" und darin serialisierten Muell.
        Auch der Naive-Fix `@{N='Recipients';E={$_.Recipients}}` reicht nicht:
        das Array wird unveraendert uebergeben und landet als
        "System.String[]" in der CSV.

        Hier wird jede Mehrwert-Property abgeflacht:
          - Recipients      -> '; '-getrennt
          - RecipientStatus -> '; '-getrennt (Index korrespondiert zu Recipients)
          - EventData       -> 'Key1=Val1; Key2=Val2'

    .EXAMPLE
        $logs | ConvertTo-FlatMessageTrackingLog |
            Export-Csv .\tracking.csv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)] $InputObject
    )
    process {
        foreach ($log in $InputObject) {
            $log | Select-Object -ExcludeProperty Recipients, RecipientStatus, EventData -Property *,
                @{N = 'Recipients'; E = {
                    if ($_.Recipients) { ($_.Recipients) -join '; ' } else { '' }
                }},
                @{N = 'RecipientStatus'; E = {
                    if ($_.RecipientStatus) { ($_.RecipientStatus) -join '; ' } else { '' }
                }},
                @{N = 'EventData'; E = {
                    if ($_.EventData) {
                        ($_.EventData | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                    } else { '' }
                }}
        }
    }
}

function Invoke-MessageTrackingSearch {
    <#
    .SYNOPSIS
        Fuehrt Get-MessageTrackingLog ueber eine Serverliste aus (oder alle).
    .NOTES
        Kernel fuer die GUI - kann aber auch eigenstaendig genutzt werden.
    #>
    [CmdletBinding()]
    param(
        [string[]]  $ServerList,
        [hashtable] $Parameters = @{},
        [switch]    $ExcludeHealth,
        [switch]    $FailuresOnly
    )

    if (-not $ServerList -or $ServerList.Count -eq 0) {
        try {
            $ServerList = (Get-TransportService -ErrorAction Stop).Name
        } catch {
            throw "Get-TransportService schlaegt fehl: $($_.Exception.Message)"
        }
    }

    # Get-MessageTrackingLog kennt -Sender nur als Einzelwert, -Recipients
    # dagegen als Array. Bei einer Sender-Liste wird pro Sender iteriert.
    $senderList = $null
    if ($Parameters.ContainsKey('Sender') -and @($Parameters.Sender).Count -gt 1) {
        $senderList = @($Parameters.Sender)
        $Parameters  = $Parameters.Clone()
        $Parameters.Remove('Sender') | Out-Null
    }

    $all = foreach ($srv in $ServerList) {
        try {
            if ($senderList) {
                foreach ($snd in $senderList) {
                    Get-MessageTrackingLog -Server $srv @Parameters -Sender $snd -ErrorAction Stop
                }
            } else {
                Get-MessageTrackingLog -Server $srv @Parameters -ErrorAction Stop
            }
        } catch {
            Write-Warning "Server ${srv} schlaegt fehl: $($_.Exception.Message)"
        }
    }

    if ($ExcludeHealth) {
        $all = $all | Where-Object {
            $_.Sender -notmatch 'HealthMailbox' -and
            (($_.Recipients -join ',') -notmatch 'HealthMailbox')
        }
    }

    if ($FailuresOnly) {
        $all = $all | Where-Object { $_.EventId -match 'fail' }
    }

    , $all
}

function Expand-MessageTrackingByMessageId {
    <#
    .SYNOPSIS
        Nimmt ein Ergebnis-Set und erweitert jede eindeutige MessageId zur
        vollstaendigen Transport-Reise ueber alle Server.

    .DESCRIPTION
        Eine recipient-gefilterte Suche sieht nur die Slice, in der dieser
        Recipient auftaucht. Expand re-queried jede MessageId ohne
        Recipient-Filter auf allen Transport-Services - dadurch werden
        EXPAND (DL-Aufloesung), REDIRECT (Weiterleitungen), TRANSFER
        (Bifurkation) und saemtliche DELIVER/SEND/FAIL-Events sichtbar,
        auch wenn sie auf anderen Servern stattfanden.

    .PARAMETER Logs
        Ergebnis einer vorherigen Tracking-Suche.

    .PARAMETER Limit
        Maximale Anzahl eindeutiger MessageIds, die expandiert werden darf.
        Default: 50. Schutz vor versehentlicher Komplett-Expansion.

    .PARAMETER ServerList
        Optional: explizite Serverliste. Leer = alle Transport-Services.

    .PARAMETER Start
    .PARAMETER End
        Zeitfenster der erneuten Abfrage (sollte das urspruengliche Fenster
        sein oder leicht ausgedehnt, damit vor- und nachgelagerte Events
        gesehen werden).

    .PARAMETER Progress
        Optional: Scriptblock, der pro Iteration mit (current,total,messageId)
        aufgerufen wird. Wird von der GUI genutzt.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]          $Logs,
        [int]      $Limit      = 50,
        [string[]] $ServerList,
        [datetime] $Start,
        [datetime] $End,
        [scriptblock] $Progress
    )

    $mids = $Logs |
        Where-Object { $_.MessageId } |
        Select-Object -ExpandProperty MessageId -Unique

    if (-not $mids -or $mids.Count -eq 0) {
        Write-Warning 'Keine MessageIds im Ergebnis - Expansion uebersprungen.'
        return , @()
    }

    if ($mids.Count -gt $Limit) {
        Write-Warning ("MessageId-Anzahl ({0}) ueberschreitet Limit ({1}) - Expansion uebersprungen." -f $mids.Count, $Limit)
        return , @()
    }

    $expandedParams = @{ ResultSize = 'unlimited' }
    if ($PSBoundParameters.ContainsKey('Start')) { $expandedParams.Start = $Start }
    if ($PSBoundParameters.ContainsKey('End'))   { $expandedParams.End   = $End   }

    $i     = 0
    $total = $mids.Count
    $all   = foreach ($mid in $mids) {
        $i++
        if ($Progress) { & $Progress $i $total $mid }
        $p = $expandedParams.Clone()
        $p.MessageId = $mid
        Invoke-MessageTrackingSearch -ServerList $ServerList -Parameters $p
    }

    , @($all)
}

function ConvertTo-MessageTrackingForensics {
    <#
    .SYNOPSIS
        Zerlegt expandierte Tracking-Logs in forensische Kategorien.

    .DESCRIPTION
        Erkennt und klassifiziert:
          - GroupExpansions   : EXPAND-Events (DL-Aufloesung inkl.
                                RelatedRecipientAddress = Gruppenadresse)
          - Redirects         : REDIRECT-Events (OriginalRecipient vs. Ziel,
                                typisch fuer Weiterleitungen / Mail-Kontakte)
          - ExternalDeliveries: SEND/DELIVER an externe Domaenen und
                                Cloud-Ziele (M365, Google, Proofpoint etc.)
          - NestedChains      : Ketten von EXPANDs, in denen Mitglieder
                                selbst wieder DLs sind
          - AuthFailures      : FAILs mit DSN 5.7.0/5.7.1/5.7.8 bzw.
                                Reason=AuthRequired / Relay denied
          - DeliveryFailures  : sonstige FAIL-Events mit Grund

    .PARAMETER Logs
        Expandierte Tracking-Logs (Roh-Objekte, nicht geflacht).

    .PARAMETER InternalDomains
        Optional: Liste der als intern geltenden Domaenen. Wenn leer,
        werden Sender-Domaenen automatisch als intern herangezogen.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true)] $Logs,
        [string[]] $InternalDomains
    )

    begin { $all = @() }
    process { $all += $Logs }
    end {
        if (-not $InternalDomains -or $InternalDomains.Count -eq 0) {
            $InternalDomains = $all |
                Where-Object Sender |
                ForEach-Object { ($_.Sender -split '@')[-1].ToLower() } |
                Sort-Object -Unique
        } else {
            $InternalDomains = $InternalDomains | ForEach-Object { $_.ToLower() }
        }

        $getEvtField = {
            param($log, $key)
            if (-not $log.EventData) { return $null }
            ($log.EventData | Where-Object { $_.Key -eq $key } | Select-Object -First 1).Value
        }

        # --- EXPAND events ---
        $expandEvents = $all | Where-Object { $_.EventId -eq 'EXPAND' }
        $groupExpansions = foreach ($e in $expandEvents) {
            $groupAddr = $e.RelatedRecipientAddress
            if (-not $groupAddr) { $groupAddr = & $getEvtField $e 'RelatedRecipientAddress' }
            [pscustomobject]@{
                Timestamp    = $e.Timestamp
                Sender       = $e.Sender
                GroupAddress = $groupAddr
                MemberCount  = @($e.Recipients).Count
                Members      = ($e.Recipients -join '; ')
                Subject      = $e.MessageSubject
                MessageId    = $e.MessageId
                Source       = $e.Source
                Server       = $e.ServerHostname
            }
        }

        # --- REDIRECT events ---
        $redirectEvents = $all | Where-Object { $_.EventId -eq 'REDIRECT' }
        $redirects = foreach ($r in $redirectEvents) {
            $orig = $r.RelatedRecipientAddress
            if (-not $orig) { $orig = & $getEvtField $r 'OriginalRecipientAddress' }
            if (-not $orig) { $orig = & $getEvtField $r 'RelatedRecipientAddress' }
            [pscustomobject]@{
                Timestamp       = $r.Timestamp
                Sender          = $r.Sender
                OriginalRcpt    = $orig
                RedirectedTo    = ($r.Recipients -join '; ')
                Subject         = $r.MessageSubject
                MessageId       = $r.MessageId
                Server          = $r.ServerHostname
            }
        }

        # --- External deliveries (SEND/DELIVER with non-internal recipients) ---
        $cloudPattern = 'mail\.protection\.outlook\.com|aspmx\..*google|googlemail\.com|pphosted\.com|proofpoint|mimecast|barracuda|cisco\.iphmx'
        $externals = foreach ($log in $all | Where-Object { $_.EventId -match '^(SEND|DELIVER)$' }) {
            $externalRcpts = foreach ($rcpt in @($log.Recipients)) {
                if (-not $rcpt -or $rcpt -notmatch '@') { continue }
                $dom = ($rcpt -split '@')[-1].ToLower()
                if ($InternalDomains -notcontains $dom) { $rcpt }
            }
            if (-not $externalRcpts) { continue }

            $rcptString = ($log.Recipients -join ' ')
            $provider   = switch -Regex ($rcptString) {
                'mail\.protection\.outlook\.com'        { 'Microsoft 365 / Exchange Online'; break }
                'aspmx\..*google|googlemail|gmail\.com' { 'Google Workspace'; break }
                'pphosted\.com|proofpoint'              { 'Proofpoint'; break }
                'mimecast'                              { 'Mimecast'; break }
                'barracuda'                             { 'Barracuda'; break }
                'cisco\.iphmx'                          { 'Cisco IronPort'; break }
                default                                 { 'External SMTP' }
            }
            [pscustomobject]@{
                Timestamp        = $log.Timestamp
                EventId          = $log.EventId
                Sender           = $log.Sender
                ExternalRcpts    = ($externalRcpts -join '; ')
                CloudProvider    = $provider
                ConnectorId      = (& $getEvtField $log 'ConnectorId')
                Subject          = $log.MessageSubject
                MessageId        = $log.MessageId
                Server           = $log.ServerHostname
            }
        }

        # --- Nested expansion chains ---
        # Group EXPANDs by MessageId. If a recipient of one EXPAND is itself
        # a RelatedRecipientAddress in another EXPAND (same MessageId), chain.
        $nestedChains = @()
        foreach ($grp in ($expandEvents | Group-Object MessageId | Where-Object { $_.Count -gt 1 })) {
            $ordered = $grp.Group | Sort-Object Timestamp
            $steps   = foreach ($e in $ordered) {
                $src = $e.RelatedRecipientAddress
                if (-not $src) { $src = & $getEvtField $e 'RelatedRecipientAddress' }
                $members = ($e.Recipients -join ', ')
                "{0}  ==>  [{1}]" -f $src, $members
            }
            $nestedChains += [pscustomobject]@{
                MessageId = $grp.Name
                Depth     = $grp.Count
                Subject   = $ordered[0].MessageSubject
                Sender    = $ordered[0].Sender
                Chain     = $steps -join '   |   '
            }
        }

        # --- Failures (auth vs. other) ---
        $authPattern = '5\.7\.0|5\.7\.1|5\.7\.8|AuthRequired|NotAuthenticated|Client was not authenticated|RelayDenied|Relaying denied|not permitted to relay'
        $failEvents  = $all | Where-Object { $_.EventId -match 'fail' }

        $authFailures = foreach ($f in $failEvents) {
            $status = ($f.RecipientStatus -join ' ')
            $reason = & $getEvtField $f 'Reason'
            $cat    = & $getEvtField $f 'FailureCategory'
            $combo  = "$status $reason $cat"
            if ($combo -match $authPattern) {
                [pscustomobject]@{
                    Timestamp       = $f.Timestamp
                    Sender          = $f.Sender
                    FailedRecipient = ($f.Recipients -join '; ')
                    Status          = $status
                    Reason          = $reason
                    Category        = $cat
                    Subject         = $f.MessageSubject
                    MessageId       = $f.MessageId
                    Server          = $f.ServerHostname
                }
            }
        }

        $deliveryFailures = foreach ($f in $failEvents) {
            $status = ($f.RecipientStatus -join ' ')
            $reason = & $getEvtField $f 'Reason'
            $cat    = & $getEvtField $f 'FailureCategory'
            $combo  = "$status $reason $cat"
            if ($combo -notmatch $authPattern) {
                [pscustomobject]@{
                    Timestamp       = $f.Timestamp
                    Sender          = $f.Sender
                    FailedRecipient = ($f.Recipients -join '; ')
                    Status          = $status
                    Reason          = $reason
                    Category        = $cat
                    EventData       = if ($f.EventData) { ($f.EventData | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; ' } else { '' }
                    Subject         = $f.MessageSubject
                    MessageId       = $f.MessageId
                    Server          = $f.ServerHostname
                }
            }
        }

        # --- Summary ---
        $summary = [ordered]@{
            TotalEvents        = @($all).Count
            UniqueMessageIds   = @($all | Where-Object MessageId | Select-Object -ExpandProperty MessageId -Unique).Count
            UniqueSenders      = @($all | Where-Object Sender    | Select-Object -ExpandProperty Sender    -Unique).Count
            InternalDomains    = ($InternalDomains -join ', ')
            TimeFrom           = if ($all) { ($all | Measure-Object Timestamp -Minimum).Minimum } else { $null }
            TimeTo             = if ($all) { ($all | Measure-Object Timestamp -Maximum).Maximum } else { $null }
            EventBreakdown     = ($all | Group-Object EventId | Sort-Object Count -Descending |
                                    ForEach-Object { "$($_.Name)=$($_.Count)" }) -join '; '
            Counts = [ordered]@{
                GroupExpansions    = @($groupExpansions).Count
                Redirects          = @($redirects).Count
                ExternalDeliveries = @($externals).Count
                NestedChains       = @($nestedChains).Count
                AuthFailures       = @($authFailures).Count
                DeliveryFailures   = @($deliveryFailures).Count
            }
        }

        [pscustomobject]@{
            Summary            = [pscustomobject]$summary
            GroupExpansions    = @($groupExpansions)
            Redirects          = @($redirects)
            ExternalDeliveries = @($externals)
            NestedChains       = @($nestedChains)
            AuthFailures       = @($authFailures)
            DeliveryFailures   = @($deliveryFailures)
        }
    }
}

function Format-MessageTrackingForensicsReport {
    <#
    .SYNOPSIS
        Baut einen lesbaren Markdown-Report aus dem Forensik-Objekt.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Forensics)

    $sb = New-Object System.Text.StringBuilder
    $nl = [Environment]::NewLine

    [void]$sb.AppendLine('# Exchange Message Tracking - Forensik-Report')
    [void]$sb.AppendLine(('Erzeugt: {0:yyyy-MM-dd HH:mm:ss}' -f (Get-Date)))
    [void]$sb.AppendLine()

    # --- Summary ---
    [void]$sb.AppendLine('## Zusammenfassung')
    $s = $Forensics.Summary
    [void]$sb.AppendLine(('- Events gesamt:           {0}' -f $s.TotalEvents))
    [void]$sb.AppendLine(('- Eindeutige MessageIds:   {0}' -f $s.UniqueMessageIds))
    [void]$sb.AppendLine(('- Eindeutige Sender:       {0}' -f $s.UniqueSenders))
    [void]$sb.AppendLine(('- Interne Domaenen:        {0}' -f $s.InternalDomains))
    [void]$sb.AppendLine(('- Zeitraum:                {0:yyyy-MM-dd HH:mm:ss} - {1:yyyy-MM-dd HH:mm:ss}' -f $s.TimeFrom, $s.TimeTo))
    [void]$sb.AppendLine(('- Event-Verteilung:        {0}' -f $s.EventBreakdown))
    [void]$sb.AppendLine()
    [void]$sb.AppendLine('### Kategorien')
    [void]$sb.AppendLine(('- Gruppen-Expansionen:     {0}' -f $s.Counts.GroupExpansions))
    [void]$sb.AppendLine(('- Redirects:               {0}' -f $s.Counts.Redirects))
    [void]$sb.AppendLine(('- Externe Zustellungen:    {0}' -f $s.Counts.ExternalDeliveries))
    [void]$sb.AppendLine(('- Nested Chains:           {0}' -f $s.Counts.NestedChains))
    [void]$sb.AppendLine(('- Auth-Fehler:             {0}' -f $s.Counts.AuthFailures))
    [void]$sb.AppendLine(('- Sonstige Fehler:         {0}' -f $s.Counts.DeliveryFailures))
    [void]$sb.AppendLine()

    # --- Group expansions (narrative) ---
    [void]$sb.AppendLine('## Gruppen-Expansionen')
    if ($Forensics.GroupExpansions.Count -eq 0) {
        [void]$sb.AppendLine('Keine EXPAND-Events im Datensatz.')
    } else {
        foreach ($g in ($Forensics.GroupExpansions | Group-Object GroupAddress)) {
            [void]$sb.AppendLine(('### {0}' -f $g.Name))
            [void]$sb.AppendLine(('- Aufgeloest: {0} Mal' -f $g.Count))
            $firstMembers = ($g.Group[0].Members -split '; ' | Select-Object -First 8) -join ', '
            if (($g.Group[0].Members -split '; ').Count -gt 8) { $firstMembers += ', ...' }
            [void]$sb.AppendLine(('- Mitglieder (Beispiel): {0}' -f $firstMembers))
            [void]$sb.AppendLine(('- Erster Zeitstempel:    {0:yyyy-MM-dd HH:mm:ss}' -f ($g.Group | Measure-Object Timestamp -Minimum).Minimum))
            [void]$sb.AppendLine()
        }
    }

    # --- Redirects ---
    [void]$sb.AppendLine('## Redirects / Weiterleitungen')
    if ($Forensics.Redirects.Count -eq 0) {
        [void]$sb.AppendLine('Keine REDIRECT-Events.')
    } else {
        foreach ($r in $Forensics.Redirects) {
            [void]$sb.AppendLine(('- {0:yyyy-MM-dd HH:mm:ss}  {1}  -->  {2}  (Subject: {3})' -f $r.Timestamp, $r.OriginalRcpt, $r.RedirectedTo, $r.Subject))
        }
    }
    [void]$sb.AppendLine()

    # --- External deliveries ---
    [void]$sb.AppendLine('## Externe Zustellungen')
    if ($Forensics.ExternalDeliveries.Count -eq 0) {
        [void]$sb.AppendLine('Keine externen Zustellungen.')
    } else {
        foreach ($p in ($Forensics.ExternalDeliveries | Group-Object CloudProvider | Sort-Object Count -Descending)) {
            [void]$sb.AppendLine(('### {0}  ({1} Events)' -f $p.Name, $p.Count))
            foreach ($e in ($p.Group | Select-Object -First 20)) {
                [void]$sb.AppendLine(('- {0:yyyy-MM-dd HH:mm:ss}  {1} -> {2}  [Subject: {3}]' -f $e.Timestamp, $e.Sender, $e.ExternalRcpts, $e.Subject))
            }
            if ($p.Count -gt 20) { [void]$sb.AppendLine(('- ... und weitere {0} (siehe CSV)' -f ($p.Count - 20))) }
            [void]$sb.AppendLine()
        }
    }

    # --- Nested chains ---
    [void]$sb.AppendLine('## Nested-Group-Ketten')
    if ($Forensics.NestedChains.Count -eq 0) {
        [void]$sb.AppendLine('Keine verschachtelten Expansionen erkannt.')
    } else {
        foreach ($n in $Forensics.NestedChains) {
            [void]$sb.AppendLine(('### MessageId: {0}' -f $n.MessageId))
            [void]$sb.AppendLine(('- Tiefe:   {0}' -f $n.Depth))
            [void]$sb.AppendLine(('- Sender:  {0}' -f $n.Sender))
            [void]$sb.AppendLine(('- Subject: {0}' -f $n.Subject))
            [void]$sb.AppendLine(('- Kette:   {0}' -f $n.Chain))
            [void]$sb.AppendLine()
        }
    }

    # --- Auth failures ---
    [void]$sb.AppendLine('## Auth-Fehler')
    if ($Forensics.AuthFailures.Count -eq 0) {
        [void]$sb.AppendLine('Keine Auth-bezogenen Fehler.')
    } else {
        foreach ($f in $Forensics.AuthFailures) {
            [void]$sb.AppendLine(('- {0:yyyy-MM-dd HH:mm:ss}  {1} -> {2}' -f $f.Timestamp, $f.Sender, $f.FailedRecipient))
            [void]$sb.AppendLine(('    Status:   {0}' -f $f.Status))
            if ($f.Reason)   { [void]$sb.AppendLine(('    Reason:   {0}' -f $f.Reason)) }
            if ($f.Category) { [void]$sb.AppendLine(('    Category: {0}' -f $f.Category)) }
        }
    }
    [void]$sb.AppendLine()

    # --- Other delivery failures ---
    [void]$sb.AppendLine('## Sonstige Zustellfehler')
    if ($Forensics.DeliveryFailures.Count -eq 0) {
        [void]$sb.AppendLine('Keine weiteren Zustellfehler.')
    } else {
        foreach ($f in ($Forensics.DeliveryFailures | Select-Object -First 50)) {
            [void]$sb.AppendLine(('- {0:yyyy-MM-dd HH:mm:ss}  {1} -> {2}' -f $f.Timestamp, $f.Sender, $f.FailedRecipient))
            [void]$sb.AppendLine(('    Status: {0}' -f $f.Status))
            if ($f.Reason)    { [void]$sb.AppendLine(('    Reason: {0}' -f $f.Reason)) }
            if ($f.EventData) { [void]$sb.AppendLine(('    EventData: {0}' -f $f.EventData)) }
        }
        if ($Forensics.DeliveryFailures.Count -gt 50) {
            [void]$sb.AppendLine(('- ... und weitere {0} (siehe CSV)' -f ($Forensics.DeliveryFailures.Count - 50)))
        }
    }

    return $sb.ToString()
}

function Export-MessageTrackingForensics {
    <#
    .SYNOPSIS
        Schreibt Markdown-Report plus CSV-Detailschnitte in einen
        Ordner mit Zeitstempel.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Forensics,
        [Parameter(Mandatory)] $ExpandedLogs,
        [Parameter(Mandatory)] [string] $ExportPath
    )

    $stamp     = Get-Date -Format 'yyyy.MM.dd_HH.mm.ss'
    $outDir    = Join-Path $ExportPath "${stamp}_tracking_forensics"
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null

    Format-MessageTrackingForensicsReport -Forensics $Forensics |
        Set-Content -Path (Join-Path $outDir '00_forensics_report.md') -Encoding UTF8

    $csvOpts = @{ Delimiter = ';'; NoTypeInformation = $true; Encoding = 'UTF8'; Force = $true }
    if ($Forensics.GroupExpansions.Count)    { $Forensics.GroupExpansions    | Export-Csv -Path (Join-Path $outDir '01_group_expansions.csv')    @csvOpts }
    if ($Forensics.Redirects.Count)          { $Forensics.Redirects          | Export-Csv -Path (Join-Path $outDir '02_redirects.csv')           @csvOpts }
    if ($Forensics.ExternalDeliveries.Count) { $Forensics.ExternalDeliveries | Export-Csv -Path (Join-Path $outDir '03_external_deliveries.csv') @csvOpts }
    if ($Forensics.NestedChains.Count)       { $Forensics.NestedChains       | Export-Csv -Path (Join-Path $outDir '04_nested_chains.csv')       @csvOpts }
    if ($Forensics.AuthFailures.Count)       { $Forensics.AuthFailures       | Export-Csv -Path (Join-Path $outDir '05_auth_failures.csv')       @csvOpts }
    if ($Forensics.DeliveryFailures.Count)   { $Forensics.DeliveryFailures   | Export-Csv -Path (Join-Path $outDir '06_delivery_failures.csv')   @csvOpts }

    $ExpandedLogs | ConvertTo-FlatMessageTrackingLog |
        Export-Csv -Path (Join-Path $outDir '99_expanded_tracking_flat.csv') @csvOpts

    return $outDir
}

# =============================================================================
# GUI
# =============================================================================

function Show-MessageTrackingGUI {
    [CmdletBinding()]
    param(
        [string] $DefaultExportPath = (Join-Path $env:USERPROFILE 'Desktop')
    )

    # -------------------------- Form --------------------------
    $form                 = New-Object System.Windows.Forms.Form
    $form.Text            = 'Exchange Message Tracking GUI  -  refaktoriert'
    $form.StartPosition   = 'CenterScreen'
    $form.ClientSize      = New-Object System.Drawing.Size(760, 820)
    $form.MinimumSize     = New-Object System.Drawing.Size(760, 820)
    $form.Font            = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.FormBorderStyle = 'Sizable'

    # -------------------------- Layout helpers --------------------------
    # rowY/rowHeight werden script-scoped gehalten, damit die geschachtelte
    # New-Row-Hilfsfunktion den Zaehler zurueckschreiben kann (PS-Funktionen
    # erzeugen sonst eine neue lokale Variable beim '+=').
    $script:rowHeight = 28
    $script:rowY      = 12

    $newRow = {
        param([string]$Label, [string]$DefaultValue = '', [bool]$Checked = $false, [int]$InputWidth = 520)
        $cb          = New-Object System.Windows.Forms.CheckBox
        $cb.Location = New-Object System.Drawing.Point(12, $script:rowY)
        $cb.Size     = New-Object System.Drawing.Size(18, 22)
        $cb.Checked  = $Checked

        $lbl          = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(34, ($script:rowY + 3))
        $lbl.Size     = New-Object System.Drawing.Size(120, 20)
        $lbl.Text     = $Label

        $tb          = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(160, $script:rowY)
        $tb.Size     = New-Object System.Drawing.Size($InputWidth, 22)
        $tb.Text     = $DefaultValue
        $tb.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

        $form.Controls.AddRange(@($cb, $lbl, $tb))
        $script:rowY += $script:rowHeight
        return @{ Check = $cb; Text = $tb }
    }

    # -------------------------- Filter rows --------------------------
    $rowSender    = & $newRow -Label 'Sender'
    $rowRecipient = & $newRow -Label 'Empfaenger'
    $rowEventId   = & $newRow -Label 'EventID'         # z.B. DELIVER, FAIL, RECEIVE
    $rowMsgId     = & $newRow -Label 'MessageID'
    $rowIntMsgId  = & $newRow -Label 'InternalMsgID'
    $rowSubject   = & $newRow -Label 'Subject'
    $rowReference = & $newRow -Label 'Reference'
    $rowServer    = & $newRow -Label 'Server(s)' -DefaultValue '' -Checked $false
    # Server-Hinweis
    $lblServerHint      = New-Object System.Windows.Forms.Label
    $lblServerHint.Text = '(leer = alle Transport-Services; sonst Komma-Liste: MAIL-01,MAIL-02)'
    $lblServerHint.Location = New-Object System.Drawing.Point(160, $script:rowY)
    $lblServerHint.Size     = New-Object System.Drawing.Size(560, 18)
    $lblServerHint.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($lblServerHint)
    $script:rowY += 22

    # -------------------------- Date pickers --------------------------
    $cbStart            = New-Object System.Windows.Forms.CheckBox
    $cbStart.Location   = New-Object System.Drawing.Point(12, $script:rowY)
    $cbStart.Size       = New-Object System.Drawing.Size(18, 22)
    $cbStart.Checked    = $true
    $lblStart           = New-Object System.Windows.Forms.Label
    $lblStart.Location  = New-Object System.Drawing.Point(34, ($script:rowY + 3))
    $lblStart.Size      = New-Object System.Drawing.Size(120, 20)
    $lblStart.Text      = 'Start'
    $dtStart            = New-Object System.Windows.Forms.DateTimePicker
    $dtStart.Location   = New-Object System.Drawing.Point(160, $script:rowY)
    $dtStart.Size       = New-Object System.Drawing.Size(240, 22)
    $dtStart.Format     = 'Custom'
    $dtStart.CustomFormat = 'dd.MM.yyyy HH:mm'
    $dtStart.Value      = (Get-Date).AddDays(-1)
    $form.Controls.AddRange(@($cbStart, $lblStart, $dtStart))
    $script:rowY += $rowHeight

    $cbEnd              = New-Object System.Windows.Forms.CheckBox
    $cbEnd.Location     = New-Object System.Drawing.Point(12, $script:rowY)
    $cbEnd.Size         = New-Object System.Drawing.Size(18, 22)
    $cbEnd.Checked      = $true
    $lblEnd             = New-Object System.Windows.Forms.Label
    $lblEnd.Location    = New-Object System.Drawing.Point(34, ($script:rowY + 3))
    $lblEnd.Size        = New-Object System.Drawing.Size(120, 20)
    $lblEnd.Text        = 'Ende'
    $dtEnd              = New-Object System.Windows.Forms.DateTimePicker
    $dtEnd.Location     = New-Object System.Drawing.Point(160, $script:rowY)
    $dtEnd.Size         = New-Object System.Drawing.Size(240, 22)
    $dtEnd.Format       = 'Custom'
    $dtEnd.CustomFormat = 'dd.MM.yyyy HH:mm'
    $dtEnd.Value        = (Get-Date)
    $form.Controls.AddRange(@($cbEnd, $lblEnd, $dtEnd))
    $script:rowY += $rowHeight

    # -------------------------- ResultSize --------------------------
    $cbRes             = New-Object System.Windows.Forms.CheckBox
    $cbRes.Location    = New-Object System.Drawing.Point(12, $script:rowY)
    $cbRes.Size        = New-Object System.Drawing.Size(18, 22)
    $cbRes.Checked     = $true
    $lblRes            = New-Object System.Windows.Forms.Label
    $lblRes.Location   = New-Object System.Drawing.Point(34, ($script:rowY + 3))
    $lblRes.Size       = New-Object System.Drawing.Size(120, 20)
    $lblRes.Text       = 'ResultSize'
    $tbRes             = New-Object System.Windows.Forms.TextBox
    $tbRes.Location    = New-Object System.Drawing.Point(160, $script:rowY)
    $tbRes.Size        = New-Object System.Drawing.Size(120, 22)
    $tbRes.Text        = 'unlimited'
    $form.Controls.AddRange(@($cbRes, $lblRes, $tbRes))
    $script:rowY += ($rowHeight + 6)

    # -------------------------- Options group --------------------------
    $grpOpts             = New-Object System.Windows.Forms.GroupBox
    $grpOpts.Text        = 'Optionen'
    $grpOpts.Location    = New-Object System.Drawing.Point(12, $script:rowY)
    $grpOpts.Size        = New-Object System.Drawing.Size(720, 190)
    $grpOpts.Anchor      = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

    $chkHealth          = New-Object System.Windows.Forms.CheckBox
    $chkHealth.Location = New-Object System.Drawing.Point(16, 22)
    $chkHealth.Size     = New-Object System.Drawing.Size(340, 22)
    $chkHealth.Text     = 'HealthMailbox-Nachrichten ausblenden'
    $chkHealth.Checked  = $true

    $chkFail            = New-Object System.Windows.Forms.CheckBox
    $chkFail.Location   = New-Object System.Drawing.Point(16, 46)
    $chkFail.Size       = New-Object System.Drawing.Size(340, 22)
    $chkFail.Text       = 'Nur Fehler (EventId -match ''fail'')'

    $chkGroup           = New-Object System.Windows.Forms.CheckBox
    $chkGroup.Location  = New-Object System.Drawing.Point(16, 70)
    $chkGroup.Size      = New-Object System.Drawing.Size(340, 22)
    $chkGroup.Text      = 'Gruppierung nach Sender zusaetzlich anzeigen'

    $chkCsv             = New-Object System.Windows.Forms.CheckBox
    $chkCsv.Location    = New-Object System.Drawing.Point(16, 94)
    $chkCsv.Size        = New-Object System.Drawing.Size(340, 22)
    $chkCsv.Text        = 'CSV-Export (flach, Semikolon, UTF8)'

    $chkDetailed           = New-Object System.Windows.Forms.CheckBox
    $chkDetailed.Location  = New-Object System.Drawing.Point(370, 22)
    $chkDetailed.Size      = New-Object System.Drawing.Size(340, 22)
    $chkDetailed.Text      = 'Detailansicht im Gridview (alle Spalten)'

    $chkUser            = New-Object System.Windows.Forms.CheckBox
    $chkUser.Location   = New-Object System.Drawing.Point(370, 46)
    $chkUser.Size       = New-Object System.Drawing.Size(340, 22)
    $chkUser.Text       = 'Zusaetzlicher Empfangsbericht (DELIVER/STOREDRIVER)'

    # --- Auto-Expand row ---
    $chkAutoExpand           = New-Object System.Windows.Forms.CheckBox
    $chkAutoExpand.Location  = New-Object System.Drawing.Point(16, 126)
    $chkAutoExpand.Size      = New-Object System.Drawing.Size(330, 22)
    $chkAutoExpand.Text      = 'Auto-Expand via MessageId, wenn Treffer <='
    $chkAutoExpand.Checked   = $true

    $numExpandLimit          = New-Object System.Windows.Forms.NumericUpDown
    $numExpandLimit.Location = New-Object System.Drawing.Point(346, 124)
    $numExpandLimit.Size     = New-Object System.Drawing.Size(60, 22)
    $numExpandLimit.Minimum  = 1
    $numExpandLimit.Maximum  = 5000
    $numExpandLimit.Value    = 50

    $lblExpandHint           = New-Object System.Windows.Forms.Label
    $lblExpandHint.Location  = New-Object System.Drawing.Point(412, 128)
    $lblExpandHint.Size      = New-Object System.Drawing.Size(298, 18)
    $lblExpandHint.Text      = '(je MessageId wird erneut ueber alle Server gesucht)'
    $lblExpandHint.ForeColor = [System.Drawing.Color]::DimGray

    $chkForensics            = New-Object System.Windows.Forms.CheckBox
    $chkForensics.Location   = New-Object System.Drawing.Point(16, 152)
    $chkForensics.Size       = New-Object System.Drawing.Size(690, 22)
    $chkForensics.Text       = 'Forensik-Export (Expansion, Redirects, Cloud/Auth-Fehler) - setzt Auto-Expand voraus'
    $chkForensics.Checked    = $false

    $grpOpts.Controls.AddRange(@(
        $chkHealth, $chkFail, $chkGroup, $chkCsv, $chkDetailed, $chkUser,
        $chkAutoExpand, $numExpandLimit, $lblExpandHint, $chkForensics
    ))
    $form.Controls.Add($grpOpts)
    $script:rowY += 198

    # -------------------------- Export path --------------------------
    $lblPath          = New-Object System.Windows.Forms.Label
    $lblPath.Location = New-Object System.Drawing.Point(12, ($script:rowY + 3))
    $lblPath.Size     = New-Object System.Drawing.Size(140, 20)
    $lblPath.Text     = 'Export-Ordner:'
    $tbPath           = New-Object System.Windows.Forms.TextBox
    $tbPath.Location  = New-Object System.Drawing.Point(160, $script:rowY)
    $tbPath.Size      = New-Object System.Drawing.Size(440, 22)
    $tbPath.Text      = $DefaultExportPath
    $tbPath.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnBrowse        = New-Object System.Windows.Forms.Button
    $btnBrowse.Location = New-Object System.Drawing.Point(610, ($script:rowY - 1))
    $btnBrowse.Size   = New-Object System.Drawing.Size(120, 24)
    $btnBrowse.Text   = 'Durchsuchen...'
    $btnBrowse.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Controls.AddRange(@($lblPath, $tbPath, $btnBrowse))
    $script:rowY += ($rowHeight + 4)

    # -------------------------- Command preview --------------------------
    $lblCmd          = New-Object System.Windows.Forms.Label
    $lblCmd.Location = New-Object System.Drawing.Point(12, $script:rowY)
    $lblCmd.Size     = New-Object System.Drawing.Size(300, 20)
    $lblCmd.Text     = 'Kommando-Vorschau:'
    $form.Controls.Add($lblCmd)
    $script:rowY += 20

    $tbCmd           = New-Object System.Windows.Forms.TextBox
    $tbCmd.Location  = New-Object System.Drawing.Point(12, $script:rowY)
    $tbCmd.Size      = New-Object System.Drawing.Size(720, 60)
    $tbCmd.Multiline = $true
    $tbCmd.ReadOnly  = $true
    $tbCmd.ScrollBars = 'Vertical'
    $tbCmd.Font      = New-Object System.Drawing.Font('Consolas', 8.5)
    $tbCmd.BackColor = [System.Drawing.Color]::WhiteSmoke
    $tbCmd.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Controls.Add($tbCmd)
    $script:rowY += 68

    # -------------------------- Action buttons --------------------------
    $btnSearch          = New-Object System.Windows.Forms.Button
    $btnSearch.Location = New-Object System.Drawing.Point(12, $script:rowY)
    $btnSearch.Size     = New-Object System.Drawing.Size(120, 32)
    $btnSearch.Text     = 'Suchen'
    $btnSearch.BackColor = [System.Drawing.Color]::FromArgb(210, 230, 255)

    $btnExpand          = New-Object System.Windows.Forms.Button
    $btnExpand.Location = New-Object System.Drawing.Point(138, $script:rowY)
    $btnExpand.Size     = New-Object System.Drawing.Size(160, 32)
    $btnExpand.Text     = 'MessageIds expandieren'
    $btnExpand.Enabled  = $false

    $btnForensics          = New-Object System.Windows.Forms.Button
    $btnForensics.Location = New-Object System.Drawing.Point(304, $script:rowY)
    $btnForensics.Size     = New-Object System.Drawing.Size(160, 32)
    $btnForensics.Text     = 'Forensik-Export'
    $btnForensics.Enabled  = $false

    $btnClose           = New-Object System.Windows.Forms.Button
    $btnClose.Location  = New-Object System.Drawing.Point(612, $script:rowY)
    $btnClose.Size      = New-Object System.Drawing.Size(120, 32)
    $btnClose.Text      = 'Schliessen'
    $btnClose.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

    $form.Controls.AddRange(@($btnSearch, $btnExpand, $btnForensics, $btnClose))
    $script:rowY += 36

    # State-Speicher fuer Expand / Forensik-Export
    $script:lastRawResults      = $null
    $script:lastExpandedResults = $null
    $script:lastSearchStart     = $null
    $script:lastSearchEnd       = $null
    $script:lastServerList      = @()

    # -------------------------- Status bar --------------------------
    $status             = New-Object System.Windows.Forms.StatusStrip
    $statusLbl          = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLbl.Text     = 'Bereit.'
    $status.Items.Add($statusLbl) | Out-Null
    $form.Controls.Add($status)

    # ==========================================================================
    # Logic
    # ==========================================================================

    # --- Browse folder
    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.SelectedPath = $tbPath.Text
        if ($dlg.ShowDialog() -eq 'OK') { $tbPath.Text = $dlg.SelectedPath }
    })

    # --- Build parameter set from current form state
    $splitMulti = {
        param([string]$Text)
        # Splittet an Komma / Semikolon, trimmt Whitespace, verwirft Leereintraege.
        if (-not $Text) { return @() }
        , @($Text -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    $getParameters = {
        $p = @{}
        if ($rowSender.Check.Checked -and $rowSender.Text.Text) {
            $vals = & $splitMulti $rowSender.Text.Text
            if ($vals.Count -eq 1) { $p.Sender = $vals[0] }
            elseif ($vals.Count -gt 1) { $p.Sender = $vals }
        }
        if ($rowRecipient.Check.Checked -and $rowRecipient.Text.Text) {
            $vals = & $splitMulti $rowRecipient.Text.Text
            if ($vals.Count -eq 1) { $p.Recipients = $vals[0] }
            elseif ($vals.Count -gt 1) { $p.Recipients = $vals }
        }
        if ($rowEventId.Check.Checked   -and $rowEventId.Text.Text)   { $p.EventId           = $rowEventId.Text.Text.Trim() }
        if ($rowMsgId.Check.Checked     -and $rowMsgId.Text.Text)     { $p.MessageId         = $rowMsgId.Text.Text.Trim() }
        if ($rowIntMsgId.Check.Checked  -and $rowIntMsgId.Text.Text)  { $p.InternalMessageId = $rowIntMsgId.Text.Text.Trim() }
        if ($rowSubject.Check.Checked   -and $rowSubject.Text.Text)   { $p.MessageSubject    = $rowSubject.Text.Text.Trim() }
        if ($rowReference.Check.Checked -and $rowReference.Text.Text) { $p.Reference         = $rowReference.Text.Text.Trim() }
        if ($cbStart.Checked) { $p.Start = [datetime]$dtStart.Value }
        if ($cbEnd.Checked)   { $p.End   = [datetime]$dtEnd.Value   }
        if ($cbRes.Checked -and $tbRes.Text) { $p.ResultSize = $tbRes.Text.Trim() }
        return $p
    }

    $getServerList = {
        if ($rowServer.Check.Checked -and $rowServer.Text.Text.Trim()) {
            return ($rowServer.Text.Text -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
        return @()
    }

    # --- Command preview (updated on any control change)
    $updatePreview = {
        $p = & $getParameters
        $srv = & $getServerList
        $parts = @()
        foreach ($k in $p.Keys) {
            $v = $p[$k]
            if     ($v -is [datetime]) { $v = "(Get-Date '{0:dd.MM.yyyy HH:mm}')" -f $v }
            elseif ($v -is [array])    { $v = "@('" + (($v | ForEach-Object { $_ }) -join "','") + "')" }
            elseif ($v -is [string])   { $v = "'$v'" }
            $parts += "-$k $v"
        }
        $paramLine = $parts -join ' '
        $srvPart   = if ($srv.Count) { "@('$($srv -join ''',''')') | %{ Get-MessageTrackingLog -Server `$_ $paramLine }" }
                     else            { "Get-TransportService | %{ Get-MessageTrackingLog -Server `$_.Name $paramLine }" }
        $tail = @()
        if ($chkHealth.Checked) { $tail += "Where-Object { `$_.Sender -notmatch 'HealthMailbox' }" }
        if ($chkFail.Checked)   { $tail += "Where-Object { `$_.EventId -match 'fail' }" }
        $full = $srvPart
        if ($tail.Count) { $full += ' | ' + ($tail -join ' | ') }
        if ($chkCsv.Checked) {
            $full += " | ConvertTo-FlatMessageTrackingLog | Export-Csv -Path <Path> -Delimiter ';' -NoTypeInformation -Encoding UTF8"
        }
        $tbCmd.Text = $full
    }

    # Wire preview updates
    foreach ($ctrl in @(
        $rowSender.Check, $rowSender.Text, $rowRecipient.Check, $rowRecipient.Text,
        $rowEventId.Check, $rowEventId.Text, $rowMsgId.Check, $rowMsgId.Text,
        $rowIntMsgId.Check, $rowIntMsgId.Text, $rowSubject.Check, $rowSubject.Text,
        $rowReference.Check, $rowReference.Text, $rowServer.Check, $rowServer.Text,
        $cbStart, $cbEnd, $cbRes, $tbRes,
        $chkHealth, $chkFail, $chkCsv
    )) {
        if ($ctrl -is [System.Windows.Forms.CheckBox]) { $ctrl.Add_CheckedChanged($updatePreview) }
        else                                           { $ctrl.Add_TextChanged($updatePreview) }
    }
    $dtStart.Add_ValueChanged($updatePreview)
    $dtEnd.Add_ValueChanged($updatePreview)
    & $updatePreview

    # --- Helper: Expand logic (shared between auto-expand and manual button)
    $doExpand = {
        param($rawResults, [int]$limit, $servers, $start, $end)
        $statusLbl.Text = 'Expansion laeuft...'
        $form.Refresh()

        $progress = {
            param($i, $total, $mid)
            $statusLbl.Text = ('Expansion {0}/{1}: {2}' -f $i, $total, $mid)
            $form.Refresh()
        }

        $expanded = Expand-MessageTrackingByMessageId `
            -Logs       $rawResults `
            -Limit      $limit `
            -ServerList $servers `
            -Start      $start `
            -End        $end `
            -Progress   $progress

        return , @($expanded)
    }

    $updateButtonState = {
        $btnExpand.Enabled    = ($null -ne $script:lastRawResults)      -and (@($script:lastRawResults).Count -gt 0)
        $btnForensics.Enabled = ($null -ne $script:lastExpandedResults) -and (@($script:lastExpandedResults).Count -gt 0)
    }

    # --- Search
    $btnSearch.Add_Click({
        $btnSearch.Enabled    = $false
        $btnExpand.Enabled    = $false
        $btnForensics.Enabled = $false
        $statusLbl.Text       = 'Suche laeuft...'
        $form.Cursor          = 'WaitCursor'
        $form.Refresh()

        try {
            $params      = & $getParameters
            $servers     = & $getServerList
            $searchStart = Get-Date

            $results = Invoke-MessageTrackingSearch `
                -ServerList    $servers `
                -Parameters    $params `
                -ExcludeHealth:$chkHealth.Checked `
                -FailuresOnly:$chkFail.Checked

            $count    = @($results).Count
            $duration = (Get-Date) - $searchStart
            $statusLbl.Text = ('{0} Eintraege in {1:n1}s.' -f $count, $duration.TotalSeconds)

            # State aktualisieren (auch bei 0 Treffern - leere Arrays sind valide state)
            $script:lastRawResults      = @($results)
            $script:lastExpandedResults = $null
            $script:lastServerList      = $servers
            $script:lastSearchStart     = if ($params.ContainsKey('Start')) { $params.Start } else { $null }
            $script:lastSearchEnd       = if ($params.ContainsKey('End'))   { $params.End   } else { $null }

            if ($count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    'Keine passenden Eintraege gefunden.',
                    'Message Tracking',
                    'OK', 'Information') | Out-Null
                return
            }

            # --- Main gridview (immer flach, damit Recipients/EventData lesbar)
            $flat = $results | ConvertTo-FlatMessageTrackingLog
            if ($chkDetailed.Checked) {
                $flat | Out-GridView -Title ("Message Tracking - {0} Eintraege" -f $count)
            } else {
                $flat | Select-Object Timestamp, EventId, Source, Sender, Recipients, MessageSubject |
                    Out-GridView -Title ("Message Tracking - {0} Eintraege" -f $count)
            }

            if ($chkUser.Checked) {
                $flat | Where-Object { $_.EventId -match 'DELIVER' -and $_.Source -match 'STOREDRIVER' } |
                    Select-Object Timestamp, EventId, Sender, Recipients, MessageSubject |
                    Out-GridView -Title 'Empfangsbericht (DELIVER/STOREDRIVER)'
            }

            if ($chkGroup.Checked) {
                $results | Group-Object Sender | Sort-Object Count -Descending |
                    Select-Object Count, Name |
                    Out-GridView -Title 'Gruppierung nach Sender'
            }

            # --- CSV-Export (Rohsuche, flach)
            if ($chkCsv.Checked) {
                $exportDir = $tbPath.Text.Trim()
                if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir -Force | Out-Null }

                $stamp = Get-Date -Format 'yyyy.MM.dd_HH.mm.ss'
                $hint  = @(
                    if ($params.Sender)      { if ($params.Sender -is [array])     { 'sender-multi' } else { "sender-$($params.Sender)" } }
                    if ($params.Recipients)  { if ($params.Recipients -is [array]) { 'rcpt-multi' }   else { "rcpt-$($params.Recipients)" } }
                    if ($params.MessageId)   { 'msgid' }
                    if ($chkFail.Checked)    { 'fail' }
                ) -join '_' -replace '[\\/:*?"<>|@]', '_'
                if (-not $hint) { $hint = 'tracking' }
                $file = Join-Path $exportDir "${stamp}_${hint}.csv"

                $flat | Export-Csv -Path $file -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Force
                $statusLbl.Text = '{0} Eintraege exportiert -> {1}' -f $count, $file
            }

            # --- Auto-Expand wenn unter Schwelle
            $uniqueMids = @($results | Where-Object MessageId | Select-Object -ExpandProperty MessageId -Unique).Count
            $limit      = [int]$numExpandLimit.Value
            $shouldExpand = $chkAutoExpand.Checked -and $uniqueMids -gt 0 -and $uniqueMids -le $limit

            if ($shouldExpand) {
                $expanded = & $doExpand $results $limit $servers $script:lastSearchStart $script:lastSearchEnd
                $script:lastExpandedResults = @($expanded)
                $statusLbl.Text = '{0} MessageIds expandiert zu {1} Events.' -f $uniqueMids, @($expanded).Count

                # Flach-Gridview der Expansion anzeigen
                $expanded | ConvertTo-FlatMessageTrackingLog |
                    Select-Object Timestamp, EventId, Source, Server*, Sender, Recipients, MessageSubject, MessageId |
                    Out-GridView -Title ('Expansion: {0} MessageIds / {1} Events' -f $uniqueMids, @($expanded).Count)

                # Automatischer Forensik-Export wenn gewuenscht
                if ($chkForensics.Checked) {
                    $forensics = $expanded | ConvertTo-MessageTrackingForensics
                    $exportDir = $tbPath.Text.Trim()
                    if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir -Force | Out-Null }
                    $outDir = Export-MessageTrackingForensics -Forensics $forensics -ExpandedLogs $expanded -ExportPath $exportDir
                    $statusLbl.Text = 'Forensik-Report -> {0}' -f $outDir

                    # Summary-Gridview
                    [pscustomobject]@{
                        Kategorie = 'Gruppen-Expansionen';    Anzahl = $forensics.Summary.Counts.GroupExpansions
                    },
                    [pscustomobject]@{ Kategorie = 'Redirects';             Anzahl = $forensics.Summary.Counts.Redirects },
                    [pscustomobject]@{ Kategorie = 'Externe Zustellungen'; Anzahl = $forensics.Summary.Counts.ExternalDeliveries },
                    [pscustomobject]@{ Kategorie = 'Nested Chains';        Anzahl = $forensics.Summary.Counts.NestedChains },
                    [pscustomobject]@{ Kategorie = 'Auth-Fehler';          Anzahl = $forensics.Summary.Counts.AuthFailures },
                    [pscustomobject]@{ Kategorie = 'Sonstige Fehler';      Anzahl = $forensics.Summary.Counts.DeliveryFailures } |
                        Out-GridView -Title ('Forensik-Zusammenfassung - {0}' -f $outDir)
                }
            }
            elseif ($chkAutoExpand.Checked -and $uniqueMids -gt $limit) {
                $statusLbl.Text = '{0} MessageIds - ueber Schwelle {1}. Auto-Expand uebersprungen.' -f $uniqueMids, $limit
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Suche schlaegt fehl: $($_.Exception.Message)",
                'Fehler',
                'OK', 'Error') | Out-Null
            $statusLbl.Text = 'Fehler.'
        }
        finally {
            $btnSearch.Enabled = $true
            & $updateButtonState
            $form.Cursor       = 'Default'
        }
    })

    # --- Manual expand
    $btnExpand.Add_Click({
        if (-not $script:lastRawResults -or @($script:lastRawResults).Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Keine Ergebnisse zum Expandieren.', 'Expand', 'OK', 'Information') | Out-Null
            return
        }
        $btnSearch.Enabled    = $false
        $btnExpand.Enabled    = $false
        $btnForensics.Enabled = $false
        $form.Cursor          = 'WaitCursor'
        try {
            $limit    = [int]$numExpandLimit.Value
            $expanded = & $doExpand $script:lastRawResults $limit $script:lastServerList $script:lastSearchStart $script:lastSearchEnd
            $script:lastExpandedResults = @($expanded)

            $mids = @($script:lastRawResults | Where-Object MessageId | Select-Object -ExpandProperty MessageId -Unique).Count
            $statusLbl.Text = '{0} MessageIds expandiert zu {1} Events.' -f $mids, @($expanded).Count

            $expanded | ConvertTo-FlatMessageTrackingLog |
                Select-Object Timestamp, EventId, Source, Server*, Sender, Recipients, MessageSubject, MessageId |
                Out-GridView -Title ('Expansion: {0} MessageIds / {1} Events' -f $mids, @($expanded).Count)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Expand schlaegt fehl: $($_.Exception.Message)",
                'Fehler', 'OK', 'Error') | Out-Null
        }
        finally {
            $btnSearch.Enabled = $true
            & $updateButtonState
            $form.Cursor       = 'Default'
        }
    })

    # --- Forensik-Export
    $btnForensics.Add_Click({
        if (-not $script:lastExpandedResults -or @($script:lastExpandedResults).Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                'Keine expandierten Ergebnisse. Zuerst "MessageIds expandieren" ausfuehren oder Auto-Expand aktivieren.',
                'Forensik', 'OK', 'Information') | Out-Null
            return
        }
        $btnForensics.Enabled = $false
        $form.Cursor          = 'WaitCursor'
        try {
            $exportDir = $tbPath.Text.Trim()
            if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir -Force | Out-Null }

            $forensics = $script:lastExpandedResults | ConvertTo-MessageTrackingForensics
            $outDir    = Export-MessageTrackingForensics -Forensics $forensics -ExpandedLogs $script:lastExpandedResults -ExportPath $exportDir
            $statusLbl.Text = 'Forensik-Report -> {0}' -f $outDir

            [pscustomobject]@{ Kategorie = 'Gruppen-Expansionen';   Anzahl = $forensics.Summary.Counts.GroupExpansions },
            [pscustomobject]@{ Kategorie = 'Redirects';             Anzahl = $forensics.Summary.Counts.Redirects },
            [pscustomobject]@{ Kategorie = 'Externe Zustellungen'; Anzahl = $forensics.Summary.Counts.ExternalDeliveries },
            [pscustomobject]@{ Kategorie = 'Nested Chains';         Anzahl = $forensics.Summary.Counts.NestedChains },
            [pscustomobject]@{ Kategorie = 'Auth-Fehler';           Anzahl = $forensics.Summary.Counts.AuthFailures },
            [pscustomobject]@{ Kategorie = 'Sonstige Fehler';       Anzahl = $forensics.Summary.Counts.DeliveryFailures } |
                Out-GridView -Title ('Forensik-Zusammenfassung - {0}' -f $outDir)

            # Auch den MD-Report oeffnen
            try { Invoke-Item (Join-Path $outDir '00_forensics_report.md') } catch { }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Forensik-Export schlaegt fehl: $($_.Exception.Message)",
                'Fehler', 'OK', 'Error') | Out-Null
        }
        finally {
            & $updateButtonState
            $form.Cursor = 'Default'
        }
    })

    $btnClose.Add_Click({ $form.Close() })

    # -------------------------- Run --------------------------
    [void]$form.ShowDialog()
    $form.Dispose()
}

# =============================================================================
# Entry point
# =============================================================================
# Wenn dot-sourced, stehen die Helfer zur Verfuegung. Direkter Aufruf startet
# die GUI:
if ($MyInvocation.InvocationName -ne '.') {
    Show-MessageTrackingGUI
}
