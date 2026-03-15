<#
.SYNOPSIS
    Converts a .PST file to a standards-compliant .MBOX file, preserving
    email content (HTML + plain text) and all file attachments as base64-
    encoded MIME parts.

.DESCRIPTION
    Uses the Outlook COM interop to mount a PST, walk every folder, and
    serialise each IPM.Note item as a proper multipart MIME message inside
    a single MBOX file. Designed for ingestion into email archival systems.

    Supports filtering by date range, folder patterns, and senders.
    Can split output into multiple files for large exports.

.PARAMETER PstPath
    Full or relative path to the input .PST file (mandatory).

.PARAMETER MboxPath
    Full or relative path to the output .MBOX file (optional).
    If omitted, the output file is created alongside the PST with the
    same base name and a .mbox extension.

.PARAMETER MaxEmails
    Optional maximum number of emails to export. If not specified,
    all emails are exported. Useful for testing with large PST files.

.PARAMETER FailedLogPath
    Optional path to write a log of failed items. If not specified,
    failed items are only displayed in the console output.

.PARAMETER DateFrom
    Optional start date for filtering emails. Emails received before
    this date are excluded.

.PARAMETER DateTo
    Optional end date for filtering emails. Emails received after
    this date are excluded.

.PARAMETER ExcludeFolders
    Optional array of folder name patterns to exclude. Folders matching
    any of these patterns (case-insensitive, wildcard supported) are skipped.

.PARAMETER ExcludeSenders
    Optional array of sender patterns to exclude. Emails from senders
    matching any of these patterns (case-insensitive, wildcard supported)
    are skipped.

.PARAMETER SplitSizeMB
    Optional maximum size in MB for each output file. When reached,
    creates a new numbered file (e.g., archive.mbox.001, archive.mbox.002).

.PARAMETER Verbose
    Enable verbose output showing detailed processing information.

.PARAMETER WhatIf
    Dry-run mode: scans the PST and reports what would be exported
    without actually writing any output files.

.PARAMETER DeDuplicate
    Remove duplicate emails based on Message-ID. This prevents
    the same email from appearing multiple times in the output.

.PARAMETER PreserveHeaders
    Array of additional header names to preserve in the output.
    By default, standard RFC headers are preserved.

.PARAMETER RestrictAttachments
    Enable security restrictions on attachments. When specified, blocks
    dangerous file extensions (exe, bat, sys, cmd, ps1, vbs, msg, etc.)
    and scans for macros, embedded OLE objects, file type mismatches,
    and archive bombs. By default, ALL attachments are exported.

.PARAMETER MaxAttachmentSizeMB
    Maximum size in MB for individual attachments. Larger attachments
    are skipped to prevent denial-of-service attacks (default: 50MB).

.PARAMETER AllowedBasePath
    Restrict output file creation to this directory. Prevents path
    traversal attacks by ensuring all output paths are within this base.

.PARAMETER BatchPath
    Path to a directory containing multiple .PST files. When specified,
    the script processes all .pst files in the directory sequentially,
    creating a corresponding .mbox file for each.

.EXAMPLE
    .\pst2mbox.ps1 "C:\Archive\mailbox.pst"
    Creates C:\Archive\mailbox.mbox

.EXAMPLE
    .\pst2mbox.ps1 "C:\Archive\mailbox.pst" "D:\Export\archive.mbox"
    Specifies custom output path

.EXAMPLE
    .\pst2mbox.ps1 .\backup.pst .\backup.mbox
    Uses current directory for both input and output

.EXAMPLE
    .\pst2mbox.ps1 "C:\Archive\mailbox.pst" -MaxEmails 1000
    Exports only the first 1000 emails

.EXAMPLE
    .\pst2mbox.ps1 "C:\Archive\mailbox.pst" -DateFrom "2024-01-01" -DateTo "2024-12-31"
    Exports only emails from year 2024

.EXAMPLE
    .\pst2mbox.ps1 "C:\Archive\mailbox.pst" -ExcludeFolders "Deleted Items","Junk"
    Excludes specific folders from export

.EXAMPLE
    .\pst2mbox.ps1 "C:\Archive\mailbox.pst" -ExcludeSenders "newsletter@","noreply@"
    Excludes emails from specific senders

.EXAMPLE
    .\pst2mbox.ps1 "C:\Archive\mailbox.pst" -SplitSizeMB 100
    Splits output into 100MB chunks

.EXAMPLE
    .\pst2mbox.ps1 "C:\Archive\mailbox.pst" -WhatIf
    Shows what would be exported without creating files

.NOTES
    Requirements:
      - Microsoft Outlook (desktop) installed
      - PowerShell 5.1+
      - Run as the user whose Outlook profile is active

    Author: PST2MBOX Contributors
    License: MIT
    Version: 2.1.0
#>

# -----------------------------------------------------------------------------
# PARAMETERS
# -----------------------------------------------------------------------------
param(
    [Parameter(Mandatory = $true, Position = 0,
               HelpMessage = "Path to the input .PST file.")]
    [string]$PstPath,

    [Parameter(Mandatory = $false, Position = 1,
               HelpMessage = "Path to the output .MBOX file (optional).")]
    [string]$MboxPath,

    [Parameter(Mandatory = $false)]
    [int]$MaxEmails = 0,

    [Parameter(Mandatory = $false)]
    [string]$FailedLogPath,

    [Parameter(Mandatory = $false)]
    [datetime]$DateFrom,

    [Parameter(Mandatory = $false)]
    [datetime]$DateTo,

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludeFolders,

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludeSenders,

    [Parameter(Mandatory = $false)]
    [int]$SplitSizeMB = 0,

    [Parameter(Mandatory = $false)]
    [switch]$WhatIf,

    [Parameter(Mandatory = $false)]
    [switch]$DeDuplicate,

    [Parameter(Mandatory = $false)]
    [string[]]$PreserveHeaders,

    [Parameter(Mandatory = $false)]
    [switch]$RestrictAttachments,

    [Parameter(Mandatory = $false)]
    [int]$MaxAttachmentSizeMB = 50,

    [Parameter(Mandatory = $false)]
    [string]$AllowedBasePath,

    [Parameter(Mandatory = $false)]
    [string]$BatchPath
)

# -----------------------------------------------------------------------------
# SECURITY: PATH VALIDATION & NORMALIZATION
# -----------------------------------------------------------------------------

function Test-SafePath {
    param(
        [string]$Path,
        [string]$BasePath,
        [string]$Type = "file"
    )
    # Resolve to absolute path and normalize
    $resolved = [System.IO.Path]::GetFullPath($Path)

    # Ensure path is within allowed base directory (prevent path traversal)
    if ($BasePath) {
        $baseResolved = [System.IO.Path]::GetFullPath($BasePath)
        if (-not $resolved.StartsWith($baseResolved, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Security: $Type path '$$Path' resolves outside allowed base directory"
        }
    }

    # Check for suspicious path characters
    if ($resolved -match '[;|&]') {
        throw "Security: Invalid characters detected in path"
    }

    return $resolved
}

function Get-SafeFileName { param([string]$name)
    $name = ($name -replace '[\\/]', '_') -replace "`0", ''
    $name = $name -replace '^\.*', ''
    if ($name.Length -gt 255) {
        $ext = [System.IO.Path]::GetExtension($name)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        if ($ext.Length -lt 20) { $name = $base.Substring(0, 255 - $ext.Length) + $ext }
        else { $name = $base.Substring(0, 250) }
    }
    return (Sanitise-Filename $name).Trim()
}

function Test-SafeFileExtension {
    param([string]$FileName)
    # Allowlist of safe attachment extensions
    $safeExtensions = @(
        '.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx',
        '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif', '.svg',
        '.zip', '.7z', '.rar', '.gz', '.tar',
        '.txt', '.csv', '.html', '.htm', '.xml', '.json', '.eml', '.ics', '.vcf',
        '.mp3', '.mp4', '.wav', '.avi', '.mkv', '.mov',
        '.rtf', '.odt', '.ods', '.odp'
    )
    $ext = [System.IO.Path]::GetExtension($FileName).ToLower()
    if ([string]::IsNullOrWhiteSpace($ext)) { return $true }  # No extension = allow
    return $safeExtensions -contains $ext
}

# -----------------------------------------------------------------------------
# RESOLVE PATHS
# -----------------------------------------------------------------------------

# Batch mode: process all PST files in a directory
if ($BatchPath) {
    if (-not (Test-Path $BatchPath)) { throw "BatchPath directory not found: $BatchPath" }
    $BatchPath = (Resolve-Path $BatchPath).Path
    $pstFiles = Get-ChildItem -Path $BatchPath -Filter "*.pst" -File
    if ($pstFiles.Count -eq 0) { Write-Host "No .pst files found in $BatchPath"; exit 0 }
    Write-Host "=== Batch Mode ===" -ForegroundColor Cyan
    Write-Host "Directory: $BatchPath"
    Write-Host "PST files found: $($pstFiles.Count)"
    Write-Host ""
    foreach ($pstFile in $pstFiles) {
        Write-Host "Processing: $($pstFile.Name)" -ForegroundColor Yellow
        & $PSCommandPath -PstPath $pstFile.FullName -MaxEmails $MaxEmails -DateFrom $DateFrom -DateTo $DateTo `
            -ExcludeFolders $ExcludeFolders -ExcludeSenders $ExcludeSenders -SplitSizeMB $SplitSizeMB `
            -DeDuplicate:$DeDuplicate -RestrictAttachments:$RestrictAttachments -MaxAttachmentSizeMB $MaxAttachmentSizeMB `
            -WhatIf:$WhatIf -Verbose:$Verbose
        Write-Host ""
    }
    Write-Host "=== Batch Complete ===" -ForegroundColor Cyan
    exit 0
}

$PstPath = (Resolve-Path $PstPath -ErrorAction Stop).Path

if ([string]::IsNullOrWhiteSpace($MboxPath)) {
    $MboxPath = [System.IO.Path]::ChangeExtension($PstPath, ".mbox")
} else {
    $MboxPath = Test-SafePath -Path $MboxPath -BasePath $PWD.Path -Type "output"
}

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"
$script:cleanupNeeded = $false

# =============================================================================
# CONSTANTS
# =============================================================================
$MAPI_TAGS = @{
    MessageId      = 'http://schemas.microsoft.com/mapi/proptag/0x1035001E'
    InReplyTo      = 'http://schemas.microsoft.com/mapi/proptag/0x1042001E'
    References     = 'http://schemas.microsoft.com/mapi/proptag/0x1039001E'
    AttachMethod   = 'http://schemas.microsoft.com/mapi/proptag/0x37050003'
    AttachMethodByReference = 6
    ListUnsubscribe = 'http://schemas.microsoft.com/mapi/proptag/0x007F001E'
}

$PST2MBOX_VERSION = '2.1.0'

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

function Get-Base64 ([byte[]]$bytes) {
    if (-not $bytes -or $bytes.Length -eq 0) { return "" }
    $b64 = [System.Convert]::ToBase64String($bytes, [System.Base64FormattingOptions]::InsertLineBreaks)
    return $b64 -replace "`r", ""
}

function Escape-FromLines ([string]$text) {
    return ($text -split "`r?`n" | ForEach-Object {
        if ($_ -match '^From ') { ">$_" } else { $_ }
    }) -join "`n"
}

function Fold-Header ([string]$name, [string]$value) {
    $full = "${name}: ${value}"
    if ($full.Length -le 76) { return $full }
    $result = "${name}:"
    $line = ""
    foreach ($p in ($value -split ' ')) {
        if ($line -eq "") { $line = " $p" }
        elseif (("$line $p").Length -gt 75) { $result += "$line`n"; $line = " $p" }
        else { $line += " $p" }
    }
    return $result + $line
}

function Sanitise-Filename ([string]$name) {
    $normalized = $name.Normalize([System.Text.NormalizationForm]::FormKC)
    return ($normalized -replace '[^\w\.\-\(\) ]', '_').Trim()
}

function New-Boundary {
    return "----=_Part_$(Get-Random -Minimum 100000 -Maximum 999999)_$(Get-Random)"
}

function Get-MimeType ([string]$filename) {
    $ext = [System.IO.Path]::GetExtension($filename).ToLower()
    switch ($ext) {
        ".pdf"  { return "application/pdf" }
        ".doc"  { return "application/msword" }
        ".docx" { return "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }
        ".xls"  { return "application/vnd.ms-excel" }
        ".xlsx" { return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
        ".ppt"  { return "application/vnd.ms-powerpoint" }
        ".pptx" { return "application/vnd.openxmlformats-officedocument.presentationml.presentation" }
        ".png"  { return "image/png" }
        ".jpg"  { return "image/jpeg" }
        ".jpeg" { return "image/jpeg" }
        ".gif"  { return "image/gif" }
        ".bmp"  { return "image/bmp" }
        ".svg"  { return "image/svg+xml" }
        ".tiff" { return "image/tiff" }
        ".tif"  { return "image/tiff" }
        ".zip"  { return "application/zip" }
        ".7z"   { return "application/x-7z-compressed" }
        ".rar"  { return "application/vnd.rar" }
        ".gz"   { return "application/gzip" }
        ".tar"  { return "application/x-tar" }
        ".txt"  { return "text/plain" }
        ".csv"  { return "text/csv" }
        ".html" { return "text/html" }
        ".htm"  { return "text/html" }
        ".xml"  { return "application/xml" }
        ".json" { return "application/json" }
        ".eml"  { return "message/rfc822" }
        ".msg"  { return "application/vnd.ms-outlook" }
        ".ics"  { return "text/calendar" }
        ".vcf"  { return "text/vcard" }
        ".mp3"  { return "audio/mpeg" }
        ".mp4"  { return "video/mp4" }
        ".wav"  { return "audio/wav" }
        ".avi"  { return "video/x-msvideo" }
        default { return "application/octet-stream" }
    }
}

function Get-StoreFilePath ([object]$store) {
    try   { return $store.FilePath }
    catch { return $null }
}

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [string]$ErrorMessage = "Operation failed",
        [int]$MaxRetries = 3,
        [int]$DelayMs = 500
    )
    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try { return & $ScriptBlock }
        catch {
            $attempt++
            if ($attempt -lt $MaxRetries) {
                Write-Warning "$ErrorMessage (attempt $attempt of $MaxRetries): $_"
                Start-Sleep -Milliseconds $DelayMs
            } else { throw }
        }
    }
}

# =============================================================================
# SANITIZER FUNCTIONS (from mboxsanitizer.ps1)
# =============================================================================

function Decode-Rfc2047 ([string]$header) {
    # Pattern: =?charset?B/Q?encoded-text?=
    $pattern = '=\?([^?]+)\?([BbQq])\?([^?]*)\?='

    return [regex]::Replace($header, $pattern, {
        param($match)

        $charset  = $match.Groups[1].Value
        $encoding = $match.Groups[2].Value.ToUpper()
        $payload  = $match.Groups[3].Value

        if ([string]::IsNullOrEmpty($payload)) {
            return $match.Value
        }

        try {
            $enc = [System.Text.Encoding]::GetEncoding($charset)
        } catch {
            $enc = [System.Text.Encoding]::UTF8
        }

        try {
            [byte[]]$bytes = $null

            if ($encoding -eq 'B') {
                $bytes = [System.Convert]::FromBase64String($payload)
            } else {
                # Quoted-printable (Q)
                $payload = $payload -replace '_', ' '
                $bytes = Decode-QuotedPrintable $payload
            }

            if (-not $bytes -or $bytes.Length -eq 0) {
                return $match.Value
            }

            return $enc.GetString($bytes)
        } catch {
            return $match.Value
        }
    })
}

function Decode-QuotedPrintable ([string]$text) {
    # Remove soft line breaks (= at end of line)
    $text = $text -replace "=`r?`n", ""
    $ms  = [System.IO.MemoryStream]::new()
    $utf = [System.Text.Encoding]::UTF8
    $i   = 0
    while ($i -lt $text.Length) {
        if ($text[$i] -eq '=' -and ($i + 2) -lt $text.Length) {
            $hex = $text.Substring($i + 1, 2)
            try {
                $byte = [System.Convert]::ToByte($hex, 16)
                $ms.WriteByte($byte)
            } catch {
                $literal = $text.Substring($i, 3)
                $litBytes = $utf.GetBytes($literal)
                $ms.Write($litBytes, 0, $litBytes.Length)
            }
            $i += 3
        } else {
            $charBytes = $utf.GetBytes($text[$i].ToString())
            $ms.Write($charBytes, 0, $charBytes.Length)
            $i++
        }
    }
    return $ms.ToArray()
}

function Find-PstStore {
    param([object]$namespace, [string]$pstPath, [string[]]$storePathsBefore)
    $pstFileName = [System.IO.Path]::GetFileName($pstPath)

    foreach ($s in $namespace.Stores) {
        $fp = Get-StoreFilePath $s
        if ($fp -and $fp -eq $pstPath) {
            Write-Host "  Store matched by exact path."
            return $s
        }
    }

    $candidates = [System.Collections.Generic.List[object]]::new()
    foreach ($s in $namespace.Stores) {
        $fp = Get-StoreFilePath $s
        if ($fp -and $fp -like "*.pst") {
            if ([System.IO.Path]::GetFileName($fp) -eq $pstFileName) {
                $candidates.Add($s)
            }
        }
    }
    if ($candidates.Count -eq 1) {
        $matchedPath = Get-StoreFilePath $candidates[0]
        Write-Host "  Store matched by filename: $matchedPath"
        return $candidates[0]
    }

    if ($storePathsBefore) {
        $newStores = [System.Collections.Generic.List[object]]::new()
        foreach ($s in $namespace.Stores) {
            $fp = Get-StoreFilePath $s
            if ($fp -and $fp -notin $storePathsBefore) {
                $newStores.Add($s)
            }
        }
        if ($newStores.Count -eq 1) {
            $matchedPath = Get-StoreFilePath $newStores[0]
            Write-Host "  Store matched by diff (new store): $matchedPath"
            return $newStores[0]
        }
    }

    Write-Warning "Could not match PST store. Stores visible to Outlook:"
    foreach ($s in $namespace.Stores) {
        $fp = Get-StoreFilePath $s
        $displayName = try { $s.DisplayName } catch { "(unknown)" }
        if ($fp) { Write-Warning "  - $displayName  ->  $fp" }
        else { Write-Warning "  - $displayName  ->  (no file path - cloud/Exchange store)" }
    }
    return $null
}

function Test-EmailMatchesFilters {
    param(
        [object]$item,
        $DateFrom,
        $DateTo,
        [string[]]$ExcludeFolders,
        [string[]]$ExcludeSenders,
        [string]$currentFolderName
    )
    if ($ExcludeFolders -and $ExcludeFolders.Count -gt 0) {
        foreach ($exFolder in $ExcludeFolders) {
            if ($currentFolderName -like "*$exFolder*") {
                Write-Verbose "  Skipping: folder '$currentFolderName' matches '$exFolder'"
                return $false
            }
        }
    }
    if ($ExcludeSenders -and $ExcludeSenders.Count -gt 0) {
        $senderEmail = try { $item.SenderEmailAddress } catch { "" }
        $senderName = try { $item.SenderName } catch { "" }
        foreach ($exSender in $ExcludeSenders) {
            if ($senderEmail -like "*$exSender*" -or $senderName -like "*$exSender*") {
                Write-Verbose "  Skipping: sender '$senderEmail' matches '$exSender'"
                return $false
            }
        }
    }
    if ($DateFrom -or $DateTo) {
        $receivedDate = try { $item.ReceivedTime } catch { $null }
        if ($receivedDate) {
            if ($DateFrom -and $receivedDate -lt $DateFrom) {
                Write-Verbose "  Skipping: date $receivedDate before $DateFrom"
                return $false
            }
            if ($DateTo -and $receivedDate -gt $DateTo) {
                Write-Verbose "  Skipping: date $receivedDate after $DateTo"
                return $false
            }
        }
    }
    return $true
}

# =============================================================================
# MIME MESSAGE BUILDER
# =============================================================================

function Build-MimeMessage {
    param(
        [string]$from, [string]$fromName, [string]$to, [string]$cc, [string]$bcc,
        [string]$subject, [string]$dateRfc, [string]$messageId, [string]$inReplyTo,
        [string]$references, [string]$importance, [string]$listUnsubscribe,
        [string]$htmlBody, [string]$plainBody, [object[]]$attachmentList,
        [hashtable]$extraHeaders
    )
    $enc = [System.Text.Encoding]::UTF8
    $sb = [System.Text.StringBuilder]::new(65536)
    $LF = "`n"

    if ($null -eq $attachmentList) { $attachmentList = @() }
    elseif ($attachmentList -isnot [System.Array]) { $attachmentList = @($attachmentList) }
    $attachmentList = @($attachmentList | Where-Object { $_ -is [hashtable] -and $_.ContainsKey('Name') })

    $L = { param([string]$line = "") [void]$sb.Append($line + $LF) }

    # Decode RFC2047 encoded headers
    if ($fromName) { $fromName = Decode-Rfc2047 $fromName }
    if ($subject) { $subject = Decode-Rfc2047 $subject }
    if ($to) { $to = Decode-Rfc2047 $to }
    if ($cc) { $cc = Decode-Rfc2047 $cc }
    if ($bcc) { $bcc = Decode-Rfc2047 $bcc }

    $hasAttachments = $attachmentList.Count -gt 0
    $hasHtml = -not [string]::IsNullOrWhiteSpace($htmlBody)
    $hasPlain = -not [string]::IsNullOrWhiteSpace($plainBody)

    $altBoundary = New-Boundary
    $mixedBoundary = New-Boundary
    $safeFrom = if ($from) { $from } else { "unknown@unknown.invalid" }

    try {
        $parsedDate = [datetime]::Parse($dateRfc)
        $envelopeDate = $parsedDate.ToUniversalTime().ToString("ddd MMM dd HH:mm:ss yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
    } catch {
        $envelopeDate = (Get-Date).ToUniversalTime().ToString("ddd MMM dd HH:mm:ss yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
    }

    & $L "From $safeFrom $envelopeDate"

    if ($fromName -and $fromName -ne $safeFrom) {
        $nameB64 = [Convert]::ToBase64String($enc.GetBytes($fromName))
        & $L "From: =?utf-8?B?${nameB64}?= <$safeFrom>"
    } else { & $L "From: $safeFrom" }

    if ($to) { & $L (Fold-Header "To" $to) }
    if ($cc) { & $L (Fold-Header "Cc" $cc) }
    if ($bcc) { & $L (Fold-Header "Bcc" $bcc) }
    if ($messageId) { & $L "Message-ID: $messageId" }
    if ($inReplyTo) { & $L "In-Reply-To: $inReplyTo" }
    if ($references) { & $L (Fold-Header "References" $references) }

    try {
        $parsedDate = [datetime]::Parse($dateRfc)
        $formattedDate = $parsedDate.ToString("ddd, dd MMM yyyy HH:mm:ss zzz", [System.Globalization.CultureInfo]::InvariantCulture)
        & $L "Date: $formattedDate"
    } catch { & $L "Date: $dateRfc" }

    $subjectB64 = [Convert]::ToBase64String($enc.GetBytes($subject))
    & $L "Subject: =?utf-8?B?${subjectB64}?="

    if ($importance -eq "High") { & $L "X-Priority: 1" }
    if ($importance -eq "Low") { & $L "X-Priority: 5" }

    & $L "MIME-Version: 1.0"
    & $L "X-Mailer: PST2MBOX-PowerShell v$PST2MBOX_VERSION"

    if ($listUnsubscribe) { & $L "List-Unsubscribe: $listUnsubscribe" }

    # Add extra preserved headers (from sanitizer integration)
    if ($extraHeaders) {
        foreach ($key in $extraHeaders.Keys) {
            & $L (Fold-Header $key $extraHeaders[$key])
        }
    }

    if ($hasAttachments) {
        & $L "Content-Type: multipart/mixed;"
        & $L "  boundary=`"$mixedBoundary`""
        & $L
        & $L "--$mixedBoundary"
        if ($hasHtml -or $hasPlain) {
            & $L "Content-Type: multipart/alternative;"
            & $L "  boundary=`"$altBoundary`""
            & $L
        }
    } elseif ($hasHtml -and $hasPlain) {
        & $L "Content-Type: multipart/alternative;"
        & $L "  boundary=`"$altBoundary`""
        & $L
    } elseif ($hasHtml) {
        & $L "Content-Type: text/html; charset=utf-8"
        & $L "Content-Transfer-Encoding: base64"
        & $L
        $escaped = Escape-FromLines $htmlBody
        & $L (Get-Base64 $enc.GetBytes($escaped))
        & $L
        return $sb.ToString()
    } else {
        & $L "Content-Type: text/plain; charset=utf-8"
        & $L "Content-Transfer-Encoding: base64"
        & $L
        $escaped = Escape-FromLines $plainBody
        & $L (Get-Base64 $enc.GetBytes($escaped))
        & $L
        return $sb.ToString()
    }

    if ($hasPlain) {
        & $L "--$altBoundary"
        & $L "Content-Type: text/plain; charset=utf-8"
        & $L "Content-Transfer-Encoding: base64"
        & $L
        $plainEscaped = Escape-FromLines $plainBody
        & $L (Get-Base64 $enc.GetBytes($plainEscaped))
    }
    if ($hasHtml) {
        & $L "--$altBoundary"
        & $L "Content-Type: text/html; charset=utf-8"
        & $L "Content-Transfer-Encoding: base64"
        & $L
        $htmlEscaped = Escape-FromLines $htmlBody
        & $L (Get-Base64 $enc.GetBytes($htmlEscaped))
    }
    if ($hasHtml -or $hasPlain) { & $L "--${altBoundary}--" }

    foreach ($att in $attachmentList) {
        $attNameB64 = [Convert]::ToBase64String($enc.GetBytes($att.Name))
        & $L "--$mixedBoundary"
        & $L "Content-Type: $($att.ContentType);"
        & $L "  name=`"=?utf-8?B?${attNameB64}?=`""
        & $L "Content-Transfer-Encoding: base64"
        & $L "Content-Disposition: attachment;"
        & $L "  filename=`"=?utf-8?B?${attNameB64}?=`""
        & $L
        & $L (Get-Base64 $att.Bytes)
    }
    if ($hasAttachments) { & $L "--${mixedBoundary}--" }
    & $L
    return $sb.ToString()
}

# =============================================================================
# EXTENDED ATTACHMENT SCANNER
# =============================================================================

function Invoke-AttachmentScanner {
    param(
        [string]$FilePath,
        [string]$FileName,
        [byte[]]$Bytes
    )

    $result = @{
        Blocked = $false
        Reason  = ""
        Warnings = [System.Collections.Generic.List[string]]::new()
    }

    # 1. MAGIC BYTE VERIFICATION - Check file signature matches extension
    $magicMatch = Test-MagicBytes -Bytes $Bytes -Extension ([System.IO.Path]::GetExtension($FileName).ToLower())
    if (-not $magicMatch.IsMatch) {
        if ($magicMatch.ExpectedType) {
            $result.Blocked = $true
            $result.Reason = "File type mismatch: extension is $([System.IO.Path]::GetExtension($FileName)) but magic bytes indicate $($magicMatch.ExpectedType)"
            return $result
        } else {
            $result.Warnings.Add("Unknown file format - magic bytes not recognized")
        }
    }

    # 2. OLE/COMPOUND DOCUMENT DETECTION
    $oleResult = Test-OLEObject -Bytes $Bytes
    if ($oleResult.IsOLE) {
        if ($oleResult.EmbeddedExe) {
            $result.Blocked = $true
            $result.Reason = "OLE document contains embedded executable"
            return $result
        }
        if ($oleResult.HasMacros) {
            $result.Warnings.Add("OLE document contains VBA macros")
        }
        $result.Warnings.Add("Compound document detected (OLE/ActiveX possible)")
    }

    # 3. ARCHIVE INSPECTION
    $archiveResult = Test-Archive -Bytes $Bytes -Extension ([System.IO.Path]::GetExtension($FileName).ToLower())
    if ($archiveResult.IsArchive) {
        if ($archiveResult.CompressionRatio -gt 1000) {
            $result.Blocked = $true
            $result.Reason = "Possible archive bomb (compression ratio: $($archiveResult.CompressionRatio):1)"
            return $result
        }
        if ($archiveResult.PasswordProtected) {
            $result.Warnings.Add("Password-protected archive - contents cannot be scanned")
        }
        if ($archiveResult.SuspiciousFiles.Count -gt 0) {
            $result.Warnings.Add("Archive contains suspicious files: $($archiveResult.SuspiciousFiles -join ', ')")
        }
        if ($archiveResult.BlockedFiles.Count -gt 0) {
            $result.Warnings.Add("Archive contains blocked file types: $($archiveResult.BlockedFiles -join ', ')")
        }
    }

    # 4. DANGEROUS PATTERN DETECTION
    $patternResult = Test-DangerousPatterns -Bytes $Bytes -FileName $FileName
    if ($patternResult.Blocked) {
        $result.Blocked = $true
        $result.Reason = $patternResult.Reason
        return $result
    }
    foreach ($w in $patternResult.Warnings) {
        $result.Warnings.Add($w)
    }

    # 5. HTA/SCF/SPECIAL FILE DETECTION
    $specialResult = Test-SpecialFiles -Bytes $Bytes -FileName $FileName
    if ($specialResult.Blocked) {
        $result.Blocked = $true
        $result.Reason = $specialResult.Reason
        return $result
    }

    return $result
}

function Test-MagicBytes {
    param(
        [byte[]]$Bytes,
        [string]$Extension
    )

    $result = @{ IsMatch = $true; ExpectedType = $null }

    if ($Bytes.Length -lt 4) { return $result }

    # Common magic byte signatures
    $signatures = @{
        [byte[]]@(0x50,0x4B,0x03,0x04) = "ZIP/Office (PK)"
        [byte[]]@(0x25,0x50,0x44,0x46) = "PDF"
        [byte[]]@(0x89,0x50,0x4E,0x47) = "PNG"
        [byte[]]@(0xFF,0xD8,0xFF) = "JPEG"
        [byte[]]@(0x47,0x49,0x46,0x38) = "GIF"
        [byte[]]@(0x52,0x49,0x46,0x46) = "RIFF (AVI/WebP)"
        [byte[]]@(0x4D,0x5A) = "Executable (DOS/Windows)"
        [byte[]]@(0xD0,0xCF,0x11,0xE0) = "OLE Compound Document"
    }

    $detectedType = $null
    foreach ($sig in $signatures.GetEnumerator()) {
        $match = $true
        for ($i = 0; $i -lt $sig.Key.Length; $i++) {
            if ($i -ge $Bytes.Length -or $Bytes[$i] -ne $sig.Key[$i]) {
                $match = $false
                break
            }
        }
        if ($match) {
            $detectedType = $sig.Value
            break
        }
    }

    # Check for mismatches
    if ($detectedType) {
        $expectedForExt = @{
            '.doc' = 'OLE'; '.xls' = 'OLE'; '.ppt' = 'OLE'
            '.docx' = 'ZIP'; '.xlsx' = 'ZIP'; '.pptx' = 'ZIP'
            '.pdf' = 'PDF'; '.png' = 'PNG'; '.jpg' = 'JPEG'; '.jpeg' = 'JPEG'
            '.gif' = 'GIF'; '.zip' = 'ZIP'; '.avi' = 'RIFF'
            '.exe' = 'Executable'; '.dll' = 'Executable'; '.scr' = 'Executable'
        }

        $isMismatch = $false
        if ($expectedForExt[$Extension]) {
            $expected = $expectedForExt[$Extension]
            if (($expected -eq 'OLE' -and $detectedType -notlike '*OLE*') -or
                ($expected -eq 'ZIP' -and $detectedType -notlike '*ZIP*') -or
                ($expected -eq 'PDF' -and $detectedType -notlike '*PDF*') -or
                ($expected -eq 'PNG' -and $detectedType -notlike '*PNG*') -or
                ($expected -eq 'JPEG' -and $detectedType -notlike '*JPEG*') -or
                ($expected -eq 'GIF' -and $detectedType -notlike '*GIF*') -or
                ($expected -eq 'Executable' -and $detectedType -notlike '*Executable*') -or
                ($expected -eq 'RIFF' -and $detectedType -notlike '*RIFF*')) {
                $isMismatch = $true
            }
        }

        # Executable masquerading as document
        if ($detectedType -like '*Executable*' -and $Extension -match '^\.(doc|xls|ppt|pdf|png|jpg|jpeg|gif|txt|csv)$') {
            $result.IsMatch = $false
            $result.ExpectedType = $detectedType
            return $result
        }

        if ($isMismatch) {
            $result.IsMatch = $false
            $result.ExpectedType = $detectedType
        }
    }

    return $result
}

function Test-OLEObject {
    param([byte[]]$Bytes)

    $result = @{ IsOLE = $false; EmbeddedExe = $false; HasMacros = $false }

    if ($Bytes.Length -lt 8) { return $result }

    # Check for OLE Compound Document header
    if ($Bytes[0] -eq 0xD0 -and $Bytes[1] -eq 0xCF -and $Bytes[2] -eq 0x11 -and $Bytes[3] -eq 0xE0) {
        $result.IsOLE = $true

        # Look for embedded executable markers within OLE stream
        $oleContent = [System.Text.Encoding]::ASCII.GetString($Bytes, 0, [Math]::Min(512, $Bytes.Length))
        if ($oleContent -match '\.exe|\.dll|\.scr|\.com|\.bat|\.cmd') {
            $result.EmbeddedExe = $true
        }

        # Check for VBA macro markers
        if ($oleContent -match 'VBA|Macro|Module|ThisDocument|ThisWorkbook') {
            $result.HasMacros = $true
        }
    }

    # Also check for VBA project in ZIP-based Office files
    if ($Bytes.Length -gt 4 -and $Bytes[0] -eq 0x50 -and $Bytes[1] -eq 0x4B) {
        $zipContent = [System.Text.Encoding]::ASCII.GetString($Bytes)
        if ($zipContent -match 'vbaProject\.bin|macros\.bin') {
            $result.IsOLE = $true
            $result.HasMacros = $true
        }
    }

    return $result
}

function Test-Archive {
    param(
        [byte[]]$Bytes,
        [string]$Extension
    )

    $result = @{
        IsArchive = $false
        CompressionRatio = 0
        PasswordProtected = $false
        SuspiciousFiles = [System.Collections.Generic.List[string]]::new()
        BlockedFiles = [System.Collections.Generic.List[string]]::new()
    }

    $archiveExts = @('.zip', '.7z', '.rar', '.gz', '.tar')
    if ($archiveExts -notcontains $Extension) {
        # Check ZIP magic even if extension doesn't match
        if ($Bytes.Length -lt 4 -or $Bytes[0] -ne 0x50 -or $Bytes[1] -ne 0x4B -or $Bytes[2] -ne 0x03 -or $Bytes[3] -ne 0x04) {
            return $result
        }
    }

    $result.IsArchive = $true

    # ZIP format detection
    if ($Bytes.Length -ge 4 -and $Bytes[0] -eq 0x50 -and $Bytes[1] -eq 0x4B -and $Bytes[2] -eq 0x03 -and $Bytes[3] -eq 0x04) {
        try {
            $ms = New-Object System.IO.MemoryStream(,$Bytes)
            $zip = New-Object System.IO.Compression.ZipArchive($ms, [System.IO.Compression.ZipArchiveMode]::Read)

            $totalUncompressed = 0
            foreach ($entry in $zip.Entries) {
                $totalUncompressed += $entry.UncompressedLength

                $entryName = $entry.FullName.ToLower()

                # Check for password protection (encrypted entries)
                if ($entry.CompressedLength -gt 0 -and $entry.UncompressedLength -gt 0 -and
                    $entry.CompressedLength -ge $entry.UncompressedLength) {
                    # This might indicate encryption
                }

                # Check for suspicious extensions in archive
                if ($entryName -match '\.(exe|bat|cmd|scr|vbs|js|ps1|hta|scf|msi)$') {
                    $result.BlockedFiles.Add($entry.FullName)
                }
                if ($entryName -match 'password|secret|credential|login' -and $entryName -match '\.(txt|doc|xls)$') {
                    $result.SuspiciousFiles.Add($entry.FullName)
                }
            }

            # Calculate compression ratio
            if ($ms.Length -gt 0) {
                $result.CompressionRatio = [math]::Round($totalUncompressed / $ms.Length, 0)
            }

            $zip.Dispose()
            $ms.Dispose()
        } catch {
            $result.Warnings = "Could not inspect archive: $_"
        }
    }

    return $result
}

function Test-DangerousPatterns {
    param(
        [byte[]]$Bytes,
        [string]$FileName
    )

    $result = @{ Blocked = $false; Reason = ""; Warnings = [System.Collections.Generic.List[string]]::new() }

    # Convert to string for pattern matching (multiple encodings)
    $asciiContent = [System.Text.Encoding]::ASCII.GetString($Bytes)
    $utf8Content = [System.Text.Encoding]::UTF8.GetString($Bytes)

    # Check for script content in non-script files
    $ext = [System.IO.Path]::GetExtension($FileName).ToLower()

    # VBScript patterns
    if ($asciiContent -match 'On\s+Error\s+Resume\s+Next|CreateObject\(|WScript\.|ExecuteGlobal') {
        if ($ext -match '^\.(doc|xls|ppt|pdf|txt)$') {
            $result.Warnings.Add("Contains VBScript-like patterns")
        }
    }

    # PowerShell patterns
    if ($asciiContent -match 'Invoke-Expression|DownloadString|WebClient|ShellExecute|Bypass|EncodedCommand') {
        if ($ext -match '^\.(doc|xls|ppt|pdf|txt|jpg|png)$') {
            $result.Warnings.Add("Contains PowerShell-like patterns")
        }
    }

    # JavaScript patterns in non-JS files
    if ($asciiContent -match 'eval\(|document\.write|ActiveXObject|\.ExecScript|WScript\.Shell') {
        if ($ext -match '^\.(doc|xls|ppt|pdf)$') {
            $result.Warnings.Add("Contains JavaScript/ActiveX patterns")
        }
    }

    # Base64 encoded content (potential dropper)
    $base64Pattern = '[A-Za-z0-9+/]{50,}={0,2}'
    if ($asciiContent -match $base64Pattern) {
        $matches = [regex]::Matches($asciiContent, $base64Pattern)
        foreach ($m in $matches) {
            try {
                $decoded = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($m.Value))
                if ($decoded -match 'cmd\.exe|powershell|wscript|cscript|reg\.exe') {
                    $result.Blocked = $true
                    $result.Reason = "Contains encoded executable commands"
                    return $result
                }
            } catch { }
        }
    }

    # PE header in non-executable (executable appended to document)
    if ($ext -match '^\.(doc|xls|ppt|pdf|jpg|png|gif)$') {
        $peOffset = [Array]::IndexOf($Bytes, 0x4D)
        while ($peOffset -ge 0) {
            if ($peOffset + 2 -lt $Bytes.Length -and $Bytes[$peOffset] -eq 0x4D -and $Bytes[$peOffset + 1] -eq 0x5A) {
                $result.Blocked = $true
                $result.Reason = "PE executable header found within document"
                return $result
            }
            $peOffset = [Array]::IndexOf($Bytes, 0x4D, $peOffset + 1)
        }
    }

    return $result
}

function Test-SpecialFiles {
    param(
        [byte[]]$Bytes,
        [string]$FileName
    )

    $result = @{ Blocked = $false; Reason = "" }

    $ext = [System.IO.Path]::GetExtension($FileName).ToLower()

    # HTA files can execute arbitrary code
    if ($ext -eq '.hta') {
        $result.Blocked = $true
        $result.Reason = "HTA files can execute arbitrary code"
        return $result
    }

    # SCF files can be used for credential theft
    if ($ext -eq '.scf') {
        $result.Blocked = $true
        $result.Reason = "SCF files can be used for credential theft"
        return $result
    }

    # Library files (DLL hijacking)
    if ($ext -match '^\.(dll|ocx|tlb)$') {
        $result.Blocked = $true
        $result.Reason = "Library files can be used for DLL hijacking"
        return $result
    }

    return $result
}

# =============================================================================
# ATTACHMENT EXTRACTOR
# =============================================================================

function Get-Attachments {
    param(
        [object]$mailItem,
        [string]$tempDir,
        [switch]$RestrictAttachments,
        [int]$MaxAttachmentSizeMB = 50
    )
    $result = [System.Collections.Generic.List[hashtable]]::new()
    $attCount = 0
    try { $attCount = $mailItem.Attachments.Count } catch { return $result }
    if ($attCount -eq 0) { return $result }

    # Dangerous extensions to block when -RestrictAttachments is used
    $blockedExtensions = @('.exe', '.bat', '.cmd', '.com', '.pif', '.scr', '.vbs', '.vbe', '.js', '.jse', '.wsf', '.wsh', '.ps1', '.msc', '.msi', '.msp', '.hta', '.scf', '.lnk', '.inf', '.reg', '.sys', '.dll', '.cpl', '.drv', '.msg', '.emf', '.wmf')

    for ($idx = 1; $idx -le $attCount; $idx++) {
        $att = $null
        try { $att = $mailItem.Attachments.Item($idx) } catch { continue }

        try {
            # Skip attachments stored by reference (cloud links - potential SSRF)
            try {
                $attachMethod = $att.PropertyAccessor.GetProperty($MAPI_TAGS.AttachMethod)
                if ($attachMethod -eq $MAPI_TAGS.AttachMethodByReference) {
                    Write-Verbose "  Skipping attachment by reference: $($att.FileName)"
                    continue
                }
            } catch { }

            # Validate filename
            $safeName = Get-SafeFileName -name $att.FileName
            if ([string]::IsNullOrWhiteSpace($safeName)) {
                Write-Verbose "  Skipping attachment with invalid filename: $($att.FileName)"
                continue
            }

            # Check attachment size before saving
            try {
                $attSize = $att.Size
                if ($attSize -gt ($MaxAttachmentSizeMB * 1MB)) {
                    Write-Warning "  Skipping oversized attachment ($([math]::Round($attSize / 1MB, 2))MB): $safeName"
                    continue
                }
            } catch { }

            $tempPath = Join-Path $tempDir $safeName
            $att.SaveAsFile($tempPath)

            if (Test-Path $tempPath) {
                # Verify saved file size matches expected
                $fileInfo = Get-Item $tempPath
                if ($fileInfo.Length -gt ($MaxAttachmentSizeMB * 1MB)) {
                    Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
                    Write-Warning "  Removed oversized file after save: $safeName"
                    continue
                }

                $bytes = [System.IO.File]::ReadAllBytes($tempPath)
                $ext = [System.IO.Path]::GetExtension($safeName).ToLower()

                # SECURITY: Check for dangerous extensions (only when -RestrictAttachments)
                if ($RestrictAttachments -and $blockedExtensions -contains $ext) {
                    Write-Warning "  Blocked dangerous extension: $safeName"
                    Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
                    continue
                }

                # Run extended scanner if -RestrictAttachments is specified
                if ($RestrictAttachments) {
                    $scanResult = Invoke-AttachmentScanner -FilePath $tempPath -FileName $safeName -Bytes $bytes
                    if ($scanResult.Blocked) {
                        Write-Warning "  Blocked by scanner: $($safeName) - $($scanResult.Reason)"
                        Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
                        continue
                    }
                    if ($scanResult.Warnings.Count -gt 0) {
                        foreach ($w in $scanResult.Warnings) {
                            Write-Warning "  Scanner warning: $safeName - $w"
                        }
                    }
                }

                $mime = Get-MimeType $safeName
                $result.Add(@{
                    Name        = $att.FileName
                    SafeName    = $safeName
                    Bytes       = $bytes
                    ContentType = $mime
                    Size        = $fileInfo.Length
                })

                Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
            }
        } catch {
            Write-Warning "  Could not extract attachment '$($att.FileName)': $_"
        }
    }
    return $result
}

# =============================================================================
# FOLDER WALKER
# =============================================================================

$script:totalEmails = 0
$script:skippedItems = 0
$script:errorEmails = 0
$script:totalFolders = 0
$script:tempDirs = [System.Collections.Generic.List[string]]::new()
$script:emailCount = 0
$script:failedItems = [System.Collections.Generic.List[object]]::new()
$script:currentFileIndex = 1
$script:currentFileSize = 0
$script:seenMessageIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:duplicateCount = 0

function Export-FolderToMbox {
    param([object]$folder, [System.IO.StreamWriter]$writer, [string]$indent = "", [string]$currentMboxPath)

    if ($MaxEmails -gt 0 -and $script:emailCount -ge $MaxEmails) { return }

    $script:totalFolders++
    $folderName = $folder.Name
    $folderCount = $folder.Items.Count
    Write-Host "${indent}Folder: $folderName ($folderCount items)"

    for ($i = 1; $i -le $folderCount; $i++) {
        if ($MaxEmails -gt 0 -and $script:emailCount -ge $MaxEmails) { return }

        $item = $null
        try { $item = Invoke-WithRetry -ScriptBlock { $folder.Items.Item($i) } -ErrorMessage "Failed to access folder item" }
        catch { $script:skippedItems++; continue }

        if ($item.MessageClass -notlike "IPM.Note*") { $script:skippedItems++; continue }

        if (-not (Test-EmailMatchesFilters -item $item -DateFrom $DateFrom -DateTo $DateTo `
                -ExcludeFolders $ExcludeFolders -ExcludeSenders $ExcludeSenders -currentFolderName $folderName)) {
            $script:skippedItems++
            Write-Verbose "  Filtered out: $($item.Subject)"
            continue
        }

        try {
            $senderEmail = $item.SenderEmailAddress
            if ($item.SenderEmailType -eq "EX") {
                try {
                    $exchUser = $item.Sender.GetExchangeUser()
                    if ($exchUser) { $senderEmail = $exchUser.PrimarySmtpAddress }
                } catch { }
            }

            $ccStr = try { $item.CC } catch { "" }
            $bccStr = try { $item.BCC } catch { "" }
            $msgId = try { $item.PropertyAccessor.GetProperty($MAPI_TAGS.MessageId) } catch { "" }
            $inReplyTo = try { $item.PropertyAccessor.GetProperty($MAPI_TAGS.InReplyTo) } catch { "" }
            $references = try { $item.PropertyAccessor.GetProperty($MAPI_TAGS.References) } catch { "" }
            $importance = try { switch ($item.Importance) { 2 { "High" } 0 { "Low" } default { "" } } } catch { "" }
            $listUnsubscribe = try { $item.PropertyAccessor.GetProperty($MAPI_TAGS.ListUnsubscribe) } catch { "" }
            $htmlText = try { $item.HTMLBody } catch { "" }
            $plainText = try { $item.Body } catch { "" }
            $sentDate = try { $item.SentOn.ToString("R") } catch { (Get-Date).ToString("R") }
            $subj = try { $item.Subject } catch { "(no subject)" }

            $attachments = @(Get-Attachments $item $tempDir -RestrictAttachments:$RestrictAttachments -MaxAttachmentSizeMB $MaxAttachmentSizeMB)

            # Build extra headers hash from PreserveHeaders parameter
            # SECURITY: Validate MAPI property tags against allowlist
            $allowedMapiTags = @(
                'http://schemas.microsoft.com/mapi/proptag/0x007F001E',  # List-Unsubscribe
                'http://schemas.microsoft.com/mapi/proptag/0x001A001E',  # Message-Class
                'http://schemas.microsoft.com/mapi/proptag/0x0037001E',  # Subject
                'http://schemas.microsoft.com/mapi/proptag/0x003A001E',  # Sender-Name
                'http://schemas.microsoft.com/mapi/proptag/0x0065001E',  # Sender-Address-Type
                'http://schemas.microsoft.com/mapi/proptag/0x0070001E',  # Sender-Email
                'http://schemas.microsoft.com/mapi/proptag/0x0E04001E',  # Sender-EntryID
                'http://schemas.microsoft.com/mapi/proptag/0x0E05001E',  # Sender-Search-Key
                'http://schemas.microsoft.com/mapi/proptag/0x0E06001E',  # Sender-Address-EntryID
                'http://schemas.microsoft.com/mapi/proptag/0x1081001E',  # Sender-SMTP-Address
                'http://schemas.microsoft.com/mapi/proptag/0x1086001E',  # Sender-Original-Display-Name
                'http://schemas.microsoft.com/mapi/proptag/0x1087001E',  # Sender-Original-EntryID
                'http://schemas.microsoft.com/mapi/proptag/0x1088001E',  # Sender-Original-Search-Key
                'http://schemas.microsoft.com/mapi/proptag/0x1089001E',  # Sender-Original-Address-EntryID
                'http://schemas.microsoft.com/mapi/proptag/0x108A001E',  # Sender-Original-Email
                'http://schemas.microsoft.com/mapi/proptag/0x108B001E',  # Sender-Original-Address-Type
                'http://schemas.microsoft.com/mapi/proptag/0x108C001E',  # Sender-Original-EntryID-EntryID
                'http://schemas.microsoft.com/mapi/proptag/0x108D001E',  # Sender-Original-Search-Key-EntryID
                'http://schemas.microsoft.com/mapi/proptag/0x108E001E',  # Sender-Original-Address-EntryID-EntryID
                'http://schemas.microsoft.com/mapi/proptag/0x108F001E',  # Sender-Original-Email-EntryID
                'http://schemas.microsoft.com/mapi/proptag/0x1090001E',  # Sender-Original-Address-Type-EntryID
                'http://schemas.microsoft.com/mapi/proptag/0x1091001E'   # Sender-Original-EntryID-Search-Key
            )
            $extraHeaders = @{}
            if ($PreserveHeaders -and $PreserveHeaders.Count -gt 0) {
                foreach ($hName in $PreserveHeaders) {
                    # SECURITY: Validate MAPI tag format
                    if ($hName -notmatch '^http://schemas\.microsoft\.com/mapi/proptag/0x[0-9A-F]{8}001E$') {
                        Write-Verbose "  Skipping invalid MAPI tag format: $hName"
                        continue
                    }
                    # SECURITY: Check against allowlist
                    if ($allowedMapiTags -notcontains $hName) {
                        Write-Verbose "  Skipping non-allowlisted MAPI tag: $hName"
                        continue
                    }
                    try {
                        $propValue = $item.PropertyAccessor.GetProperty($hName)
                        if ($propValue) { $extraHeaders[$hName] = $propValue }
                    } catch { }
                }
            }

            # Check for duplicate Message-ID if DeDuplicate is enabled
            if ($DeDuplicate -and $msgId -and -not [string]::IsNullOrWhiteSpace($msgId)) {
                if (-not $script:seenMessageIds.Add($msgId)) {
                    $script:duplicateCount++
                    Write-Verbose "  Skipping duplicate: $subj (Message-ID: $msgId)"
                    continue
                }
            }

            $mimeText = Build-MimeMessage -from $senderEmail -fromName $item.SenderName -to $item.To `
                -cc $ccStr -bcc $bccStr -subject $subj -dateRfc $sentDate -messageId $msgId `
                -inReplyTo $inReplyTo -references $references -importance $importance `
                -listUnsubscribe $listUnsubscribe -htmlBody $htmlText -plainBody $plainText `
                -attachmentList $attachments -extraHeaders $extraHeaders

            $writer.Write($mimeText)
            $script:currentFileSize += [System.Text.Encoding]::UTF8.GetByteCount($mimeText)

            $script:totalEmails++
            $script:emailCount++

            if ($script:totalEmails % 100 -eq 0) {
                Write-Host "  ... $($script:totalEmails) emails written" -ForegroundColor DarkGray
                $writer.Flush()
            }
            if ($script:totalEmails % 1000 -eq 0) {
                $elapsed = $stopwatch.Elapsed.TotalSeconds
                $rate = [math]::Round($script:totalEmails / $elapsed, 2)
                Write-Host "  [PROGRESS] $($script:totalEmails) emails | ${elapsed}s | ${rate}/sec" -ForegroundColor Cyan
            }

            if ($SplitSizeMB -gt 0) {
                $maxBytes = $SplitSizeMB * 1MB
                if ($script:currentFileSize -ge $maxBytes) {
                    Write-Host "  Splitting file at ${SplitSizeMB}MB..." -ForegroundColor Yellow
                    $writer.Flush()
                    $writer.Close()

                    $baseName = [System.IO.Path]::GetBaseName($currentMboxPath)
                    $dirName = [System.IO.Path]::GetDirectoryName($currentMboxPath)
                    $script:currentFileIndex++
                    $nextMboxPath = Join-Path $dirName "${baseName}.mbox.$('{0:D3}' -f $script:currentFileIndex)"

                    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
                    $writer = New-Object System.IO.StreamWriter($nextMboxPath, $false, $utf8NoBom)
                    $writer.NewLine = "`n"
                    $script:currentFileSize = 0
                    Write-Host "  Writing to: $nextMboxPath" -ForegroundColor DarkGray
                }
            }
        } catch {
            $subjFallback = try { $item.Subject } catch { "(unknown)" }
            # SECURITY: Sanitize error message to prevent path disclosure
            $safeError = $_.Exception.Message -replace [regex]::Escape([System.IO.Path]::GetFullPath(".")), "[REDACTED]"
            $safeError = $safeError -replace [regex]::Escape($env:TEMP), "[TEMP]"
            Write-Warning "${indent}  Failed to process '$subjFallback': $safeError"
            $script:errorEmails++

            $senderVal = try { $item.SenderEmailAddress } catch { "unknown" }
            $receivedVal = try { $item.ReceivedTime.ToString() } catch { "unknown" }

            $script:failedItems.Add([PSCustomObject]@{
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Folder = $folderName
                Subject = $subjFallback
                Sender = $senderVal
                Received = $receivedVal
                Error = $safeError
            })
        }
    }

    if ($MaxEmails -eq 0 -or $script:emailCount -lt $MaxEmails) {
        foreach ($sub in $folder.Folders) {
            Export-FolderToMbox -folder $sub -writer $writer -indent "$indent  " -currentMboxPath $currentMboxPath
        }
    }
}

# =============================================================================
# MAIN
# =============================================================================

Write-Host "=== PST to MBOX Converter v$PST2MBOX_VERSION ===" -ForegroundColor Cyan
Write-Host "Input : $PstPath"
Write-Host "Output: $MboxPath"

$hasFilters = $DateFrom -or $DateTo -or $ExcludeFolders -or $ExcludeSenders
if ($hasFilters) {
    Write-Host "Filters:" -ForegroundColor Yellow
    if ($DateFrom) { Write-Host "  Date From: $DateFrom" }
    if ($DateTo) { Write-Host "  Date To: $DateTo" }
    if ($ExcludeFolders) { Write-Host "  Exclude Folders: $($ExcludeFolders -join ', ')" }
    if ($ExcludeSenders) { Write-Host "  Exclude Senders: $($ExcludeSenders -join ', ')" }
}
if ($MaxEmails -gt 0) { Write-Host "Max emails: $MaxEmails" -ForegroundColor Yellow }
if ($SplitSizeMB -gt 0) { Write-Host "Split size: ${SplitSizeMB}MB" -ForegroundColor Yellow }
if ($WhatIf) { Write-Host "Mode: DRY-RUN (no output written)" -ForegroundColor Yellow }
Write-Host ""

if (-not (Test-Path $PstPath)) { Write-Error "PST file not found: $PstPath"; exit 1 }

if (-not $WhatIf -and (Test-Path $MboxPath)) { Remove-Item $MboxPath -Force }

function Show-DialogWarning ([string]$phase) {
    Write-Host ""
    Write-Warning "Outlook appears stuck during: $phase"
    Write-Warning "This usually means a modal dialog is waiting for your input."
    Write-Warning "Please check:"
    Write-Warning "  1. The Windows TASKBAR for a flashing Outlook icon"
    Write-Warning "  2. Press Alt+Tab to find hidden windows"
    Write-Warning "  3. Check the system tray (bottom-right near the clock)"
    Write-Warning "  4. Look behind this PowerShell window"
    Write-Warning ""
    Write-Warning "Common dialogs that block:"
    Write-Warning "  - Profile selection ('Choose Profile')"
    Write-Warning "  - Recovery prompt ('Outlook was not closed properly')"
    Write-Warning "  - Security dialog ('A program is trying to access...')"
    Write-Warning "  - PST repair dialog"
    Write-Warning "  - License / activation prompt"
    Write-Warning ""
    Write-Host "Waiting for you to dismiss the dialog..." -ForegroundColor Yellow
}

Write-Host "Starting Outlook COM..."
$outlook = $null
$namespace = $null
$comReady = $false
$comTimeout = 30

$runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
$runspace.ApartmentState = [System.Threading.ApartmentState]::STA
$runspace.Open()

$ps = [System.Management.Automation.PowerShell]::Create()
$ps.Runspace = $runspace
[void]$ps.AddScript({
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $null = $ns.Folders.Count
    return @{ Outlook = $ol; Namespace = $ns }
})

$asyncResult = $ps.BeginInvoke()
$waited = 0
while (-not $asyncResult.IsCompleted -and $waited -lt $comTimeout) {
    Start-Sleep -Milliseconds 500
    $waited += 0.5
    if ($waited % 5 -eq 0) { Write-Host "  ... still waiting ($waited sec)" -ForegroundColor DarkGray }
}

if ($asyncResult.IsCompleted) {
    try { $result = $ps.EndInvoke($asyncResult); $comReady = $true }
    catch { Write-Warning "Outlook COM initialisation error: $_" }
}

$ps.Dispose()
$runspace.Close()
$runspace.Dispose()

if (-not $comReady) { Show-DialogWarning "Outlook COM initialisation" }

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

$probeStart = [System.Diagnostics.Stopwatch]::StartNew()
$probeOk = $false
while ($probeStart.Elapsed.TotalSeconds -lt 15) {
    try { $null = $namespace.Folders.Count; $probeOk = $true; break }
    catch { Start-Sleep -Milliseconds 500 }
}

if (-not $probeOk) {
    Show-DialogWarning "MAPI namespace access"
    while ($true) {
        try { $null = $namespace.Folders.Count; break }
        catch { Start-Sleep -Seconds 1 }
    }
}

Write-Host "Outlook is responsive." -ForegroundColor Green

$storePathsBefore = [System.Collections.Generic.List[string]]::new()
foreach ($s in $namespace.Stores) {
    $fp = Get-StoreFilePath $s
    if ($fp) { $storePathsBefore.Add($fp) }
}

Write-Host "Mounting PST..."
$namespace.AddStore($PstPath)
Start-Sleep -Seconds 2

$mountProbeOk = $false
$mountProbeStart = [System.Diagnostics.Stopwatch]::StartNew()
while ($mountProbeStart.Elapsed.TotalSeconds -lt 10) {
    try { $null = $namespace.Stores.Count; $mountProbeOk = $true; break }
    catch { Start-Sleep -Milliseconds 500 }
}

if (-not $mountProbeOk) {
    Show-DialogWarning "PST mount (AddStore)"
    while ($true) {
        try { $null = $namespace.Stores.Count; break }
        catch { Start-Sleep -Seconds 1 }
    }
    Write-Host "Outlook is responsive after mount. Continuing..." -ForegroundColor Green
}

$store = Find-PstStore -namespace $namespace -pstPath $PstPath -storePathsBefore $storePathsBefore.ToArray()

if (-not $store) { Write-Error "Could not locate mounted PST store. See store list above."; exit 1 }

$rootFolder = $store.GetRootFolder()

if ($WhatIf) {
    Write-Host ""
    Write-Host "=== DRY-RUN MODE ===" -ForegroundColor Cyan
    Write-Host "Scanning PST without writing output..."
    Write-Host ""

    $dryRunCount = 0
    $dryRunSkipped = 0

    function Scan-Folder {
        param([object]$folder, [string]$indent = "")
        $folderCount = $folder.Items.Count
        Write-Host "${indent}Folder: $($folder.Name) ($folderCount items)"

        for ($i = 1; $i -le $folderCount; $i++) {
            $item = $null
            try { $item = $folder.Items.Item($i) } catch { continue }

            if ($item.MessageClass -notlike "IPM.Note*") { $dryRunSkipped++; continue }

            if (Test-EmailMatchesFilters -item $item -DateFrom $DateFrom -DateTo $DateTo `
                    -ExcludeFolders $ExcludeFolders -ExcludeSenders $ExcludeSenders `
                    -currentFolderName $folder.Name) {
                $dryRunCount++
            } else { $dryRunSkipped++ }
        }

        foreach ($sub in $folder.Folders) { Scan-Folder -folder $sub -indent "$indent  " }
    }

    Scan-Folder -folder $rootFolder

    Write-Host ""
    Write-Host "=== DRY-RUN Summary ===" -ForegroundColor Cyan
    Write-Host "Emails that would be exported: $dryRunCount"
    Write-Host "Items that would be skipped:  $dryRunSkipped"
    Write-Host "Total folders scanned:        $($script:totalFolders)"

    $namespace.RemoveStore($rootFolder) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
    [System.GC]::Collect()
    exit 0
}

# SECURITY: Create secure temporary directory
$tempDir = Join-Path $env:TEMP "pst2mbox_$([System.IO.Path]::GetRandomFileName())"

# Verify TEMP is on a local drive (prevent network path attacks)
try {
    $tempDrive = [System.IO.Path]::GetPathRoot($env:TEMP)
    $driveType = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='$tempDrive'" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DriveType
    if ($driveType -ne 3) {
        Write-Warning "TEMP directory is not on a local fixed drive. Using fallback."
        $tempDir = [System.IO.Path]::GetTempPath()
    }
} catch {
    Write-Verbose "Could not verify drive type: $_"
}

try {
    $tempDir = [System.IO.Path]::Combine($tempDir, "pst2mbox_$([System.IO.Path]::GetRandomFileName())")
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    # SECURITY: Set restrictive ACL on temp directory (current user only)
    try {
        $acl = Get-Acl $tempDir
        $acl.SetAccessRuleProtection($true, $false)
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            [System.Security.Principal.WindowsIdentity]::GetCurrent().Name,
            "FullControl",
            "Allow")
        $acl.AddAccessRule($rule)
        Set-Acl $tempDir $acl -ErrorAction SilentlyContinue
    } catch {
        Write-Verbose "Could not set restrictive ACL on temp directory: $_"
    }

    $script:tempDirs.Add($tempDir)
} catch {
    throw "Failed to create secure temporary directory: $_"
}

$utf8NoBom = New-Object System.Text.UTF8Encoding $false
$writer = New-Object System.IO.StreamWriter($MboxPath, $false, $utf8NoBom)
$writer.NewLine = "`n"

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    Write-Host "Exporting..." -ForegroundColor Green
    Export-FolderToMbox -folder $rootFolder -writer $writer -currentMboxPath $MboxPath
} finally {
    $writer.Flush()
    $writer.Close()
    $stopwatch.Stop()

    Write-Host ""
    Write-Host "Unmounting PST..."
    try { $namespace.RemoveStore($rootFolder) } catch { }

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null

    if ($script:tempDirs) {
        Write-Host "Cleaning up temporary files..."
        foreach ($dir in $script:tempDirs) {
            if (Test-Path $dir) { Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }

    if ($FailedLogPath -and $script:failedItems.Count -gt 0) {
        # SECURITY: Validate log path
        try {
            $FailedLogPath = Test-SafePath -Path $FailedLogPath -BasePath $PWD.Path -Type "log"
        } catch {
            Write-Warning "FailedLogPath rejected: $_"
            $FailedLogPath = $null
        }
        if ($FailedLogPath) {
            Write-Host "Writing failed items log to: $FailedLogPath"
            $failedLogDir = [System.IO.Path]::GetDirectoryName($FailedLogPath)
            if ($failedLogDir -and -not (Test-Path $failedLogDir)) {
                New-Item -ItemType Directory -Path $failedLogDir -Force | Out-Null
            }
            # SECURITY: Use restricted file permissions (owner read/write only)
            $script:failedItems | Out-File -FilePath $FailedLogPath -Encoding UTF8
            try {
                $acl = Get-Acl $FailedLogPath
                $acl.SetAccessRuleProtection($true, $false)
                $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                    [System.Security.Principal.WindowsIdentity]::GetCurrent().Name,
                    "Read,Write",
                    "Allow")
                $acl.AddAccessRule($rule)
                Set-Acl $FailedLogPath $acl -ErrorAction SilentlyContinue
            } catch {
                Write-Verbose "Could not set restrictive ACL on log file: $_"
            }
        }
    }

    [System.GC]::Collect()
}

Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Cyan
Write-Host "Emails exported : $($script:totalEmails)"
Write-Host "Items skipped   : $($script:skippedItems)  (calendars, contacts, tasks, etc.)"
Write-Host "Errors          : $($script:errorEmails)"
Write-Host "Elapsed         : $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"
Write-Host "MBOX written to : $MboxPath"

if ($SplitSizeMB -gt 0 -and $script:currentFileIndex -gt 1) {
    Write-Host "Split files     : $script:currentFileIndex files created" -ForegroundColor Yellow
}

if ($script:failedItems.Count -gt 0) {
    if ($FailedLogPath) { Write-Host "Failed items    : Logged to $FailedLogPath" -ForegroundColor Yellow }
    else { Write-Host "Failed items    : $($script:failedItems.Count) errors (use -FailedLogPath to save)" -ForegroundColor Yellow }
}

$script:completedNormally = $true