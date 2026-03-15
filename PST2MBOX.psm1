# =============================================================================
# PST2MBOX PowerShell Module
# =============================================================================
$PST2MBOX_VERSION = '2.2.0'

$MAPI_TAGS = @{
    MessageId      = 'http://schemas.microsoft.com/mapi/proptag/0x1035001E'
    InReplyTo      = 'http://schemas.microsoft.com/mapi/proptag/0x1042001E'
    References     = 'http://schemas.microsoft.com/mapi/proptag/0x1039001E'
    AttachMethod   = 'http://schemas.microsoft.com/mapi/proptag/0x37050003'
    AttachMethodByReference = 6
    ListUnsubscribe = 'http://schemas.microsoft.com/mapi/proptag/0x007F001E'
}

function Get-Base64 { param([byte[]]$bytes)
    if (-not $bytes -or $bytes.Length -eq 0) { return "" }
    return ([System.Convert]::ToBase64String($bytes, [System.Base64FormattingOptions]::InsertLineBreaks)) -replace "`r", ""
}

function Escape-FromLines { param([string]$text)
    return ($text -split "`r?`n" | ForEach-Object { if ($_ -match '^From ') { ">$_" } else { $_ } }) -join "`n"
}

function Fold-Header { param([string]$name, [string]$value)
    $full = "${name}: ${value}"
    if ($full.Length -le 76) { return $full }
    $result = "${name}:"; $line = ""
    foreach ($p in ($value -split ' ')) {
        if ($line -eq "") { $line = " $p" }
        elseif (("$line $p").Length -gt 75) { $result += "$line`n"; $line = " $p" }
        else { $line += " $p" }
    }
    return $result + $line
}

function Sanitise-Filename { param([string]$name)
    return ($name -replace '[^\w\.\-\(\) ]', '_').Trim()
}

function New-Boundary {
    return "----=_Part_$(Get-Random -Minimum 100000 -Maximum 999999)_$(Get-Random)"
}

function Get-MimeType { param([string]$filename)
    $ext = [System.IO.Path]::GetExtension($filename).ToLower()
    switch ($ext) {
        ".pdf" { return "application/pdf" } ".doc" { return "application/msword" }
        ".docx" { return "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }
        ".xls" { return "application/vnd.ms-excel" } ".xlsx" { return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
        ".ppt" { return "application/vnd.ms-powerpoint" } ".pptx" { return "application/vnd.openxmlformats-officedocument.presentationml.presentation" }
        ".png" { return "image/png" } ".jpg" { return "image/jpeg" } ".jpeg" { return "image/jpeg" }
        ".gif" { return "image/gif" } ".bmp" { return "image/bmp" } ".svg" { return "image/svg+xml" }
        ".tiff" { return "image/tiff" } ".tif" { return "image/tiff" }
        ".zip" { return "application/zip" } ".7z" { return "application/x-7z-compressed" }
        ".rar" { return "application/vnd.rar" } ".gz" { return "application/gzip" } ".tar" { return "application/x-tar" }
        ".txt" { return "text/plain" } ".csv" { return "text/csv" }
        ".html" { return "text/html" } ".htm" { return "text/html" }
        ".xml" { return "application/xml" } ".json" { return "application/json" }
        ".eml" { return "message/rfc822" } ".ics" { return "text/calendar" } ".vcf" { return "text/vcard" }
        ".mp3" { return "audio/mpeg" } ".mp4" { return "video/mp4" }
        ".wav" { return "audio/wav" } ".avi" { return "video/x-msvideo" }
        default { return "application/octet-stream" }
    }
}

function Get-StoreFilePath { param([object]$store)
    try { return $store.FilePath } catch { return $null }
}

function Invoke-WithRetry {
    param([scriptblock]$ScriptBlock, [string]$ErrorMessage = "Operation failed", [int]$MaxRetries = 3, [int]$DelayMs = 500)
    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try { return & $ScriptBlock } catch {
            $attempt++
            if ($attempt -lt $MaxRetries) { Write-Warning "$ErrorMessage (attempt $attempt of $MaxRetries): $_"; Start-Sleep -Milliseconds $DelayMs }
            else { throw }
        }
    }
}

function Find-PstStore {
    param([object]$namespace, [string]$pstPath, [string[]]$storePathsBefore)
    $pstFileName = [System.IO.Path]::GetFileName($pstPath)
    foreach ($s in $namespace.Stores) {
        $fp = Get-StoreFilePath $s
        if ($fp -and $fp -eq $pstPath) { Write-Host "  Store matched by exact path."; return $s }
    }
    $candidates = [System.Collections.Generic.List[object]]::new()
    foreach ($s in $namespace.Stores) {
        $fp = Get-StoreFilePath $s
        if ($fp -and $fp -like "*.pst") {
            if ([System.IO.Path]::GetFileName($fp) -eq $pstFileName) { $candidates.Add($s) }
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
            if ($fp -and $fp -notin $storePathsBefore) { $newStores.Add($s) }
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

function Decode-Rfc2047 { param([string]$header)
    $pattern = '=\?([^?]+)\?([BbQq])\?([^?]*)\?='
    return [regex]::Replace($header, $pattern, {
        param($match)
        $charset = $match.Groups[1].Value; $encoding = $match.Groups[2].Value.ToUpper(); $payload = $match.Groups[3].Value
        if ([string]::IsNullOrEmpty($payload)) { return $match.Value }
        try { $enc = [System.Text.Encoding]::GetEncoding($charset) } catch { $enc = [System.Text.Encoding]::UTF8 }
        try {
            [byte[]]$bytes = $null
            if ($encoding -eq 'B') { $bytes = [System.Convert]::FromBase64String($payload) }
            else { $payload = $payload -replace '_', ' '; $bytes = Decode-QuotedPrintable -text $payload }
            if (-not $bytes -or $bytes.Length -eq 0) { return $match.Value }
            return $enc.GetString($bytes)
        } catch { return $match.Value }
    })
}

function Decode-QuotedPrintable { param([string]$text)
    $text = $text -replace "=`r?`n", ""
    $ms = [System.IO.MemoryStream]::new(); $utf = [System.Text.Encoding]::UTF8; $i = 0
    while ($i -lt $text.Length) {
        if ($text[$i] -eq '=' -and ($i + 2) -lt $text.Length) {
            $hex = $text.Substring($i + 1, 2)
            try { $byte = [System.Convert]::ToByte($hex, 16); $ms.WriteByte($byte) }
            catch { $literal = $text.Substring($i, 3); $litBytes = $utf.GetBytes($literal); $ms.Write($litBytes, 0, $litBytes.Length) }
            $i += 3
        } else { $charBytes = $utf.GetBytes($text[$i].ToString()); $ms.Write($charBytes, 0, $charBytes.Length); $i++ }
    }
    return $ms.ToArray()
}

function Test-SafePath {
    param([string]$Path, [string]$BasePath, [string]$Type = "file")
    $resolved = [System.IO.Path]::GetFullPath($Path)
    if ($BasePath) {
        $baseResolved = [System.IO.Path]::GetFullPath($BasePath)
        if (-not $resolved.StartsWith($baseResolved, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Security: $Type path '$Path' resolves outside allowed base directory"
        }
    }
    if ($resolved -match '[;|&]') { throw "Security: Invalid characters detected in path" }
    return $resolved
}

function Get-SafeFileName { param([string]$name)
    $name = ($name -replace '[\\/]', '_') -replace "`0", ''
    $name = $name -replace '^\.+', ''
    if ($name.Length -gt 255) {
        $ext = [System.IO.Path]::GetExtension($name)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
        if ($ext.Length -lt 20) { $name = $base.Substring(0, 255 - $ext.Length) + $ext }
        else { $name = $base.Substring(0, 250) }
    }
    return (Sanitise-Filename $name).Trim()
}

function Test-SafeFileExtension { param([string]$FileName)
    $safeExtensions = @('.pdf','.doc','.docx','.xls','.xlsx','.ppt','.pptx','.png','.jpg','.jpeg','.gif','.bmp','.tiff','.tif','.svg','.zip','.7z','.rar','.gz','.tar','.txt','.csv','.html','.htm','.xml','.json','.eml','.ics','.vcf','.mp3','.mp4','.wav','.avi','.mkv','.mov','.rtf','.odt','.ods','.odp')
    $ext = [System.IO.Path]::GetExtension($FileName).ToLower()
    if ([string]::IsNullOrWhiteSpace($ext)) { return $true }
    return $safeExtensions -contains $ext
}

function Test-EmailMatchesFilters {
    param([object]$item, [datetime]$DateFrom, [datetime]$DateTo, [string[]]$ExcludeFolders, [string[]]$ExcludeSenders, [string]$currentFolderName)
    if ($ExcludeFolders -and $ExcludeFolders.Count -gt 0) {
        foreach ($exFolder in $ExcludeFolders) { if ($currentFolderName -like "*$exFolder*") { return $false } }
    }
    if ($ExcludeSenders -and $ExcludeSenders.Count -gt 0) {
        $senderEmail = try { $item.SenderEmailAddress } catch { "" }
        $senderName = try { $item.SenderName } catch { "" }
        foreach ($exSender in $ExcludeSenders) { if ($senderEmail -like "*$exSender*" -or $senderName -like "*$exSender*") { return $false } }
    }
    if ($DateFrom -or $DateTo) {
        $receivedDate = try { $item.ReceivedTime } catch { $null }
        if ($receivedDate) {
            if ($DateFrom -and $receivedDate -lt $DateFrom) { return $false }
            if ($DateTo -and $receivedDate -gt $DateTo) { return $false }
        }
    }
    return $true
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

function Get-Attachments {
    param([object]$mailItem, [string]$tempDir, [switch]$RestrictAttachments, [int]$MaxAttachmentSizeMB = 50)
    $result = [System.Collections.Generic.List[hashtable]]::new()
    $attCount = 0; try { $attCount = $mailItem.Attachments.Count } catch { return $result }
    if ($attCount -eq 0) { return $result }

    # Dangerous extensions to block when -RestrictAttachments is used
    $blockedExtensions = @('.exe', '.bat', '.cmd', '.com', '.pif', '.scr', '.vbs', '.vbe', '.js', '.jse', '.wsf', '.wsh', '.ps1', '.msc', '.msi', '.msp', '.hta', '.scf', '.lnk', '.inf', '.reg', '.sys', '.dll', '.cpl', '.drv', '.msg', '.emf', '.wmf')

    for ($idx = 1; $idx -le $attCount; $idx++) {
        $att = $null; try { $att = $mailItem.Attachments.Item($idx) } catch { continue }
        try {
            try {
                $attachMethod = $att.PropertyAccessor.GetProperty($MAPI_TAGS.AttachMethod)
                if ($attachMethod -eq $MAPI_TAGS.AttachMethodByReference) { continue }
            } catch { }
            $safeName = Get-SafeFileName -name $att.FileName
            if ([string]::IsNullOrWhiteSpace($safeName)) { continue }
            try { $attSize = $att.Size; if ($attSize -gt ($MaxAttachmentSizeMB * 1MB)) { continue } } catch { }
            $tempPath = Join-Path $tempDir $safeName; $att.SaveAsFile($tempPath)
            if (Test-Path $tempPath) {
                $fileInfo = Get-Item $tempPath
                if ($fileInfo.Length -gt ($MaxAttachmentSizeMB * 1MB)) { Remove-Item $tempPath -Force; continue }
                $bytes = [System.IO.File]::ReadAllBytes($tempPath)
                $ext = [System.IO.Path]::GetExtension($safeName).ToLower()

                # SECURITY: Block dangerous extensions (only when -RestrictAttachments)
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
                $result.Add(@{ Name = $att.FileName; SafeName = $safeName; Bytes = $bytes; ContentType = $mime; Size = $fileInfo.Length })
                Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
            }
        } catch { Write-Warning "  Could not extract attachment '$($att.FileName)': $_" }
    }
    return $result
}

function Get-ExportStatistics {
    param([string]$MboxPath)
    if (-not (Test-Path $MboxPath)) { throw "MBOX file not found: $MboxPath" }
    $content = Get-Content -Path $MboxPath -Raw -Encoding UTF8; $lines = $content -split "`n"
    $stats = @{ TotalEmails = 0; TotalAttachments = 0; TotalSizeBytes = (Get-Item $MboxPath).Length; FirstDate = $null; LastDate = $null; UniqueSenders = @{} }
    $currentEmail = $false
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^From ') { $stats.TotalEmails++; $currentEmail = $true }
        if ($currentEmail) {
            if ($lines[$i] -match '^From: (.+?) <') { $sender = $matches[1]; $stats.UniqueSenders[$sender] = $true }
            if ($lines[$i] -match '^Date: (.+)') { try { $date = [datetime]::Parse($matches[1]); if (-not $stats.FirstDate -or $date -lt $stats.FirstDate) { $stats.FirstDate = $date }; if (-not $stats.LastDate -or $date -gt $stats.LastDate) { $stats.LastDate = $date } } catch { } }
            if ($lines[$i] -match '^Content-Disposition: attachment') { $stats.TotalAttachments++ }
        }
        if ($lines[$i] -match '^--.+_--$') { $currentEmail = $false }
    }
    return [PSCustomObject]@{ TotalEmails = $stats.TotalEmails; TotalAttachments = $stats.TotalAttachments; FileSizeBytes = $stats.TotalSizeBytes; FileSizeMB = [math]::Round($stats.TotalSizeBytes / 1MB, 2); DateRange = "$($stats.FirstDate) to $($stats.LastDate)"; UniqueSenders = $stats.UniqueSenders.Count }
}

Export-ModuleMember -Function Get-Base64, Escape-FromLines, Fold-Header, Sanitise-Filename, New-Boundary, Get-MimeType, Get-StoreFilePath, Invoke-WithRetry, Find-PstStore, Decode-Rfc2047, Decode-QuotedPrintable, Test-SafePath, Get-SafeFileName, Test-SafeFileExtension, Test-EmailMatchesFilters, Get-Attachments, Get-ExportStatistics, Invoke-AttachmentScanner, Test-MagicBytes, Test-OLEObject, Test-Archive, Test-DangerousPatterns, Test-SpecialFiles
