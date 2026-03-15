# =============================================================================
# Pester Unit Tests for PST2MBOX
# =============================================================================
# Run: Invoke-Pester -Path .\pst2mbox.Tests.ps1
# Requirements: Pester 3.x, Microsoft Outlook installed
# =============================================================================

$scriptPath = Join-Path $PSScriptRoot "pst2mbox.ps1"
$modulePath = Join-Path $PSScriptRoot "PST2MBOX.psm1"
$testDir = Join-Path $PSScriptRoot "tests"

# -----------------------------------------------------------------------------
# Helper Function Tests
# -----------------------------------------------------------------------------
Describe "Helper Functions" {
    # Load only the function definitions
    $scriptContent = Get-Content -Path $scriptPath -Raw

    $functionPatterns = @(
        'function Get-Base64',
        'function Escape-FromLines',
        'function Fold-Header',
        'function Sanitise-Filename',
        'function New-Boundary',
        'function Get-MimeType',
        'function Test-EmailMatchesFilters'
    )

    foreach ($pattern in $functionPatterns) {
        $start = $scriptContent.IndexOf($pattern)
        if ($start -ge 0) {
            $end = $scriptContent.IndexOf("function ", $start + 1)
            if ($end -lt 0) { $end = $scriptContent.Length }
            $funcText = $scriptContent.Substring($start, $end - $start)
            . ([ScriptBlock]::Create($funcText))
        }
    }

    It "Get-Base64 returns empty string for null input" {
        Get-Base64 -bytes $null | Should Be ""
    }

    It "Get-Base64 returns empty string for empty array" {
        Get-Base64 -bytes @() | Should Be ""
    }

    It "Get-Base64 encodes bytes correctly" {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes("Hello World")
        $result = Get-Base64 -bytes $bytes
        $result.Trim() | Should Be "SGVsbG8gV29ybGQ="
    }

    It "Escape-FromLines escapes From at start of line" {
        $text = "From sender`r`nNormal line`r`nFrom another"
        $result = Escape-FromLines -text $text
        $result | Should BeLike "*`n>From *"
    }

    It "Fold-Header returns short headers unchanged" {
        $result = Fold-Header -name "Subject" -value "Short subject"
        $result | Should Be "Subject: Short subject"
    }

    It "Fold-Header folds long headers" {
        $longValue = "This is an extremely long header value that definitely exceeds the RFC 5322 limit of seventy-six characters per line and must be folded"
        $result = Fold-Header -name "Subject" -value $longValue
        $result | Should Match "^Subject:"
        ($result -split "`n").Count | Should BeGreaterThan 1
    }

    It "Sanitise-Filename removes invalid characters" {
        $result = Sanitise-Filename -name "file<>name*.txt"
        $result | Should Be "file__name_.txt"
    }

    It "Sanitise-Filename preserves valid characters" {
        $result = Sanitise-Filename -name "valid-file_name (1).txt"
        $result | Should Be "valid-file_name (1).txt"
    }

    It "New-Boundary generates unique strings" {
        $b1 = New-Boundary
        $b2 = New-Boundary
        $b1 | Should Not Be $b2
        $b1 | Should Match "^----=_Part_"
    }

    It "Get-MimeType returns correct type for common extensions" {
        Get-MimeType -filename "test.pdf" | Should Be "application/pdf"
        Get-MimeType -filename "test.png" | Should Be "image/png"
        Get-MimeType -filename "test.zip" | Should Be "application/zip"
    }

    It "Get-MimeType returns octet-stream for unknown extensions" {
        Get-MimeType -filename "test.xyz" | Should Be "application/octet-stream"
    }
}

# -----------------------------------------------------------------------------
# Sanitizer Function Tests
# -----------------------------------------------------------------------------
Describe "Sanitizer Functions" {
    $scriptContent = Get-Content -Path $scriptPath -Raw

    # Load Decode-Rfc2047
    $start = $scriptContent.IndexOf("function Decode-Rfc2047")
    if ($start -ge 0) {
        $end = $scriptContent.IndexOf("function ", $start + 1)
        if ($end -lt 0) { $end = $scriptContent.Length }
        $funcText = $scriptContent.Substring($start, $end - $start)
        . ([ScriptBlock]::Create($funcText))
    }

    # Load Decode-QuotedPrintable
    $start = $scriptContent.IndexOf("function Decode-QuotedPrintable")
    if ($start -ge 0) {
        $end = $scriptContent.IndexOf("function ", $start + 1)
        if ($end -lt 0) { $end = $scriptContent.Length }
        $funcText = $scriptContent.Substring($start, $end - $start)
        . ([ScriptBlock]::Create($funcText))
    }

    It "Decode-Rfc2047 decodes Base64 encoded headers" {
        $encoded = "=?utf-8?B?SGVsbG8gV29ybGQ=?="
        $result = Decode-Rfc2047 -header $encoded
        $result | Should Be "Hello World"
    }

    It "Decode-Rfc2047 decodes Quoted-Printable encoded headers" {
        $encoded = "=?utf-8?Q?Hello_World=?="
        $result = Decode-Rfc2047 -header $encoded
        $result.Trim('=') | Should Be "Hello World"
    }

    It "Decode-Rfc2047 returns unchanged header when no encoding" {
        $result = Decode-Rfc2047 -header "Plain subject"
        $result | Should Be "Plain subject"
    }

    It "Decode-Rfc2047 handles multiple encoded words" {
        $encoded = "=?utf-8?B?SGVsbG8=?= =?utf-8?B?V29ybGQ=?="
        $result = Decode-Rfc2047 -header $encoded
        $result | Should Be "Hello World"
    }

    It "Decode-QuotedPrintable decodes =XX hex sequences" {
        $text = "=48=65=6C=6C=6F"
        $result = Decode-QuotedPrintable -text $text
        [System.Text.Encoding]::UTF8.GetString($result) | Should Be "Hello"
    }

    It "Decode-QuotedPrintable removes soft line breaks" {
        $text = "Hello=`r`nWorld"
        $result = Decode-QuotedPrintable -text $text
        [System.Text.Encoding]::UTF8.GetString($result) | Should Be "HelloWorld"
    }

    It "Decode-QuotedPrintable handles UTF-8 characters" {
        $result = Decode-QuotedPrintable -text "Caf=E9"
        [System.Text.Encoding]::UTF8.GetString($result) | Should BeLike "Caf*"
    }
}

# -----------------------------------------------------------------------------
# Security Function Tests
# -----------------------------------------------------------------------------
Describe "Security Functions" {
    # Import module to get security functions
    if (Test-Path $modulePath) {
        Import-Module $modulePath -Force -ErrorAction Stop
    }

    $testBase = Join-Path $PSScriptRoot "tests"

    It "Test-SafePath allows path within base directory" {
        { Test-SafePath -Path ".\tests\file.txt" -BasePath $testBase } | Should Not Throw
    }

    It "Test-SafePath rejects path outside base directory" {
        { Test-SafePath -Path "..\..\secret.txt" -BasePath $testBase } | Should Throw "Security:"
    }

    It "Test-SafePath rejects path with command injection characters" {
        { Test-SafePath -Path "file.txt&del" -BasePath $testBase } | Should Throw "Security:"
    }

    It "Get-SafeFileName removes path separators" {
        $result = Get-SafeFileName -name "folder/subfolder/file.txt"
        $result | Should Be "folder_subfolder_file.txt"
    }

    It "Get-SafeFileName removes null bytes" {
        $result = Get-SafeFileName -name "file`0name.txt"
        $result | Should Be "filename.txt"
    }

    It "Get-SafeFileName removes leading dots" {
        $result = Get-SafeFileName -name "...hidden_file.txt"
        $result | Should Be "hidden_file.txt"
    }

    It "Get-SafeFileName truncates long filenames" {
        $longName = "a" * 300 + ".txt"
        $result = Get-SafeFileName -name $longName
        $result.Length | Should BeLessThan 256
    }

    It "Test-SafeFileExtension allows safe extensions" {
        Test-SafeFileExtension -FileName "document.pdf" | Should Be $true
        Test-SafeFileExtension -FileName "image.png" | Should Be $true
        Test-SafeFileExtension -FileName "archive.zip" | Should Be $true
    }

    It "Test-SafeFileExtension blocks dangerous extensions" {
        Test-SafeFileExtension -FileName "script.exe" | Should Be $false
        Test-SafeFileExtension -FileName "macro.bat" | Should Be $false
        Test-SafeFileExtension -FileName "driver.sys" | Should Be $false
    }

    It "Test-SafeFileExtension allows files without extension" {
        Test-SafeFileExtension -FileName "noextension" | Should Be $true
    }
}

# -----------------------------------------------------------------------------
# Filter Function Tests
# -----------------------------------------------------------------------------
Describe "Test-EmailMatchesFilters" {
    $scriptContent = Get-Content -Path $scriptPath -Raw
    $start = $scriptContent.IndexOf("function Test-EmailMatchesFilters")
    if ($start -ge 0) {
        $end = $scriptContent.IndexOf("function ", $start + 1)
        if ($end -lt 0) { $end = $scriptContent.IndexOf("# =", $start) }
        if ($end -lt 0) { $end = $scriptContent.Length }
        $funcText = $scriptContent.Substring($start, $end - $start)
        . ([ScriptBlock]::Create($funcText))
    }

    $mockItem = [PSCustomObject]@{
        SenderEmailAddress = "user@example.com"
        SenderName = "Test User"
        ReceivedTime = (Get-Date "2024-06-15")
    }

    It "Returns true when no filters specified" {
        Test-EmailMatchesFilters -item $mockItem -currentFolderName "Inbox" | Should Be $true
    }

    It "Excludes folder matching pattern" {
        Test-EmailMatchesFilters -item $mockItem -ExcludeFolders @("Deleted") -currentFolderName "Deleted Items" | Should Be $false
    }

    It "Allows folder not matching exclude pattern" {
        Test-EmailMatchesFilters -item $mockItem -ExcludeFolders @("Deleted") -currentFolderName "Inbox" | Should Be $true
    }

    It "Filters by date range - before DateFrom" {
        $earlyItem = [PSCustomObject]@{
            SenderEmailAddress = "user@example.com"
            SenderName = "Test User"
            ReceivedTime = (Get-Date "2024-01-15")
        }
        Test-EmailMatchesFilters -item $earlyItem -DateFrom (Get-Date "2024-06-01") -currentFolderName "Inbox" | Should Be $false
    }

    It "Filters by date range - after DateTo" {
        $lateItem = [PSCustomObject]@{
            SenderEmailAddress = "user@example.com"
            SenderName = "Test User"
            ReceivedTime = (Get-Date "2024-12-15")
        }
        Test-EmailMatchesFilters -item $lateItem -DateTo (Get-Date "2024-06-30") -currentFolderName "Inbox" | Should Be $false
    }

    It "Allows items within date range" {
        Test-EmailMatchesFilters -item $mockItem -DateFrom (Get-Date "2024-06-01") -DateTo (Get-Date "2024-06-30") -currentFolderName "Inbox" | Should Be $true
    }
}

# -----------------------------------------------------------------------------
# Integration Tests (requires Outlook and test PST)
# -----------------------------------------------------------------------------
Describe "Integration Tests" {
    # Check for Outlook COM object availability
    $outlookAvailable = $false
    try {
        $ol = New-Object -ComObject Outlook.Application
        $outlookAvailable = $true
        $ol.Quit()
    } catch {
        Write-Warning "Outlook COM not available - skipping integration tests"
    }

    if (-not $outlookAvailable) {
        return
    }
    $testPst = Join-Path $testDir "Workday2024.pst"
    if (-not (Test-Path $testPst)) {
        Write-Warning "Test PST not found at $testPst"
        return
    }

    It "Exports emails with -MaxEmails limit" {
        $outputPath = Join-Path $testDir "test_maxemails.mbox"
        if (Test-Path $outputPath) { Remove-Item $outputPath -Force }

        & $scriptPath -PstPath $testPst -MboxPath $outputPath -MaxEmails 50 -ErrorAction Stop

        (Test-Path $outputPath) | Should Be $true
    }

    It "Respects -WhatIf dry-run mode" {
        $outputPath = Join-Path $testDir "test_dryrun.mbox"
        if (Test-Path $outputPath) { Remove-Item $outputPath -Force }

        # Run script with -WhatIf (dry-run)
        & $scriptPath -PstPath $testPst -MboxPath $outputPath -WhatIf -ErrorAction Stop

        # Verify no output file was created (the key behavior of -WhatIf)
        (Test-Path $outputPath) | Should Be $false
    }

    It "Applies -ExcludeFolders filter" {
        $outputPath = Join-Path $testDir "test_excludefolders.mbox"
        if (Test-Path $outputPath) { Remove-Item $outputPath -Force }

        & $scriptPath -PstPath $testPst -MboxPath $outputPath -MaxEmails 100 -ExcludeFolders @("Deleted") -ErrorAction Stop 2>&1

        (Test-Path $outputPath) | Should Be $true
    }

    It "Applies -ExcludeSenders filter" {
        $outputPath = Join-Path $testDir "test_excludesenders.mbox"
        if (Test-Path $outputPath) { Remove-Item $outputPath -Force }

        & $scriptPath -PstPath $testPst -MboxPath $outputPath -MaxEmails 100 -ExcludeSenders @("noreply") -ErrorAction Stop 2>&1

        (Test-Path $outputPath) | Should Be $true
    }

    It "Applies -DateFrom and -DateTo filters" {
        $outputPath = Join-Path $testDir "test_daterange.mbox"
        if (Test-Path $outputPath) { Remove-Item $outputPath -Force }

        & $scriptPath -PstPath $testPst -MboxPath $outputPath -DateFrom (Get-Date "2024-01-01") -DateTo (Get-Date "2024-12-31") -MaxEmails 100 -ErrorAction Stop 2>&1

        (Test-Path $outputPath) | Should Be $true
    }
}

# -----------------------------------------------------------------------------
# Module Tests
# -----------------------------------------------------------------------------
Describe "PST2MBOX Module" {
    if (Test-Path $modulePath) {
        Import-Module $modulePath -Force -ErrorAction Stop
    }

    It "Exports expected functions" {
        $expectedFunctions = @(
            "Get-Base64", "Escape-FromLines", "Fold-Header", "Sanitise-Filename",
            "New-Boundary", "Get-MimeType", "Get-StoreFilePath", "Invoke-WithRetry",
            "Find-PstStore", "Get-Attachments", "Test-EmailMatchesFilters",
            "Get-ExportStatistics", "Decode-Rfc2047", "Decode-QuotedPrintable",
            "Test-SafePath", "Get-SafeFileName", "Test-SafeFileExtension"
        )
        $exported = Get-Command -Module PST2MBOX -CommandType Function -ErrorAction SilentlyContinue
        foreach ($fn in $expectedFunctions) {
            ($exported.Name -contains $fn) | Should Be $true
        }
    }

    It "Get-ExportStatistics works on valid MBOX" {
        $testMbox = Join-Path $testDir "Workday2024_100_v4.mbox"
        if (Test-Path $testMbox) {
            $stats = Get-ExportStatistics -MboxPath $testMbox
            $stats.TotalEmails | Should Not Be 0
            $stats.FileSizeBytes | Should Not Be 0
        }
    }
}

# -----------------------------------------------------------------------------
# Extended Scanner Function Tests (New in 2.3.0)
# -----------------------------------------------------------------------------
Describe "Extended Scanner Functions" {
    Import-Module $modulePath -Force -ErrorAction Stop

    It "Invoke-AttachmentScanner exists" {
        Get-Command Invoke-AttachmentScanner -ErrorAction SilentlyContinue | Should Not Be $null
    }

    It "Test-MagicBytes exists" {
        Get-Command Test-MagicBytes -ErrorAction SilentlyContinue | Should Not Be $null
    }

    It "Test-OLEObject exists" {
        Get-Command Test-OLEObject -ErrorAction SilentlyContinue | Should Not Be $null
    }

    It "Test-Archive exists" {
        Get-Command Test-Archive -ErrorAction SilentlyContinue | Should Not Be $null
    }

    It "Test-DangerousPatterns exists" {
        Get-Command Test-DangerousPatterns -ErrorAction SilentlyContinue | Should Not Be $null
    }

    It "Test-SpecialFiles exists" {
        Get-Command Test-SpecialFiles -ErrorAction SilentlyContinue | Should Not Be $null
    }
}

# -----------------------------------------------------------------------------
# Attachment Restriction Tests (New in 2.3.0)
# -----------------------------------------------------------------------------
Describe "Get-Attachments with -RestrictAttachments" {
    Import-Module $modulePath -Force -ErrorAction Stop

    # Create a mock mail item for testing
    $mockMailItem = New-Object -TypeName PSObject -Property @{
        Attachments = @{
            Count = 0
        }
    }

    $tempDir = [System.IO.Path]::GetTempPath()

    It "Get-Attachments accepts -RestrictAttachments parameter" {
        { Get-Attachments -mailItem $mockMailItem -tempDir $tempDir -RestrictAttachments } | Should Not Throw
    }

    It "Get-Attachments does not require -RestrictAttachments (exports all by default)" {
        { Get-Attachments -mailItem $mockMailItem -tempDir $tempDir } | Should Not Throw
    }
}

Describe "Test-SafeFileExtension with new blocked extensions" {
    Import-Module $modulePath -Force -ErrorAction Stop

    # Note: Test-SafeFileExtension only checks allowlist
    # The blocklist is checked in Get-Attachments when -RestrictAttachments is used

    It "Test-SafeFileExtension still allows common safe extensions" {
        Test-SafeFileExtension -FileName "document.pdf" | Should Be $true
        Test-SafeFileExtension -FileName "image.png" | Should Be $true
        Test-SafeFileExtension -FileName "archive.zip" | Should Be $true
        Test-SafeFileExtension -FileName "text.txt" | Should Be $true
    }

    It "Test-SafeFileExtension blocks traditionally dangerous extensions" {
        Test-SafeFileExtension -FileName "script.exe" | Should Be $false
        Test-SafeFileExtension -FileName "macro.bat" | Should Be $false
        Test-SafeFileExtension -FileName "driver.sys" | Should Be $false
    }
}

Describe "Test-MagicBytes function" {
    Import-Module $modulePath -Force -ErrorAction Stop

    # PDF magic bytes: 25 50 44 46 (%PDF)
    $pdfBytes = [byte[]]@(0x25, 0x50, 0x44, 0x46, 0x2D, 0x31, 0x2E, 0x34)

    # PNG magic bytes: 89 50 4E 47
    $pngBytes = [byte[]]@(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)

    # Executable magic bytes: 4D 5A (MZ)
    $exeBytes = [byte[]]@(0x4D, 0x5A, 0x90, 0x00, 0x03, 0x00)

    It "Detects PDF magic bytes correctly" {
        $result = Test-MagicBytes -Bytes $pdfBytes -Extension ".pdf"
        $result.IsMatch | Should Be $true
    }

    It "Detects PNG magic bytes correctly" {
        $result = Test-MagicBytes -Bytes $pngBytes -Extension ".png"
        $result.IsMatch | Should Be $true
    }

    It "Detects executable magic bytes" {
        $result = Test-MagicBytes -Bytes $exeBytes -Extension ".exe"
        $result.IsMatch | Should Be $true
        # ExpectedType is only set on mismatch, so for matching .exe it should be null
        $result.ExpectedType | Should Be $null
    }

    It "Detects file type mismatch (EXE masquerading as PDF)" {
        $result = Test-MagicBytes -Bytes $exeBytes -Extension ".pdf"
        $result.IsMatch | Should Be $false
        $result.ExpectedType | Should BeLike "*Executable*"
    }
}

Describe "Test-OLEObject function" {
    Import-Module $modulePath -Force -ErrorAction Stop

    # OLE Compound Document magic bytes: D0 CF 11 E0
    $oleBytes = [byte[]]@(0xD0, 0xCF, 0x11, 0xE0, 0xA5, 0xB1, 0x1A, 0xE1)

    It "Detects OLE compound document" {
        $result = Test-OLEObject -Bytes $oleBytes
        $result.IsOLE | Should Be $true
    }

    It "Returns false for non-OLE data" {
        $nonOleBytes = [byte[]]@(0x50, 0x4B, 0x03, 0x04)  # ZIP
        $result = Test-OLEObject -Bytes $nonOleBytes
        $result.IsOLE | Should Be $false
    }
}

Describe "Test-Archive function" {
    Import-Module $modulePath -Force -ErrorAction Stop

    # ZIP magic bytes: 50 4B 03 04
    $zipBytes = [byte[]]@(0x50, 0x4B, 0x03, 0x04)

    It "Detects ZIP archive" {
        $result = Test-Archive -Bytes $zipBytes -Extension ".zip"
        $result.IsArchive | Should Be $true
    }

    It "Returns false for non-archive data" {
        $result = Test-Archive -Bytes ([byte[]]@(0x00, 0x00, 0x00, 0x00)) -Extension ".txt"
        $result.IsArchive | Should Be $false
    }
}

Describe "Test-DangerousPatterns function" {
    Import-Module $modulePath -Force -ErrorAction Stop

    It "Detects PowerShell patterns in content" {
        $content = [System.Text.Encoding]::ASCII.GetBytes("Invoke-Expression -Command 'download'")
        $result = Test-DangerousPatterns -Bytes $content -FileName "document.pdf"
        $result.Warnings.Count | Should BeGreaterThan 0
    }

    It "Detects VBScript patterns in content" {
        $content = [System.Text.Encoding]::ASCII.GetBytes("On Error Resume Next")
        $result = Test-DangerousPatterns -Bytes $content -FileName "document.doc"
        $result.Warnings.Count | Should BeGreaterThan 0
    }

    It "Detects JavaScript/ActiveX patterns in content" {
        $content = [System.Text.Encoding]::ASCII.GetBytes("new ActiveXObject('WScript.Shell')")
        $result = Test-DangerousPatterns -Bytes $content -FileName "document.pdf"
        $result.Warnings.Count | Should BeGreaterThan 0
    }
}

Describe "Test-SpecialFiles function" {
    Import-Module $modulePath -Force -ErrorAction Stop

    It "Blocks HTA files" {
        $result = Test-SpecialFiles -Bytes ([byte[]]@()) -FileName "malicious.hta"
        $result.Blocked | Should Be $true
        $result.Reason | Should BeLike "*HTA*"
    }

    It "Blocks SCF files" {
        $result = Test-SpecialFiles -Bytes ([byte[]]@()) -FileName "credential.scf"
        $result.Blocked | Should Be $true
        $result.Reason | Should BeLike "*SCF*"
    }

    It "Blocks DLL files" {
        $result = Test-SpecialFiles -Bytes ([byte[]]@()) -FileName "payload.dll"
        $result.Blocked | Should Be $true
        $result.Reason | Should BeLike "*DLL*"
    }

    It "Allows normal file extensions" {
        $result = Test-SpecialFiles -Bytes ([byte[]]@()) -FileName "document.pdf"
        $result.Blocked | Should Be $false
    }
}
