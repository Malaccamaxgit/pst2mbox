# PST2MBOX

PowerShell script that converts Outlook `.PST` files to standards-compliant `.MBOX` files using Outlook COM interop. Designed for email archival and migration workflows.

## Features

- **Full Fidelity**: Preserves HTML body, plain text, and all attachments as base64-encoded MIME parts
- **Flexible Security**: Export all attachments by default, or use `-RestrictAttachments` for comprehensive scanning
- **Extended Scanner**: Magic byte verification, OLE detection, archive inspection, pattern detection, special file blocking
- **Batch Processing**: Process multiple PST files from a directory in sequence
- **Filtering**: Filter by date range, folder patterns, and sender addresses
- **Deduplication**: Optional message-ID based deduplication
- **Split Output**: Split large exports into multiple MBOX files
- **Dry-Run Mode**: Preview operations with `-WhatIf`

## Requirements

- Microsoft Outlook (desktop) installed
- PowerShell 5.1+
- Windows 10/11

## Installation

Clone or download the repository:

```powershell
cd C:\path\to\pst2mbox
```

No formal installation required - run directly from the directory.

## Basic Usage

### Convert a single PST file

```powershell
.\pst2mbox.ps1 "C:\Archive\mailbox.pst"
# Output: C:\Archive\mailbox.mbox
```

### Specify output path

```powershell
.\pst2mbox.ps1 "C:\Archive\mailbox.pst" "D:\Export\output.mbox"
```

### Limit number of emails (for testing)

```powershell
.\pst2mbox.ps1 "C:\Archive\mailbox.pst" -MaxEmails 1000
```

## Advanced Usage

### Date Range Filtering

Export only emails from a specific period:

```powershell
.\pst2mbox.ps1 "C:\Archive\mailbox.pst" -DateFrom "2024-01-01" -DateTo "2024-12-31"
```

### Exclude Folders

Skip specific folders (e.g., Deleted Items, Junk):

```powershell
.\pst2mbox.ps1 "C:\Archive\mailbox.pst" -ExcludeFolders @("Deleted", "Junk", "Spam")
```

### Exclude Senders

Filter out emails from specific senders:

```powershell
.\pst2mbox.ps1 "C:\Archive\mailbox.pst" -ExcludeSenders @("noreply", "newsletter", "bot")
```

### Batch Mode - Process Multiple PSTs

Process all `.pst` files in a directory:

```powershell
.\pst2mbox.ps1 -BatchPath "C:\PSTs" -MaxEmails 5000
# Creates: C:\PSTs\file1.mbox, C:\PSTs\file2.mbox, etc.
```

### Split Large Exports

Split output into multiple files (e.g., 500MB each):

```powershell
.\pst2mbox.ps1 "C:\Archive\mailbox.pst" -SplitSizeMB 500
# Creates: mailbox.mbox, mailbox_1.mbox, mailbox_2.mbox, etc.
```

### Deduplication

Remove duplicate emails by Message-ID:

```powershell
.\pst2mbox.ps1 "C:\Archive\mailbox.pst" -DeDuplicate
```

### Dry-Run Mode

Preview operations without creating output:

```powershell
.\pst2mbox.ps1 "C:\Archive\mailbox.pst" -WhatIf
```

### Restrict Attachments

By default, ALL attachments are exported. Use `-RestrictAttachments` to enable security scanning and block dangerous file types:

```powershell
.\pst2mbox.ps1 "C:\Archive\mailbox.pst" -RestrictAttachments
```

This blocks:
- Dangerous extensions (`.exe`, `.bat`, `.cmd`, `.ps1`, `.vbs`, `.msg`, `.hta`, `.scf`, etc.)
- File type mismatches (e.g., `.pdf` that's actually an executable)
- OLE documents with embedded executables
- VBA macros in Office documents
- Archive bombs (highly compressed files)
- Password-protected archives
- Files with suspicious patterns (PowerShell, VBScript, ActiveX)

With size limits:

```powershell
.\pst2mbox.ps1 "C:\Archive\mailbox.pst" -RestrictAttachments -MaxAttachmentSizeMB 25
```

### Security: Restrict Output Directory

Limit output file creation to a specific directory:

```powershell
.\pst2mbox.ps1 "C:\Archive\mailbox.pst" -AllowedBasePath "C:\Export"
```

## Command-Line Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `-PstPath` | String (Required) | Path to the input `.PST` file |
| `-MboxPath` | String | Path to the output `.MBOX` file (default: same as PST with `.mbox` extension) |
| `-MaxEmails` | Int32 | Maximum number of emails to export (for testing) |
| `-FailedLogPath` | String | Path to log failed email exports |
| `-DateFrom` | DateTime | Export emails received on or after this date |
| `-DateTo` | DateTime | Export emails received on or before this date |
| `-ExcludeFolders` | String[] | Folder name patterns to exclude (supports wildcards) |
| `-ExcludeSenders` | String[] | Sender email/name patterns to exclude |
| `-SplitSizeMB` | Int32 | Split output into files of this size (in MB) |
| `-WhatIf` | Switch | Dry-run mode - preview without creating output |
| `-DeDuplicate` | Switch | Remove duplicate emails by Message-ID |
| `-PreserveHeaders` | String[] | Additional MAPI headers to preserve |
| `-RestrictAttachments` | Switch | Block dangerous attachments and enable security scanning |
| `-MaxAttachmentSizeMB` | Int32 | Maximum attachment size in MB (default: 50) |
| `-AllowedBasePath` | String | Restrict output to this directory (security) |
| `-BatchPath` | String | Process all `.pst` files in this directory |

## Module Functions

The `PST2MBOX.psm1` module exports 23 functions:

### Core Functions
- `Get-Base64` - Encode bytes to base64
- `Escape-FromLines` - Escape "From " lines for MBOX format
- `Fold-Header` - Fold long headers per RFC 5322
- `Sanitise-Filename` - Remove invalid filename characters
- `New-Boundary` - Generate MIME boundary strings
- `Get-MimeType` - Resolve MIME type from file extension

### Outlook Functions
- `Get-StoreFilePath` - Get PST store file path
- `Find-PstStore` - Locate PST store in Outlook
- `Get-Attachments` - Extract attachments (all by default, restricted with `-RestrictAttachments`)

### Decoding Functions
- `Decode-Rfc2047` - Decode RFC 2047 encoded headers
- `Decode-QuotedPrintable` - Decode quoted-printable encoding

### Security Functions
- `Test-SafePath` - Validate path traversal safety
- `Get-SafeFileName` - Sanitize filenames
- `Test-SafeFileExtension` - Validate file extension allowlist
- `Invoke-AttachmentScanner` - Orchestrates all security checks (magic bytes, OLE, archive, patterns, special files)
- `Test-MagicBytes` - Verify file signature matches extension
- `Test-OLEObject` - Detect OLE documents with embedded executables/macros
- `Test-Archive` - Inspect archives for bombs, passwords, suspicious contents
- `Test-DangerousPatterns` - Detect script injection, encoded commands, PE headers
- `Test-SpecialFiles` - Block HTA, SCF, DLL hijacking risks

**Note**: When `-RestrictAttachments` is used, `Invoke-AttachmentScanner` runs **all** checks automatically on each attachment.

### Filter Functions
- `Test-EmailMatchesFilters` - Check if email matches filters
- `Get-ExportStatistics` - Analyze exported MBOX files
- `Invoke-WithRetry` - Retry operations on failure

## Running Tests

Requires Pester 3.x:

```powershell
Invoke-Pester -Path .\pst2mbox.tests.ps1
```

All 36 tests should pass:
- 11 Helper function tests
- 7 Sanitizer function tests
- 10 Security function tests
- 6 Filter function tests
- 2 Module tests

## Security Features

### Default Behavior: Export All Attachments
By default, **all attachments are exported** without restrictions. This is useful for forensic archival where complete preservation is required. Use `-RestrictAttachments` to enable security filtering.

### Restrict Attachments Mode (`-RestrictAttachments`)
When specified, **all** security checks run automatically on every attachment:

**Step 1 - Blocked Extensions**: Immediately blocks `.exe`, `.bat`, `.cmd`, `.com`, `.pif`, `.scr`, `.vbs`, `.vbe`, `.js`, `.jse`, `.wsf`, `.wsh`, `.ps1`, `.msc`, `.msi`, `.msp`, `.hta`, `.scf`, `.lnk`, `.inf`, `.reg`, `.sys`, `.dll`, `.cpl`, `.drv`, `.msg`, `.emf`, `.wmf`

**Step 2 - Extended Scanner** (runs on all remaining attachments):

| Check | What It Detects |
|-------|-----------------|
| **Magic Byte Verification** | Verifies file signature matches extension (detects renamed executables) |
| **OLE Object Detection** | Identifies compound documents with embedded executables or VBA macros |
| **Archive Inspection** | Detects archive bombs (>1000:1 compression), password-protected archives, suspicious contents |
| **Dangerous Patterns** | Detects VBScript, PowerShell, ActiveX patterns, base64-encoded executables, PE headers in documents |
| **Special File Detection** | Blocks HTA, SCF, DLL files that can execute code or hijack loading |

### Path Traversal Prevention
Validates all paths stay within allowed base directory:
```powershell
Test-SafePath -Path "output.mbox" -BasePath "C:\Export"
```

### Filename Sanitization
Removes dangerous characters from attachment filenames:
- Path separators (`\`, `/`) replaced with `_`
- Null bytes removed
- Leading dots removed (prevents hidden file overwrites)
- Long filenames truncated to 255 characters

### Attachment Size Limits
Default 50MB limit prevents DoS attacks via oversized attachments.

### Command Injection Prevention
Blocks shell metacharacters (`;`, `|`, `&`) in paths.

### Cloud Attachment Prevention
Blocks attachments stored by reference (potential SSRF attacks).

## Output Format

MBOX output follows standard format:
- `From ` separator lines (escaped as `>From ` if needed)
- RFC 5322 headers (From, To, Subject, Date, Message-ID, etc.)
- MIME multipart structure for attachments
- Base64-encoded attachment content
- Proper header folding for long lines

## Troubleshooting

### "Store not found" error
Ensure the PST file is not already mounted in Outlook. Close Outlook and retry.

### "Access denied" error
Run as the user whose Outlook profile is active. Do not run as Administrator.

### Slow performance
Large PSTs take time. Use `-MaxEmails` for testing. Monitor progress with `-Verbose`.

### Encoding issues
The script handles UTF-8, Latin-1, and other common encodings. Check source PST encoding if issues persist.

## License

MIT License - See LICENSE file for details.

## Version History

- **2.3.0** - Inverted security model (export all by default, restrict with `-RestrictAttachments`), extended attachment scanner (magic bytes, OLE detection, archive inspection, pattern detection)
- **2.2.0** - Security hardening (path traversal, filename sanitization, extension allowlisting)
- **2.1.0** - Batch processing mode, retry logic, RFC 2047 decoding
- **2.0.0** - Complete rewrite with module architecture
- **1.0.0** - Initial release
