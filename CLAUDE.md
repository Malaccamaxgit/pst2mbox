# pst2mbox Project Context

## Repository
- **URL**: https://github.com/Malaccamaxgit/pst2mbox
- **Branch**: main
- **Author**: Benjamin Alloul <benjamin.alloul+pst2mbox@gmail.com>

## Project Overview
PowerShell script that converts Outlook `.PST` files to standards-compliant `.MBOX` files using Outlook COM interop. Designed for email archival and migration workflows.

## Version
**2.3.0** - Latest

## Key Architecture Decisions

### Security Model (v2.3.0)
- **Default**: Export ALL attachments without restrictions (forensic archival)
- **Opt-in restriction**: Use `-RestrictAttachments` flag to enable security scanning

### Extended Scanner (when -RestrictAttachments is used)
All checks run automatically on every attachment:
1. **Blocked Extensions**: `.exe`, `.bat`, `.cmd`, `.ps1`, `.vbs`, `.hta`, `.scf`, `.dll`, `.msg`, etc.
2. **Magic Byte Verification**: Detects file type mismatches
3. **OLE Object Detection**: Identifies embedded executables/VBA macros
4. **Archive Inspection**: Detects archive bombs, password-protected archives
5. **Dangerous Patterns**: Detects script injection, encoded commands, PE headers
6. **Special File Detection**: Blocks HTA, SCF, library files

## Key Files
- `pst2mbox.ps1` - Main script (entry point)
- `PST2MBOX.psm1` - Module with exported functions
- `pst2mbox.tests.ps1` - Pester test suite (66 tests)
- `README.md` - User documentation

## Testing
```powershell
Invoke-Pester -Path .\pst2mbox.tests.ps1
# 66 tests should pass
```

## Usage Examples
```powershell
# Export all attachments (default)
.\pst2mbox.ps1 "C:\Archive\mailbox.pst"

# Restrict dangerous attachments
.\pst2mbox.ps1 "C:\Archive\mailbox.pst" -RestrictAttachments

# Dry-run mode
.\pst2mbox.ps1 "C:\Archive\mailbox.pst" -WhatIf
```

## Requirements
- Microsoft Outlook (desktop) installed
- PowerShell 5.1+
- Windows 10/11
