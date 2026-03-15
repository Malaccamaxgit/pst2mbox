$content = Get-Content 'E:/Github/pst2mbox/pst2mbox.v2.ps1' -Raw -Encoding UTF8
$errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseInput($content, [ref]$null, [ref]$errors)
if ($errors.Count -gt 0) {
    Write-Host "Syntax errors found:"
    $errors | ForEach-Object { Write-Host "Line $($_.Extent.StartLineNumber): $($_.Message)" }
} else {
    Write-Host "Syntax OK"
}
