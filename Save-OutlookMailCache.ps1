#Requires -Version 5.1
<#
.SYNOPSIS
    Preserves every Outlook data file (OST/PST) that may still hold a mailbox's messages, so they can be recovered.

.DESCRIPTION
    Run this on the computer where classic Outlook is connected to the mailbox. It never changes, repairs or deletes
    the original files. It:
      1. Stops Outlook from syncing: closes Outlook (asks first), because every sync and every new message can
         overwrite space where removed messages may still be recoverable.
      2. Finds the OST/PST files for the account (current and older ones, for example "name(2).ost"), in every
         user profile it can read.
      3. With administrator rights, also finds older copies of those files in Windows restore points / shadow
         copies ("Previous Versions"). An older copy can still contain the whole inbox.
      4. Copies everything to the destination, checks each copy with SHA-256, marks the copies read-only and
         writes Report.txt and Report.csv.

    The copies can then be opened with an OST-to-PST converter on another computer. Outlook itself cannot open
    an OST file directly.

.PARAMETER Account
    The email address whose data files to save. Files whose names contain it are copied. If none match, every
    OST and PST found is copied, because data file names do not always contain the address.

.PARAMETER Destination
    Folder for the copies, ideally on a USB drive. Default: a new folder next to this script.

.PARAMETER ListOnly
    Only lists what would be copied. Copies nothing and does not need Outlook to be closed.

.PARAMETER IgnoreRunningOutlook
    Copies even while Outlook is running (files Outlook holds open may fail to copy). Not recommended.

.EXAMPLE
    .\Save-OutlookMailCache.ps1 -Destination E:\MailRescue
#>
[CmdletBinding()]
param(
    [string]$Account = 'db4gtzehner@yahoo.com',
    [string]$Destination,
    [switch]$ListOnly,
    [switch]$IgnoreRunningOutlook
)

$ErrorActionPreference = 'Stop'
$Version = '1.0.0'

function Write-Step([string]$Text) { Write-Host ''; Write-Host "== $Text" -ForegroundColor Cyan }
function Write-Ok([string]$Text) { Write-Host "   $Text" -ForegroundColor Green }
function Write-Warn([string]$Text) { Write-Host "   $Text" -ForegroundColor Yellow }
function Write-Info([string]$Text) { Write-Host "   $Text" }
function Read-YesNo([string]$Question) {
    while ($true) {
        $a = Read-Host "$Question [Y/N]"
        if ($a -match '^(y|yes)$') { return $true }
        if ($a -match '^(n|no)$') { return $false }
    }
}
function Test-IsAdmin {
    $p = New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())
    $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
function Format-Size([int64]$Bytes) {
    if ($Bytes -ge 1GB) { '{0:N2} GB' -f ($Bytes / 1GB) } else { '{0:N1} MB' -f ($Bytes / 1MB) }
}
function Get-Sha256([string]$Path) {
    $sha = [Security.Cryptography.SHA256]::Create()
    $fs = New-Object IO.FileStream($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try { ([BitConverter]::ToString($sha.ComputeHash($fs))).Replace('-', '') } finally { $fs.Dispose(); $sha.Dispose() }
}

Write-Host "Outlook Mail Cache Rescue $Version" -ForegroundColor White
Write-Host "Account: $Account    Computer: $env:COMPUTERNAME    User: $env:USERNAME"
$isAdmin = Test-IsAdmin
if (-not $isAdmin) { Write-Warn 'Not running as administrator: other users'' profiles and restore-point copies will be skipped.' }

# ---------- 1. stop Outlook from syncing ----------
Write-Step 'Step 1: Outlook must not be syncing'
$running = @(Get-Process OUTLOOK -ErrorAction SilentlyContinue)
if ($running.Count -gt 0 -and -not $ListOnly -and -not $IgnoreRunningOutlook) {
    Write-Warn 'Outlook is running. While it runs it keeps syncing and writing new mail into the data file.'
    if (-not (Read-YesNo 'Close Outlook now? (it is asked to close normally; nothing is deleted)')) {
        Write-Warn 'Stopped. Close Outlook (or choose Send / Receive > Work Offline and then close it) and run this again.'
        exit 2
    }
    try {
        $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        $ol.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ol)
    } catch {
        Write-Warn "Outlook could not be asked to close automatically ($($_.Exception.Message))."
    }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while (@(Get-Process OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0 -and $sw.Elapsed.TotalSeconds -lt 90) { Start-Sleep -Seconds 2 }
    while (@(Get-Process OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0) {
        Write-Warn 'Outlook is still open. Close it yourself (File > Exit), wait until it disappears from Task Manager,'
        [void](Read-Host '   then press Enter to continue')
    }
    Write-Ok 'Outlook is closed. Do not open it again until this finishes.'
} elseif ($running.Count -gt 0) {
    Write-Warn 'Outlook is running (allowed by -ListOnly / -IgnoreRunningOutlook).'
} else {
    Write-Ok 'Outlook is not running.'
}

# ---------- 2. find data files ----------
Write-Step 'Step 2: Finding Outlook data files'
$profiles = if ($isAdmin) { @(Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName }) } else { @($env:USERPROFILE) }
$relDirs = @('AppData\Local\Microsoft\Outlook', 'Documents\Outlook Files', 'OneDrive\Documents\Outlook Files')
$found = New-Object System.Collections.ArrayList
foreach ($p in $profiles) {
    foreach ($rel in $relDirs) {
        $dir = Join-Path $p $rel
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.ost', '.pst' })) {
            [void]$found.Add([pscustomobject]@{ Source = $f.FullName; Origin = 'Current'; Size = $f.Length; Modified = $f.LastWriteTime; Name = $f.Name })
        }
    }
}

# ---------- 3. restore-point copies ----------
$shadowNote = ''
if ($isAdmin) {
    try {
        $shadows = @(Get-CimInstance Win32_ShadowCopy -ErrorAction Stop | Sort-Object InstallDate -Descending)
        $volumes = @{}
        foreach ($v in @(Get-CimInstance Win32_Volume -ErrorAction SilentlyContinue)) { if ($v.DriveLetter) { $volumes[$v.DeviceID] = $v.DriveLetter } }
        foreach ($s in $shadows) {
            $letter = $volumes[$s.VolumeName]
            if (-not $letter) { continue }
            $when = $s.InstallDate
            foreach ($p in $profiles) {
                if (-not $p.StartsWith($letter, [StringComparison]::OrdinalIgnoreCase)) { continue }
                foreach ($rel in $relDirs) {
                    $dir = $s.DeviceObject + '\' + (Join-Path $p $rel).Substring(3)
                    if (-not [IO.Directory]::Exists($dir)) { continue }
                    foreach ($path in [IO.Directory]::GetFiles($dir)) {
                        $ext = [IO.Path]::GetExtension($path).ToLowerInvariant()
                        if ($ext -ne '.ost' -and $ext -ne '.pst') { continue }
                        $fi = New-Object IO.FileInfo($path)
                        [void]$found.Add([pscustomobject]@{ Source = $path; Origin = "Restore point $($when.ToString('yyyy-MM-dd HH:mm'))"; Size = $fi.Length; Modified = $fi.LastWriteTime; Name = $fi.Name })
                    }
                }
            }
        }
        $shadowNote = "$($shadows.Count) restore point(s) / shadow copies checked."
    } catch { $shadowNote = "Restore points could not be read: $($_.Exception.Message)" }
} else {
    $shadowNote = 'Restore points skipped (needs administrator).'
}
Write-Info $shadowNote

# Keep the account's files; if none are named for it, keep everything found.
$acct = $Account.ToLowerInvariant()
$matching = @($found | Where-Object { $_.Name.ToLowerInvariant().Contains($acct) })
if ($matching.Count -gt 0) { $selected = $matching; Write-Ok "$($matching.Count) file(s) named for $Account." }
else { $selected = @($found); Write-Warn "No file is named for $Account; copying all $($found.Count) Outlook data file(s) found." }
# A restore-point copy identical to one already listed (same name, size, time) adds nothing.
$seen = @{}
$selected = @($selected | Where-Object { $k = "$($_.Name)|$($_.Size)|$($_.Modified.Ticks)"; if ($seen.ContainsKey($k)) { $false } else { $seen[$k] = $true; $true } })
if ($selected.Count -eq 0) { Write-Warn 'No Outlook data files were found. Is this the right computer and Windows user?'; exit 3 }
$total = [int64]0
foreach ($f in $selected) { $total += $f.Size; Write-Info ("{0,-45} {1,10}  {2:yyyy-MM-dd HH:mm}  {3}" -f $f.Name, (Format-Size $f.Size), $f.Modified, $f.Origin) }
Write-Info "Total: $(Format-Size $total)"
if ($ListOnly) { Write-Host ''; Write-Ok 'List only: nothing was copied.'; exit 0 }

# ---------- 4. copy and verify ----------
Write-Step 'Step 3: Copying and verifying'
if (-not $Destination) { $Destination = Join-Path $PSScriptRoot ("MailRescue-$env:COMPUTERNAME-" + (Get-Date -Format 'yyyyMMdd-HHmm')) }
[void](New-Item -ItemType Directory -Path $Destination -Force)
$root = [IO.Path]::GetPathRoot((Resolve-Path -LiteralPath $Destination).Path)
$drive = New-Object IO.DriveInfo($root)
if ($drive.AvailableFreeSpace -lt [int64]($total * 1.05)) { throw "Not enough free space on $root ($(Format-Size $drive.AvailableFreeSpace) free, $(Format-Size $total) needed)." }
$rows = New-Object System.Collections.ArrayList
$i = 0
foreach ($f in $selected) {
    $i++
    $tag = if ($f.Origin -eq 'Current') { 'current' } else { 'restorepoint-' + ($f.Origin -replace '[^0-9]', '') }
    $target = Join-Path $Destination ("{0}__{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($f.Name), $tag, [IO.Path]::GetExtension($f.Name))
    Write-Info "[$i/$($selected.Count)] $($f.Name) ($($f.Origin), $(Format-Size $f.Size))"
    $row = [pscustomobject]@{ File = $f.Name; Origin = $f.Origin; Source = $f.Source; Copy = $target; Size = $f.Size; Modified = $f.Modified; Sha256 = ''; Verified = $false; Error = '' }
    try {
        $in = New-Object IO.FileStream($f.Source, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        try { $out = [IO.File]::Create($target); try { $in.CopyTo($out, 4MB) } finally { $out.Dispose() } } finally { $in.Dispose() }
        $h1 = Get-Sha256 $f.Source; $h2 = Get-Sha256 $target
        $row.Sha256 = $h2; $row.Verified = ($h1 -eq $h2)
        (Get-Item -LiteralPath $target).IsReadOnly = $true
        if ($row.Verified) { Write-Ok 'copied and verified' } else { Write-Warn 'copied, but the copy does not match the original (was it still in use?)' }
    } catch { $row.Error = $_.Exception.Message; Write-Warn "FAILED: $($row.Error)" }
    [void]$rows.Add($row)
}

# ---------- report ----------
$rows | Export-Csv -LiteralPath (Join-Path $Destination 'Report.csv') -NoTypeInformation -Encoding UTF8
$ok = @($rows | Where-Object Verified).Count
$report = @(
    "Outlook Mail Cache Rescue $Version"
    "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm')   Computer: $env:COMPUTERNAME   User: $env:USERNAME   Administrator: $isAdmin"
    "Account: $Account"
    $shadowNote
    "Copied and verified: $ok of $($rows.Count)"
    ''
) + @($rows | ForEach-Object { "{0} | {1} | {2} | {3:yyyy-MM-dd HH:mm} | verified={4} {5}`r`n    from {6}" -f $_.File, $_.Origin, (Format-Size $_.Size), $_.Modified, $_.Verified, $_.Error, $_.Source }) + @(
    ''
    'Next steps:'
    ' - Keep these copies safe and unchanged. Work only on further copies of them.'
    ' - Open them with an OST-to-PST converter on another computer to see what can be recovered.'
    ' - Ask Yahoo to restore lost mail (Yahoo Help: restore lost or deleted emails) as soon as possible.'
    ' - Check the Yahoo account for unknown sign-ins, filters, forwarding and other devices before reconnecting.'
)
[IO.File]::WriteAllLines((Join-Path $Destination 'Report.txt'), $report)
Write-Step 'Done'
Write-Ok "$ok of $($rows.Count) file(s) copied and verified to: $Destination"
Write-Info 'See Report.txt there. Outlook can be opened again now, but every sync can overwrite recoverable data in the original file.'
if (@($rows | Where-Object { -not $_.Verified }).Count -gt 0) { exit 1 }
exit 0
