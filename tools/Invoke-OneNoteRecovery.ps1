<#
.SYNOPSIS
    Retry and merge failed OneNote PDF/HTML exports.

.DESCRIPTION
    This is a recovery companion for older OneNote exporter workflows that
    create a OneNote-Export-failed-pages.log file with FAILED, FORMAT and ID
    fields. It retries failed OneNote Publish calls, uses short temporary paths
    for path-length failures, and merges recovered files into a predictable
    _recovered folder.

    The COM retry steps require Windows, OneNote desktop and Windows PowerShell
    5.1. Merge and copy-only repair steps can be re-run safely.

.EXAMPLE
    .\tools\Invoke-OneNoteRecovery.ps1 -Step All

.EXAMPLE
    .\tools\Invoke-OneNoteRecovery.ps1 -Step Merge -ExportRoot "C:\Exports\OneNote-Export"
#>

[CmdletBinding()]
param(
    [ValidateSet("All", "RetryV1", "RetryV2", "Merge", "FixV1Misclassified", "FixV1Flat")]
    [string]$Step = "All",

    [string]$FailedLog = "$env:USERPROFILE\Downloads\OneNote-Export-failed-pages.log",
    [string]$ExportRoot = "$env:USERPROFILE\Downloads\OneNote-Export",
    [string]$RecoveryRoot = "$env:USERPROFILE\Downloads\OneNote-Export-Recovered",
    [string]$ShortRoot = "$env:USERPROFILE\Downloads\ONR2",
    [string]$MergedRootName = "_recovered",

    [string]$RetryResultsCsv = "",
    [string]$V2ResultsCsv = "",
    [string]$RecoveredIndexCsv = "",
    [string]$V1FixedIndexCsv = "",

    [string]$V1MisclassifiedOutFolderName = "",
    [string]$V1FlatOutFolderName = "",

    [switch]$IncludeUnnamed,
    [switch]$Overwrite,
    [switch]$OverwriteExisting
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$FormatMap = @{
    "pfPDF"  = 3
    "pfHTML" = 7
}

if ([string]::IsNullOrWhiteSpace($RetryResultsCsv)) {
    $RetryResultsCsv = Join-Path $RecoveryRoot "retry-results.csv"
}
if ([string]::IsNullOrWhiteSpace($V2ResultsCsv)) {
    $V2ResultsCsv = Join-Path $ShortRoot "retry-v2-results.csv"
}
if ([string]::IsNullOrWhiteSpace($RecoveredIndexCsv)) {
    $RecoveredIndexCsv = Join-Path (Join-Path $ExportRoot $MergedRootName) "_recovered-index.csv"
}
if ([string]::IsNullOrWhiteSpace($V1MisclassifiedOutFolderName)) {
    $V1MisclassifiedOutFolderName = Join-Path $MergedRootName "v1-from-recovered-folder"
}
if ([string]::IsNullOrWhiteSpace($V1FlatOutFolderName)) {
    $V1FlatOutFolderName = Join-Path $MergedRootName "v1-flat-failed"
}
if ([string]::IsNullOrWhiteSpace($V1FixedIndexCsv)) {
    $V1FixedIndexCsv = Join-Path (Join-Path $ExportRoot $V1MisclassifiedOutFolderName) "_v1-fixed-index.csv"
}

function Write-StepHeader {
    param([Parameter(Mandatory)][string]$Title)

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor DarkCyan
    Write-Host $Title -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor DarkCyan
}

function Ensure-Dir {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $Path | Out-Null
    }
}

function Assert-WindowsForOneNoteCom {
    if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
        throw "OneNote COM automation requires Windows with OneNote desktop installed."
    }
}

function Get-Sha1Short {
    param(
        [AllowEmptyString()][string]$Text,
        [ValidateRange(4, 40)][int]$Length = 12
    )

    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $hashBytes = $sha1.ComputeHash($bytes)
        $hex = -join ($hashBytes | ForEach-Object { $_.ToString("x2") })
        return $hex.Substring(0, [Math]::Min($Length, $hex.Length))
    }
    finally {
        $sha1.Dispose()
    }
}

function Sanitize-Component {
    param([AllowEmptyString()][string]$Name)

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($char in $invalid) {
        $Name = $Name.Replace($char, "-")
    }

    $Name = $Name -replace "\s+", " "
    $Name = $Name.Trim().TrimEnd(".")

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return "_empty"
    }

    return $Name
}

function Shorten-Component {
    param(
        [AllowEmptyString()][string]$Name,
        [ValidateRange(12, 180)][int]$MaxLength = 55
    )

    $clean = Sanitize-Component $Name
    if ($clean.Length -le $MaxLength) {
        return $clean
    }

    $hash = Get-Sha1Short $clean 8
    $prefixLength = $MaxLength - 9
    return ($clean.Substring(0, $prefixLength).TrimEnd("-", " ", ".") + "-" + $hash)
}

function Get-RelativePathFromRoot {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Root
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fullRoot = [System.IO.Path]::GetFullPath($Root).TrimEnd("\", "/")

    if ($fullPath.StartsWith($fullRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $fullPath.Substring($fullRoot.Length).TrimStart("\", "/")
    }

    return Split-Path $fullPath -Leaf
}

function Copy-DirectoryContent {
    param(
        [Parameter(Mandatory)][string]$SourceDir,
        [Parameter(Mandatory)][string]$DestDir
    )

    Ensure-Dir $DestDir

    Get-ChildItem -LiteralPath $SourceDir -Force | ForEach-Object {
        $dest = Join-Path $DestDir $_.Name
        if ($_.PSIsContainer) {
            Copy-Item -LiteralPath $_.FullName -Destination $dest -Recurse -Force
        }
        else {
            Copy-Item -LiteralPath $_.FullName -Destination $dest -Force
        }
    }
}

function Copy-TreeToDestination {
    param(
        [Parameter(Mandatory)][string]$SourceDir,
        [Parameter(Mandatory)][string]$DestinationDir
    )

    Ensure-Dir $DestinationDir
    Get-ChildItem -LiteralPath $SourceDir -Force | ForEach-Object {
        $dest = Join-Path $DestinationDir $_.Name
        if ($_.PSIsContainer) {
            Copy-Item -LiteralPath $_.FullName -Destination $dest -Recurse -Force
        }
        else {
            Copy-Item -LiteralPath $_.FullName -Destination $dest -Force
        }
    }
}

function Copy-FileAndHtmlResources {
    param(
        [Parameter(Mandatory)][string]$SourceFile,
        [Parameter(Mandatory)][string]$DestDir,
        [switch]$Overwrite
    )

    if (-not (Test-Path -LiteralPath $SourceFile -PathType Leaf)) {
        throw "Source file not found: $SourceFile"
    }

    Ensure-Dir $DestDir

    $fileName = [System.IO.Path]::GetFileName($SourceFile)
    $destFile = Join-Path $DestDir $fileName

    if ((Test-Path -LiteralPath $destFile) -and (-not $Overwrite)) {
        return [PSCustomObject]@{
            CopiedFile = $destFile
            Note       = "already_exists"
        }
    }

    Copy-Item -LiteralPath $SourceFile -Destination $destFile -Force:$Overwrite

    $ext = [System.IO.Path]::GetExtension($SourceFile)
    if ($ext -ieq ".htm" -or $ext -ieq ".html") {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($SourceFile)
        $sourceDir = Split-Path $SourceFile -Parent

        foreach ($resourceFolderName in @("$base-Dateien", "$base_files")) {
            $resourceDir = Join-Path $sourceDir $resourceFolderName
            if (-not (Test-Path -LiteralPath $resourceDir -PathType Container)) {
                continue
            }

            $destResourceDir = Join-Path $DestDir ([System.IO.Path]::GetFileName($resourceDir))
            if ((Test-Path -LiteralPath $destResourceDir) -and $Overwrite) {
                Remove-Item -LiteralPath $destResourceDir -Recurse -Force
            }
            if (-not (Test-Path -LiteralPath $destResourceDir)) {
                Copy-Item -LiteralPath $resourceDir -Destination $destResourceDir -Recurse -Force
            }
        }
    }

    return [PSCustomObject]@{
        CopiedFile = $destFile
        Note       = "copied"
    }
}

function Get-ShortRecoveryPath {
    param(
        [Parameter(Mandatory)][string]$OriginalPath,
        [Parameter(Mandatory)][string]$ExportRoot,
        [Parameter(Mandatory)][string]$RecoveryRoot
    )

    $relative = Get-RelativePathFromRoot -Path $OriginalPath -Root $ExportRoot
    $parts = $relative -split "[\\/]"
    $shortParts = @()

    for ($i = 0; $i -lt $parts.Count; $i++) {
        if ([string]::IsNullOrWhiteSpace($parts[$i])) {
            continue
        }

        if ($i -eq ($parts.Count - 1)) {
            $ext = [System.IO.Path]::GetExtension($parts[$i])
            $base = [System.IO.Path]::GetFileNameWithoutExtension($parts[$i])
            $shortParts += ((Shorten-Component $base 70) + $ext)
        }
        else {
            $shortParts += (Shorten-Component $parts[$i] 45)
        }
    }

    $result = $RecoveryRoot
    foreach ($part in $shortParts) {
        $result = Join-Path $result $part
    }

    return $result
}

function Parse-FailedLog {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Failed log not found: $Path"
    }

    $raw = Get-Content -LiteralPath $Path -Raw
    $blocks = $raw -split "(?m)^\s*----\s*$"
    $items = @()

    foreach ($block in $blocks) {
        $record = @{}
        foreach ($line in ($block -split "\r?\n")) {
            if ($line -match "^\s*(FAILED|FORMAT|ID)\s*:\s*(.*?)\s*$") {
                $record[$Matches[1].ToUpperInvariant()] = $Matches[2]
            }
        }

        if ($record.ContainsKey("FAILED") -and $record.ContainsKey("FORMAT") -and $record.ContainsKey("ID")) {
            $items += [PSCustomObject]@{
                Format = $record["FORMAT"]
                Path   = $record["FAILED"]
                Id     = $record["ID"]
            }
        }
    }

    return @($items)
}

function Get-FreeDriveLetter {
    $used = @(Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Name)
    foreach ($letter in @("O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")) {
        if ($used -notcontains $letter) {
            return $letter
        }
    }

    throw "No free drive letter found for subst."
}

function Start-SubstDrive {
    param([Parameter(Mandatory)][string]$TargetPath)

    Ensure-Dir $TargetPath
    $driveLetter = Get-FreeDriveLetter
    $drive = "${driveLetter}:"

    Write-Host "Creating short drive mapping: $drive -> $TargetPath" -ForegroundColor Cyan
    $output = cmd /c "subst $drive `"$TargetPath`"" 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "subst failed for $drive -> $TargetPath. Output: $output"
    }

    return $drive
}

function Stop-SubstDrive {
    param([AllowEmptyString()][string]$Drive)

    if ([string]::IsNullOrWhiteSpace($Drive)) {
        return
    }

    Write-Host "Removing short drive mapping $Drive" -ForegroundColor Cyan
    cmd /c "subst $Drive /d" | Out-Null
}

function New-OneNoteComApplication {
    Assert-WindowsForOneNoteCom
    return New-Object -ComObject OneNote.Application
}

function Release-ComObject {
    param([AllowNull()][object]$ComObject)

    if ($null -ne $ComObject) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ComObject) | Out-Null
    }
}

function Invoke-OneNotePublish {
    param(
        [Parameter(Mandatory)][object]$OneNote,
        [Parameter(Mandatory)][string]$PageId,
        [Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][int]$Format
    )

    $OneNote.Publish($PageId, $Destination, $Format, "")
    if (-not (Test-Path -LiteralPath $Destination -PathType Leaf)) {
        throw "OneNote reported no exception, but output was not created: $Destination"
    }
}

function Read-RecoveryRows {
    param(
        [Parameter(Mandatory)][string]$CsvPath,
        [Parameter(Mandatory)][string]$BatchName
    )

    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
        Write-Host "CSV missing, skipping: $CsvPath" -ForegroundColor Yellow
        return @()
    }

    $rows = Import-Csv -LiteralPath $CsvPath
    $rows | Where-Object {
        $_.Status -in @("OK", "ALREADY_EXISTS") -and
        -not [string]::IsNullOrWhiteSpace($_.SavedPath)
    } | ForEach-Object {
        [PSCustomObject]@{
            Batch        = $BatchName
            Status       = $_.Status
            Format       = $_.Format
            Mode         = if (Get-Member -InputObject $_ -Name "Mode") { $_.Mode } else { "" }
            OriginalPath = $_.OriginalPath
            SavedPath    = $_.SavedPath
        }
    }
}

function Export-Rows {
    param(
        [AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string]$Path
    )

    Ensure-Dir (Split-Path $Path -Parent)
    @($Rows) | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Invoke-RetryV1 {
    param(
        [string]$FailedLog,
        [string]$ExportRoot,
        [string]$RecoveryRoot,
        [string]$RetryResultsCsv,
        [switch]$IncludeUnnamed,
        [switch]$OverwriteExisting
    )

    Write-StepHeader "Step: RetryV1 - retry failed OneNote pages"
    Ensure-Dir $RecoveryRoot

    $failedItems = Parse-FailedLog $FailedLog
    if (-not $IncludeUnnamed) {
        $failedItems = $failedItems | Where-Object {
            ([System.IO.Path]::GetFileNameWithoutExtension($_.Path)) -notlike "Unbenannt*"
        }
    }

    $failedItems = @($failedItems | Where-Object { $FormatMap.ContainsKey($_.Format) } | Sort-Object Format, Path, Id -Unique)
    Write-Host "Retry items: $($failedItems.Count)" -ForegroundColor Cyan

    $tempRoot = Join-Path $RecoveryRoot "_temp"
    if (Test-Path -LiteralPath $tempRoot) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force
    }
    Ensure-Dir $tempRoot

    $results = @()
    $oneNote = $null

    try {
        $oneNote = New-OneNoteComApplication

        foreach ($item in $failedItems) {
            $targetPath = $item.Path
            $targetDir = Split-Path $targetPath -Parent
            $extension = [System.IO.Path]::GetExtension($targetPath)
            $originalFileName = [System.IO.Path]::GetFileName($targetPath)
            $hash = Get-Sha1Short ($item.Id + "|" + $targetPath) 12
            $tempItemDir = Join-Path $tempRoot $hash
            $tempFileName = $originalFileName

            if ($tempFileName.Length -gt 140) {
                $base = [System.IO.Path]::GetFileNameWithoutExtension($tempFileName)
                $tempFileName = (Shorten-Component $base 100) + $extension
            }

            $tempPath = Join-Path $tempItemDir $tempFileName
            Write-Host ""
            Write-Host "Retrying $($item.Format): $targetPath" -ForegroundColor Yellow

            try {
                if (Test-Path -LiteralPath $tempItemDir) {
                    Remove-Item -LiteralPath $tempItemDir -Recurse -Force
                }
                Ensure-Dir $tempItemDir

                Invoke-OneNotePublish -OneNote $oneNote -PageId $item.Id -Destination $tempPath -Format $FormatMap[$item.Format]

                $savedPath = $null
                $mode = $null

                try {
                    Ensure-Dir $targetDir
                    if ((Test-Path -LiteralPath $targetPath) -and $OverwriteExisting) {
                        Remove-Item -LiteralPath $targetPath -Force
                    }

                    if ((Test-Path -LiteralPath $targetPath) -and (-not $OverwriteExisting)) {
                        $savedPath = $targetPath
                        $mode = "already_exists"
                    }
                    else {
                        Copy-TreeToDestination -SourceDir $tempItemDir -DestinationDir $targetDir
                        $savedPath = if ($tempFileName -eq $originalFileName) { $targetPath } else { Join-Path $targetDir $tempFileName }
                        $mode = "original_path"
                    }
                }
                catch {
                    $shortRecoveryPath = Get-ShortRecoveryPath -OriginalPath $targetPath -ExportRoot $ExportRoot -RecoveryRoot $RecoveryRoot
                    $shortRecoveryDir = Split-Path $shortRecoveryPath -Parent
                    Ensure-Dir $shortRecoveryDir
                    Copy-TreeToDestination -SourceDir $tempItemDir -DestinationDir $shortRecoveryDir
                    $savedPath = Join-Path $shortRecoveryDir $tempFileName
                    $mode = "recovery_short_path"
                }

                Write-Host "OK -> $savedPath" -ForegroundColor Green
                $results += [PSCustomObject]@{
                    Status       = "OK"
                    Mode         = $mode
                    Format       = $item.Format
                    OriginalPath = $targetPath
                    SavedPath    = $savedPath
                    Id           = $item.Id
                    Error        = ""
                }
            }
            catch {
                Write-Host "FAILED again: $($_.Exception.Message)" -ForegroundColor Red
                $results += [PSCustomObject]@{
                    Status       = "FAILED"
                    Mode         = "retry_failed"
                    Format       = $item.Format
                    OriginalPath = $targetPath
                    SavedPath    = ""
                    Id           = $item.Id
                    Error        = $_.Exception.Message
                }
            }
        }
    }
    finally {
        Release-ComObject $oneNote
    }

    Export-Rows -Rows $results -Path $RetryResultsCsv

    $failedAgain = @($results | Where-Object { $_.Status -eq "FAILED" })
    if ($failedAgain.Count -gt 0) {
        $failedAgainPath = Join-Path $RecoveryRoot "still-failed.txt"
        $failedAgain | ForEach-Object {
            "$($_.Format)`t$($_.OriginalPath)`t$($_.Error)"
        } | Set-Content -Path $failedAgainPath -Encoding UTF8
        Write-Host "Still failed list: $failedAgainPath" -ForegroundColor Yellow
    }

    Write-Host "Result CSV: $RetryResultsCsv" -ForegroundColor Cyan
}

function Invoke-RetryV2 {
    param(
        [string]$RetryResultsCsv,
        [string]$V2ResultsCsv,
        [string]$ShortRoot,
        [switch]$IncludeUnnamed,
        [switch]$Overwrite
    )

    Write-StepHeader "Step: RetryV2 - retry remaining failures through short subst path"
    if (-not (Test-Path -LiteralPath $RetryResultsCsv -PathType Leaf)) {
        throw "Retry results CSV not found: $RetryResultsCsv"
    }

    $drive = ""
    $oneNote = $null
    $results = @()

    try {
        $drive = Start-SubstDrive -TargetPath $ShortRoot
        $items = Import-Csv -LiteralPath $RetryResultsCsv | Where-Object {
            $_.Status -eq "FAILED" -and $FormatMap.ContainsKey($_.Format)
        }
        if (-not $IncludeUnnamed) {
            $items = $items | Where-Object {
                ([System.IO.Path]::GetFileNameWithoutExtension($_.OriginalPath)) -notlike "Unbenannt*"
            }
        }
        $items = @($items | Sort-Object Format, OriginalPath, Id -Unique)
        Write-Host "Retrying failed items: $($items.Count)" -ForegroundColor Cyan

        $oneNote = New-OneNoteComApplication
        $outDirDrive = Join-Path $drive "recovered"
        Ensure-Dir $outDirDrive

        foreach ($item in $items) {
            $originalPath = $item.OriginalPath
            $originalBase = [System.IO.Path]::GetFileNameWithoutExtension($originalPath)
            $originalExt = [System.IO.Path]::GetExtension($originalPath)
            $hash = Get-Sha1Short ($item.Id + "|" + $originalPath + "|" + $item.Format) 12
            $shortBase = Shorten-Component $originalBase 42

            $pageFolderName = "$hash-$shortBase"
            $pageFolderDrive = Join-Path $outDirDrive $pageFolderName
            Ensure-Dir $pageFolderDrive

            $outFileName = "$shortBase$originalExt"
            $outFileDrive = Join-Path $pageFolderDrive $outFileName
            $outFileReal = Join-Path (Join-Path (Join-Path $ShortRoot "recovered") $pageFolderName) $outFileName

            Write-Host ""
            Write-Host "Retrying $($item.Format): $originalPath" -ForegroundColor Yellow
            Write-Host "Short path: $outFileDrive"

            try {
                if ((Test-Path -LiteralPath $outFileDrive) -and $Overwrite) {
                    Remove-Item -LiteralPath $outFileDrive -Force
                }
                if ((Test-Path -LiteralPath $outFileDrive) -and (-not $Overwrite)) {
                    $results += [PSCustomObject]@{
                        Status       = "ALREADY_EXISTS"
                        Format       = $item.Format
                        OriginalPath = $originalPath
                        SavedPath    = $outFileReal
                        Id           = $item.Id
                        Error        = ""
                    }
                    continue
                }

                Invoke-OneNotePublish -OneNote $oneNote -PageId $item.Id -Destination $outFileDrive -Format $FormatMap[$item.Format]

                $results += [PSCustomObject]@{
                    Status       = "OK"
                    Format       = $item.Format
                    OriginalPath = $originalPath
                    SavedPath    = $outFileReal
                    Id           = $item.Id
                    Error        = ""
                }
                Write-Host "OK -> $outFileReal" -ForegroundColor Green
            }
            catch {
                $results += [PSCustomObject]@{
                    Status       = "FAILED"
                    Format       = $item.Format
                    OriginalPath = $originalPath
                    SavedPath    = $outFileReal
                    Id           = $item.Id
                    Error        = $_.Exception.Message
                }
                Write-Host "FAILED again: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    finally {
        Release-ComObject $oneNote
        Stop-SubstDrive -Drive $drive
    }

    Export-Rows -Rows $results -Path $V2ResultsCsv

    $stillFailed = @($results | Where-Object { $_.Status -eq "FAILED" })
    $stillFailedTxt = Join-Path $ShortRoot "still-failed-v2.txt"
    $stillFailed | ForEach-Object {
        "$($_.Format)`t$($_.OriginalPath)`t$($_.Error)"
    } | Set-Content -Path $stillFailedTxt -Encoding UTF8

    Write-Host "Result CSV: $V2ResultsCsv" -ForegroundColor Cyan
    Write-Host "Still failed: $stillFailedTxt" -ForegroundColor Cyan
}

function Invoke-MergeRecoveries {
    param(
        [string]$ExportRoot,
        [string]$V1Csv,
        [string]$V2Csv,
        [string]$MergedRootName,
        [switch]$Overwrite
    )

    Write-StepHeader "Step: Merge - merge recovered V1/V2 files into export root"
    $mergedRoot = Join-Path $ExportRoot $MergedRootName
    Ensure-Dir $mergedRoot

    $rows = @()
    $rows += Read-RecoveryRows -CsvPath $V1Csv -BatchName "v1"
    $rows += Read-RecoveryRows -CsvPath $V2Csv -BatchName "v2"

    $deduped = $rows |
        Sort-Object @{Expression = { if ($_.Batch -eq "v2") { 0 } else { 1 } }}, @{Expression = { if ($_.Status -eq "OK") { 0 } else { 1 } }}, OriginalPath, Format |
        Group-Object OriginalPath, Format |
        ForEach-Object { $_.Group[0] }

    $index = @()

    foreach ($row in $deduped) {
        $source = $row.SavedPath
        $original = $row.OriginalPath

        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
            $index += [PSCustomObject]@{
                Status       = "SOURCE_MISSING"
                Batch        = $row.Batch
                Format       = $row.Format
                OriginalPath = $original
                SourcePath   = $source
                MergedPath   = ""
                Note         = "SavedPath from CSV does not exist"
            }
            continue
        }

        $sourceFull = [System.IO.Path]::GetFullPath($source)
        $exportFull = [System.IO.Path]::GetFullPath($ExportRoot).TrimEnd("\", "/")
        if ($sourceFull.StartsWith($exportFull, [System.StringComparison]::OrdinalIgnoreCase) -and
            $sourceFull -notlike "*\$MergedRootName\*") {

            $index += [PSCustomObject]@{
                Status       = "ALREADY_IN_EXPORT"
                Batch        = $row.Batch
                Format       = $row.Format
                OriginalPath = $original
                SourcePath   = $source
                MergedPath   = $source
                Note         = "Already located inside main export"
            }
            continue
        }

        $relative = Get-RelativePathFromRoot -Path $original -Root $ExportRoot
        $parts = @($relative -split "[\\/]" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $formatFolder = if ($row.Format -eq "pfPDF") { "pdf" } elseif ($row.Format -eq "pfHTML") { "html" } else { "other" }
        $context = @($parts | Select-Object -First ([Math]::Min(2, [Math]::Max(0, $parts.Count - 1)))) | ForEach-Object { Shorten-Component $_ 45 }
        $pageName = Shorten-Component ([System.IO.Path]::GetFileNameWithoutExtension($original)) 65
        $hash = Get-Sha1Short ($original + "|" + $row.Format) 10

        $destDir = Join-Path $mergedRoot $formatFolder
        foreach ($part in $context) {
            $destDir = Join-Path $destDir $part
        }
        $destDir = Join-Path $destDir "$hash-$pageName"

        try {
            $copyResult = Copy-FileAndHtmlResources -SourceFile $source -DestDir $destDir -Overwrite:$Overwrite
            $index += [PSCustomObject]@{
                Status       = "MERGED"
                Batch        = $row.Batch
                Format       = $row.Format
                OriginalPath = $original
                SourcePath   = $source
                MergedPath   = $copyResult.CopiedFile
                Note         = $copyResult.Note
            }
        }
        catch {
            $index += [PSCustomObject]@{
                Status       = "MERGE_FAILED"
                Batch        = $row.Batch
                Format       = $row.Format
                OriginalPath = $original
                SourcePath   = $source
                MergedPath   = ""
                Note         = $_.Exception.Message
            }
        }
    }

    $indexPath = Join-Path $mergedRoot "_recovered-index.csv"
    Export-Rows -Rows $index -Path $indexPath

    $readmePath = Join-Path $mergedRoot "_README.txt"
    @"
OneNote recovery merge

This folder contains recovered OneNote exports that could not be written back
to their original path, usually because of path length or problematic file names.

Use _recovered-index.csv:
- OriginalPath is the old target.
- SourcePath is the successful recovery file.
- MergedPath is the copied file in this folder.
- HTML files may have sibling resource folders such as *-Dateien or *_files.
"@ | Set-Content -Path $readmePath -Encoding UTF8

    Write-Host "Merged root: $mergedRoot" -ForegroundColor Cyan
    Write-Host "Index: $indexPath" -ForegroundColor Cyan
}

function Invoke-FixV1Misclassified {
    param(
        [string]$IndexCsv,
        [string]$ExportRoot,
        [string]$RecoveryRoot,
        [string]$OutFolderName,
        [switch]$Overwrite
    )

    Write-StepHeader "Step: FixV1Misclassified - copy V1 files marked as already in export"
    if (-not (Test-Path -LiteralPath $IndexCsv -PathType Leaf)) {
        throw "Index CSV not found: $IndexCsv"
    }

    $outRoot = Join-Path $ExportRoot $OutFolderName
    Ensure-Dir $outRoot
    $recoveryFull = [System.IO.Path]::GetFullPath($RecoveryRoot).TrimEnd("\", "/")
    $rows = Import-Csv -LiteralPath $IndexCsv
    $toFix = @($rows | Where-Object {
        $_.Status -eq "ALREADY_IN_EXPORT" -and
        [System.IO.Path]::GetFullPath($_.SourcePath).StartsWith($recoveryFull, [System.StringComparison]::OrdinalIgnoreCase)
    })

    $fixed = @()
    foreach ($row in $toFix) {
        $formatFolder = if ($row.Format -eq "pfPDF") { "pdf" } elseif ($row.Format -eq "pfHTML") { "html" } else { "other" }
        $hash = Get-Sha1Short ($row.OriginalPath + "|" + $row.Format) 10
        $pageFolder = Shorten-Component ([System.IO.Path]::GetFileNameWithoutExtension($row.OriginalPath)) 65
        $destDir = Join-Path (Join-Path $outRoot $formatFolder) "$hash-$pageFolder"

        try {
            $copyResult = Copy-FileAndHtmlResources -SourceFile $row.SourcePath -DestDir $destDir -Overwrite:$Overwrite
            $fixed += [PSCustomObject]@{
                Status       = "FIXED_COPIED"
                Format       = $row.Format
                OriginalPath = $row.OriginalPath
                SourcePath   = $row.SourcePath
                FixedPath    = $copyResult.CopiedFile
                Error        = ""
            }
        }
        catch {
            $fixed += [PSCustomObject]@{
                Status       = "FIX_FAILED"
                Format       = $row.Format
                OriginalPath = $row.OriginalPath
                SourcePath   = $row.SourcePath
                FixedPath    = ""
                Error        = $_.Exception.Message
            }
        }
    }

    $outCsv = Join-Path $outRoot "_v1-fixed-index.csv"
    Export-Rows -Rows $fixed -Path $outCsv
    Write-Host "Fixed index: $outCsv" -ForegroundColor Cyan
}

function Invoke-FixV1Flat {
    param(
        [string]$V1FixedIndex,
        [string]$ExportRoot,
        [string]$OutFolderName,
        [switch]$Overwrite
    )

    Write-StepHeader "Step: FixV1Flat - flatten V1 fix failures into short folders"
    if (-not (Test-Path -LiteralPath $V1FixedIndex -PathType Leaf)) {
        throw "Index not found: $V1FixedIndex"
    }

    $outRoot = Join-Path $ExportRoot $OutFolderName
    Ensure-Dir $outRoot
    $rows = @(Import-Csv -LiteralPath $V1FixedIndex | Where-Object { $_.Status -eq "FIX_FAILED" })
    $results = @()

    foreach ($row in $rows) {
        $hash = Get-Sha1Short ($row.OriginalPath + "|" + $row.Format) 12
        $shortFolder = Join-Path $outRoot $hash
        Ensure-Dir $shortFolder

        $shortFileName = if ($row.Format -eq "pfPDF") { "page.pdf" } else { "page.htm" }
        $destFile = Join-Path $shortFolder $shortFileName

        try {
            if (-not (Test-Path -LiteralPath $row.SourcePath -PathType Leaf)) {
                throw "Source file not found: $($row.SourcePath)"
            }

            if ((Test-Path -LiteralPath $destFile) -and (-not $Overwrite)) {
                $copyNote = "already_exists"
            }
            else {
                Copy-Item -LiteralPath $row.SourcePath -Destination $destFile -Force
                $copyNote = "copied"
            }

            if ($row.Format -eq "pfHTML") {
                $sourceBase = [System.IO.Path]::GetFileNameWithoutExtension($row.SourcePath)
                $sourceDir = Split-Path $row.SourcePath -Parent
                foreach ($resourceFolderName in @("$sourceBase-Dateien", "$sourceBase_files")) {
                    $resourceDir = Join-Path $sourceDir $resourceFolderName
                    if (Test-Path -LiteralPath $resourceDir -PathType Container) {
                        $destResourceDir = Join-Path $shortFolder "page-Dateien"
                        if ((Test-Path -LiteralPath $destResourceDir) -and $Overwrite) {
                            Remove-Item -LiteralPath $destResourceDir -Recurse -Force
                        }
                        if (-not (Test-Path -LiteralPath $destResourceDir)) {
                            Copy-DirectoryContent -SourceDir $resourceDir -DestDir $destResourceDir
                        }

                        $html = Get-Content -LiteralPath $destFile -Raw
                        $html = $html -replace [regex]::Escape([System.IO.Path]::GetFileName($resourceDir)), "page-Dateien"
                        Set-Content -LiteralPath $destFile -Value $html -Encoding UTF8
                        break
                    }
                }
            }

            $results += [PSCustomObject]@{
                Status       = "FLAT_COPIED"
                Format       = $row.Format
                OriginalPath = $row.OriginalPath
                SourcePath   = $row.SourcePath
                FlatPath     = $destFile
                Error        = $copyNote
            }
        }
        catch {
            $results += [PSCustomObject]@{
                Status       = "FLAT_FAILED"
                Format       = $row.Format
                OriginalPath = $row.OriginalPath
                SourcePath   = $row.SourcePath
                FlatPath     = ""
                Error        = $_.Exception.Message
            }
        }
    }

    $outCsv = Join-Path $outRoot "_v1-flat-index.csv"
    Export-Rows -Rows $results -Path $outCsv
    Write-Host "Flat index: $outCsv" -ForegroundColor Cyan
}

$effectiveRetryV1Overwrite = ($Overwrite -or $OverwriteExisting)

switch ($Step) {
    "RetryV1" {
        Invoke-RetryV1 -FailedLog $FailedLog -ExportRoot $ExportRoot -RecoveryRoot $RecoveryRoot -RetryResultsCsv $RetryResultsCsv -IncludeUnnamed:$IncludeUnnamed -OverwriteExisting:$effectiveRetryV1Overwrite
    }
    "RetryV2" {
        Invoke-RetryV2 -RetryResultsCsv $RetryResultsCsv -V2ResultsCsv $V2ResultsCsv -ShortRoot $ShortRoot -IncludeUnnamed:$IncludeUnnamed -Overwrite:$Overwrite
    }
    "Merge" {
        Invoke-MergeRecoveries -ExportRoot $ExportRoot -V1Csv $RetryResultsCsv -V2Csv $V2ResultsCsv -MergedRootName $MergedRootName -Overwrite:$Overwrite
    }
    "FixV1Misclassified" {
        Invoke-FixV1Misclassified -IndexCsv $RecoveredIndexCsv -ExportRoot $ExportRoot -RecoveryRoot $RecoveryRoot -OutFolderName $V1MisclassifiedOutFolderName -Overwrite:$Overwrite
    }
    "FixV1Flat" {
        Invoke-FixV1Flat -V1FixedIndex $V1FixedIndexCsv -ExportRoot $ExportRoot -OutFolderName $V1FlatOutFolderName -Overwrite:$Overwrite
    }
    "All" {
        Invoke-RetryV1 -FailedLog $FailedLog -ExportRoot $ExportRoot -RecoveryRoot $RecoveryRoot -RetryResultsCsv $RetryResultsCsv -IncludeUnnamed:$IncludeUnnamed -OverwriteExisting:$effectiveRetryV1Overwrite
        Invoke-RetryV2 -RetryResultsCsv $RetryResultsCsv -V2ResultsCsv $V2ResultsCsv -ShortRoot $ShortRoot -IncludeUnnamed:$IncludeUnnamed -Overwrite:$Overwrite
        Invoke-MergeRecoveries -ExportRoot $ExportRoot -V1Csv $RetryResultsCsv -V2Csv $V2ResultsCsv -MergedRootName $MergedRootName -Overwrite:$Overwrite
        Invoke-FixV1Misclassified -IndexCsv $RecoveredIndexCsv -ExportRoot $ExportRoot -RecoveryRoot $RecoveryRoot -OutFolderName $V1MisclassifiedOutFolderName -Overwrite:$Overwrite
        Invoke-FixV1Flat -V1FixedIndex $V1FixedIndexCsv -ExportRoot $ExportRoot -OutFolderName $V1FlatOutFolderName -Overwrite:$Overwrite
    }
}
