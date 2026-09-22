#############################################################################
##                    Publisher Converter                                  ##
#############################################################################
##                                                                         ##
## Name: Publisher_To_PDF_WORD.ps1                                         ##
## Date: 21.09.26                                                          ##
## Author: Rhys Dunn                                                       ##
## Description: Converts .PUB's to PDFS or DocXs                           ##
##                                                                         ##
#############################################################################

#requires -Version 5.1

#Change RootFolder to the path containing .PUB files
$RootFolder = "CHANGE_ME"

$LogFolder = "C:\PublisherLogs"

$OverwriteExistingOutput = $false

$PauseBeforeExit = $true

$ShowApplications = $true

$FileTimeoutSeconds = 120

$MaximumRetries = 1
$RetryDelaySeconds = 3

$KeepIntermediatePdf = $true

$PublisherPdfFormat   = 2      # pbFixedFormatTypePDF
$PublisherRtfFormat   = 6      # pbFileRTF
$WordDocxFormat       = 16     # wdFormatDocumentDefault
$WordDoNotSaveChanges = 0      # wdDoNotSaveChanges
$WordAlertsNone       = 0      # wdAlertsNone
$WordDisableMacros    = 3      # msoAutomationSecurityForceDisable


function Wait-BeforeExit {
    param (
        [Parameter()]
        [string]$Message = "Press Enter to close this window"
    )

    if (-not $PauseBeforeExit) {
        return
    }

    Write-Host ""
    Write-Host $Message -ForegroundColor Cyan

    try {
        [void](Read-Host)
    }
    catch {
        Start-Sleep -Seconds 10
    }
}


function Exit-Script {
    param (
        [Parameter()]
        [int]$ExitCode = 0,

        [Parameter()]
        [string]$Message = "Press Enter to close this window"
    )

    Wait-BeforeExit -Message $Message

    exit $ExitCode
}

if ($host.Name -eq "Windows PowerShell ISE Host") {

    Write-Host ""
    Write-Host "=====================================================" -ForegroundColor Red
    Write-Host " THIS SCRIPT CANNOT RUN IN POWERSHELL ISE" -ForegroundColor Red
    Write-Host "=====================================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "PowerShell ISE deadlocks with Office COM automation."
    Write-Host "Please run it from the normal PowerShell console:"
    Write-Host ""
    Write-Host "  powershell -ExecutionPolicy Bypass -File `"C:\Path\To\PubConvert.ps1`"" -ForegroundColor Cyan
    Write-Host ""

    return
}


function Get-YesNoAnswer {
    param ([Parameter(Mandatory)][string]$Question)

    while ($true) {

        $Answer = Read-Host -Prompt "$Question (yes/no)"

        if ($null -eq $Answer) { $Answer = "" }

        $Answer = $Answer.Trim().ToLower()

        if ($Answer -in @("yes", "y")) { return $true }
        if ($Answer -in @("no", "n"))  { return $false }

        Write-Host "Please type yes or no." -ForegroundColor Yellow
    }
}


function Get-OutputFormatAnswer {
    param ([Parameter(Mandatory)][string]$Question)

    while ($true) {

        $Answer = Read-Host -Prompt "$Question (pdf/word)"

        if ($null -eq $Answer) { $Answer = "" }

        $Answer = $Answer.Trim().ToLower()

        if ($Answer -in @("pdf", "p"))                 { return "PDF" }
        if ($Answer -in @("word", "w", "docx", "doc")) { return "WORD" }

        Write-Host "Please type pdf or word." -ForegroundColor Yellow
    }
}




$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $RootFolder -PathType Container)) {

    Write-Host ""
    Write-Host "ERROR: the root folder does not exist:" -ForegroundColor Red
    Write-Host "  $RootFolder" -ForegroundColor Red
    Write-Host ""
    Write-Host "Edit the `$RootFolder setting at the top of the script." -ForegroundColor Yellow

    Exit-Script -ExitCode 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Publisher bulk converter" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Folder to search: $RootFolder"
Write-Host ""

$DryRun = Get-YesNoAnswer `
    -Question "Do you want a DRY RUN first (count files without converting)?"

$OutputFormat = Get-OutputFormatAnswer `
    -Question "Convert the Publisher files to which format?"

if ($OutputFormat -eq "PDF") {
    $OutputExtension = ".pdf"
    $OutputDescription = "PDF"
}
else {
    $OutputExtension = ".docx"
    $OutputDescription = "Word document"
}

Write-Host ""

if ($DryRun) {
    Write-Host "Mode: DRY RUN - nothing will be converted." -ForegroundColor Yellow
}
else {
    Write-Host "Mode: LIVE RUN - files will be converted." -ForegroundColor Green
}

Write-Host "Output format: $OutputDescription ($OutputExtension)"

if ((-not $DryRun) -and ($OutputFormat -eq "WORD")) {
    Write-Host ""
    Write-Host "PLEASE NOTE about Word output:" -ForegroundColor Yellow
    Write-Host "  Text and character formatting will carry across." -ForegroundColor Yellow
    Write-Host "  Graphics, columns and page design will NOT be" -ForegroundColor Yellow
    Write-Host "  preserved faithfully. Choose PDF for an archive copy." -ForegroundColor Yellow
}

if ((-not $DryRun) -and ($FileTimeoutSeconds -gt 0)) {
    Write-Host ""
    Write-Host "NOTE: a watchdog closes Publisher and Word if a file takes" -ForegroundColor Yellow
    Write-Host "      longer than $FileTimeoutSeconds seconds." -ForegroundColor Yellow
    Write-Host "      Close your own Publisher and Word documents first." -ForegroundColor Yellow
}

Write-Host ""



if ([string]::IsNullOrWhiteSpace($LogFolder)) {
    $LogOutputFolder = $RootFolder
}
else {
    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) {
        try {
            New-Item -Path $LogFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Host ""
            Write-Host "ERROR: the log folder could not be created:" -ForegroundColor Red
            Write-Host "  $LogFolder" -ForegroundColor Red
            Write-Host "  $($_.Exception.Message)" -ForegroundColor Red

            Exit-Script -ExitCode 1
        }
    }

    $LogOutputFolder = $LogFolder
}

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

if ($DryRun) { $RunLabel = "DryRun" } else { $RunLabel = "Convert" }

$LogFile = Join-Path -Path $LogOutputFolder `
    -ChildPath "Publisher_$($RunLabel)_$($OutputFormat)_$Timestamp.log"

$CsvReport = Join-Path -Path $LogOutputFolder `
    -ChildPath "Publisher_$($RunLabel)_$($OutputFormat)_$Timestamp.csv"

$WatchdogMarker = Join-Path -Path $env:TEMP `
    -ChildPath "PubConvert_Watchdog_$Timestamp.marker"


function Write-ConversionLog {
    param (
        [Parameter(Mandatory)][string]$Message,

        [ValidateSet("INFO", "STEP", "SUCCESS", "SKIPPED", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )

    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Entry = "$Time [$Level] $Message"

    switch ($Level) {
        "STEP"    { Write-Host $Entry -ForegroundColor Gray }
        "SUCCESS" { Write-Host $Entry -ForegroundColor Green }
        "SKIPPED" { Write-Host $Entry -ForegroundColor Yellow }
        "WARNING" { Write-Host $Entry -ForegroundColor DarkYellow }
        "ERROR"   { Write-Host $Entry -ForegroundColor Red }
        default   { Write-Host $Entry }
    }

    Add-Content -LiteralPath $LogFile -Value $Entry -Encoding UTF8
}


function Release-ComObject {
    param ([Parameter()]$ComObject)

    if ($null -ne $ComObject) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
        catch {
        }
    }
}


$script:PdfWarningSavedState = @()

function Disable-WordPdfWarning {


    if ($script:PdfWarningSavedState.Count -gt 0) {
        return
    }

    $OfficeRoot = "HKCU:\Software\Microsoft\Office"

    if (-not (Test-Path -LiteralPath $OfficeRoot)) {
        Write-ConversionLog -Level "WARNING" `
            -Message "Office registry key not found. The PDF dialog may appear."
        return
    }

    $VersionKeys = Get-ChildItem -LiteralPath $OfficeRoot -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' }

    if (-not $VersionKeys) {
        return
    }

    foreach ($VersionKey in $VersionKeys) {

        $Version = $VersionKey.PSChildName
        $WordPath = "$OfficeRoot\$Version\Word"
        $OptionsPath = "$OfficeRoot\$Version\Word\Options"

        if (-not (Test-Path -LiteralPath $WordPath)) {
            continue
        }

        try {
            $KeyExisted = Test-Path -LiteralPath $OptionsPath

            if (-not $KeyExisted) {
                New-Item -Path $OptionsPath -Force -ErrorAction Stop | Out-Null
            }

            $PreviousValue = $null
            $ValueExisted = $false

            try {
                $Property = Get-ItemProperty -LiteralPath $OptionsPath `
                    -Name "DisableConvertPDFWarning" -ErrorAction Stop

                $PreviousValue = $Property.DisableConvertPDFWarning
                $ValueExisted = $true
            }
            catch {
                $ValueExisted = $false
            }

            $script:PdfWarningSavedState += [PSCustomObject]@{
                Path          = $OptionsPath
                KeyExisted    = $KeyExisted
                ValueExisted  = $ValueExisted
                PreviousValue = $PreviousValue
            }

            Set-ItemProperty -LiteralPath $OptionsPath `
                -Name "DisableConvertPDFWarning" -Value 1 -Type DWord -ErrorAction Stop

            Write-ConversionLog -Message "PDF conversion dialog suppressed for Office $Version."
        }
        catch {
            Write-ConversionLog -Level "WARNING" `
                -Message "Could not suppress the PDF dialog for Office $Version`: $($_.Exception.Message)"
        }
    }
}


function Restore-WordPdfWarning {

    foreach ($State in $script:PdfWarningSavedState) {

        try {
            if ($State.ValueExisted) {
                Set-ItemProperty -LiteralPath $State.Path `
                    -Name "DisableConvertPDFWarning" `
                    -Value $State.PreviousValue -Type DWord -ErrorAction Stop
            }
            else {
                Remove-ItemProperty -LiteralPath $State.Path `
                    -Name "DisableConvertPDFWarning" -ErrorAction SilentlyContinue
            }
        }
        catch {
        }
    }

    $script:PdfWarningSavedState = @()
}




$script:EnumCache = @{}

function Get-OfficeEnumValue {
    param (
        [Parameter(Mandatory)][string]$TypeName,
        [Parameter(Mandatory)][string]$MemberName,
        [Parameter(Mandatory)][AllowNull()]$Fallback
    )

    $CacheKey = "$TypeName|$MemberName"

    if ($script:EnumCache.ContainsKey($CacheKey)) {
        return $script:EnumCache[$CacheKey]
    }

    $Resolved = $null

    foreach ($Assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {

        try {
            $EnumType = $Assembly.GetType($TypeName, $false)

            if (($null -ne $EnumType) -and $EnumType.IsEnum) {
                if ([Enum]::GetNames($EnumType) -contains $MemberName) {
                    $Resolved = [Enum]::Parse($EnumType, $MemberName)
                    break
                }
            }
        }
        catch {
        }
    }

    if ($null -eq $Resolved) { $Resolved = $Fallback }

    $script:EnumCache[$CacheKey] = $Resolved

    return $Resolved
}


function Get-PbDoNotSaveChangesValue {
    return Get-OfficeEnumValue `
        -TypeName "Microsoft.Office.Interop.Publisher.PbSaveOptions" `
        -MemberName "pbDoNotSaveChanges" `
        -Fallback $null
}


function Initialize-OfficeEnums {

    $script:PublisherPdfFormat = Get-OfficeEnumValue `
        -TypeName "Microsoft.Office.Interop.Publisher.PbFixedFormatType" `
        -MemberName "pbFixedFormatTypePDF" -Fallback 2

    $script:PublisherRtfFormat = Get-OfficeEnumValue `
        -TypeName "Microsoft.Office.Interop.Publisher.PbFileFormat" `
        -MemberName "pbFileRTF" -Fallback 6

    $script:WordDocxFormat = Get-OfficeEnumValue `
        -TypeName "Microsoft.Office.Interop.Word.WdSaveFormat" `
        -MemberName "wdFormatDocumentDefault" -Fallback 16

    $script:WordDoNotSaveChanges = Get-OfficeEnumValue `
        -TypeName "Microsoft.Office.Interop.Word.WdSaveOptions" `
        -MemberName "wdDoNotSaveChanges" -Fallback 0
}


function Open-Publication {
    param (
        [Parameter(Mandatory)]$PublisherApp,
        [Parameter(Mandatory)][string]$Path
    )

    $LastErrorMessage = "No Open attempt was made."
    $SaveOption = Get-PbDoNotSaveChangesValue

    if ($null -ne $SaveOption) {
        try {
            return $PublisherApp.Open($Path, $false, $false, $SaveOption)
        }
        catch {
            $LastErrorMessage = $_.Exception.Message
        }
    }

    try {
        return $PublisherApp.Open($Path, $false, $false)
    }
    catch {
        $LastErrorMessage = $_.Exception.Message
    }

    try {
        return $PublisherApp.Open($Path)
    }
    catch {
        $LastErrorMessage = $_.Exception.Message
    }

    throw "Publisher could not open the publication. Last error: $LastErrorMessage"
}


function Save-WordDocumentAs {
    param (
        [Parameter(Mandatory)]$Document,
        [Parameter(Mandatory)][string]$DestinationPath,
        [Parameter(Mandatory)]$Format
    )

    $Errors = @()

    try {
        $Document.SaveAs2($DestinationPath, $Format)
        return "SaveAs2 direct"
    }
    catch {
        $Errors += "SaveAs2 direct: $($_.Exception.Message)"
    }

    try {
        $Document.SaveAs($DestinationPath, $Format)
        return "SaveAs direct"
    }
    catch {
        $Errors += "SaveAs direct: $($_.Exception.Message)"
    }

    $LocalTempPath = Join-Path -Path $env:TEMP `
        -ChildPath ("PubConv_" + [guid]::NewGuid().ToString("N").Substring(0, 8) + ".docx")

    $SavedLocally = $false

    try {
        $Document.SaveAs2($LocalTempPath, $Format)
        $SavedLocally = $true
    }
    catch {
        $Errors += "SaveAs2 local: $($_.Exception.Message)"

        try {
            $Document.SaveAs($LocalTempPath, $Format)
            $SavedLocally = $true
        }
        catch {
            $Errors += "SaveAs local: $($_.Exception.Message)"
        }
    }

    if ($SavedLocally) {

        Start-Sleep -Milliseconds 500

        if (Test-Path -LiteralPath $LocalTempPath -PathType Leaf) {

            try {
                $DestinationFolder = Split-Path -Path $DestinationPath -Parent

                if (-not (Test-Path -LiteralPath $DestinationFolder -PathType Container)) {
                    New-Item -Path $DestinationFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
                }

                Copy-Item -LiteralPath $LocalTempPath `
                    -Destination $DestinationPath -Force -ErrorAction Stop

                Remove-Item -LiteralPath $LocalTempPath -Force -ErrorAction SilentlyContinue

                return "Local save then copy"
            }
            catch {
                $Errors += "Copy to destination: $($_.Exception.Message)"
                Remove-Item -LiteralPath $LocalTempPath -Force -ErrorAction SilentlyContinue
            }
        }
        else {
            $Errors += "Local save reported success but no file was found."
        }
    }

    throw ("Word could not save the document. Attempts: " + ($Errors -join " | "))
}


function Start-Watchdog {
    param (
        [Parameter(Mandatory)][int]$TimeoutSeconds,
        [Parameter(Mandatory)][string]$MarkerPath
    )

    if ($TimeoutSeconds -le 0) { return $null }

    return Start-Job -ScriptBlock {

        param ($Seconds, $Marker)

        Start-Sleep -Seconds $Seconds

        $Killed = @()

        foreach ($Name in @("MSPUB", "WINWORD")) {

            $Processes = Get-Process -Name $Name -ErrorAction SilentlyContinue

            if ($Processes) {
                foreach ($Process in $Processes) {
                    try {
                        $Process.Kill()
                        $Killed += $Name
                    }
                    catch {
                    }
                }
            }
        }

        Set-Content -LiteralPath $Marker -Value ($Killed -join ",") -Force

    } -ArgumentList $TimeoutSeconds, $MarkerPath
}


function Stop-Watchdog {
    param ([Parameter()]$WatchdogJob)

    if ($null -eq $WatchdogJob) { return }

    try { Stop-Job -Job $WatchdogJob -ErrorAction SilentlyContinue } catch { }
    try { Remove-Job -Job $WatchdogJob -Force -ErrorAction SilentlyContinue } catch { }
}


function Test-WatchdogFired {
    param ([Parameter(Mandatory)][string]$MarkerPath)

    if (Test-Path -LiteralPath $MarkerPath -PathType Leaf) {
        Remove-Item -LiteralPath $MarkerPath -Force -ErrorAction SilentlyContinue
        return $true
    }

    return $false
}


Write-ConversionLog -Message "Run started."
Write-ConversionLog -Message "Host: $($host.Name)"
Write-ConversionLog -Message "Root folder: $RootFolder"
Write-ConversionLog -Message "Log folder: $LogOutputFolder"
Write-ConversionLog -Message "Output format: $OutputDescription ($OutputExtension)"
Write-ConversionLog -Message "Dry run: $DryRun"
Write-ConversionLog -Message "Per-file timeout: $FileTimeoutSeconds seconds"

try {
    $PubFiles = @(
        Get-ChildItem -LiteralPath $RootFolder -Recurse -File -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -ieq ".pub" }
    )
}
catch {
    Write-ConversionLog -Level "ERROR" -Message "Unable to search the root folder: $($_.Exception.Message)"
    Exit-Script -ExitCode 1
}

$TotalFiles = $PubFiles.Count

Write-ConversionLog -Message "Publisher files found: $TotalFiles"

if ($TotalFiles -eq 0) {

    Write-ConversionLog -Level "WARNING" -Message "No Publisher files were found."

    Write-Host ""
    Write-Host "No .pub files were found under:" -ForegroundColor Yellow
    Write-Host "  $RootFolder" -ForegroundColor Yellow

    Exit-Script -ExitCode 0
}


if ($DryRun) {

    Write-ConversionLog -Level "WARNING" -Message "DRY RUN. Nothing will be converted."

    $WouldConvert = 0
    $WouldSkip = 0
    $PreviewResults = [System.Collections.Generic.List[object]]::new()

    foreach ($File in $PubFiles) {

        $OutputPath = [System.IO.Path]::ChangeExtension($File.FullName, $OutputExtension)
        $OutputExists = Test-Path -LiteralPath $OutputPath -PathType Leaf

        if ($OutputExists -and (-not $OverwriteExistingOutput)) {
            $PlannedAction = "Would skip ($OutputExtension already exists)"
            $WouldSkip++
        }
        elseif ($OutputExists -and $OverwriteExistingOutput) {
            $PlannedAction = "Would overwrite existing $OutputExtension"
            $WouldConvert++
        }
        else {
            $PlannedAction = "Would convert to $OutputDescription"
            $WouldConvert++
        }

        Write-ConversionLog -Message "$PlannedAction`: $($File.FullName)"

        $PreviewResults.Add([PSCustomObject]@{
            SourcePUB     = $File.FullName
            PlannedOutput = $OutputPath
            OutputFormat  = $OutputFormat
            PlannedAction = $PlannedAction
            PUBSizeBytes  = $File.Length
            LastModified  = $File.LastWriteTime
        })
    }

    try {
        $PreviewResults | Export-Csv -LiteralPath $CsvReport -NoTypeInformation -Encoding UTF8
        Write-ConversionLog -Level "SUCCESS" -Message "Dry-run preview saved: $CsvReport"
    }
    catch {
        Write-ConversionLog -Level "ERROR" -Message "Preview could not be saved: $($_.Exception.Message)"
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " DRY RUN COMPLETE - nothing was changed" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Target format:         $OutputDescription ($OutputExtension)"
    Write-Host "Publisher files found: $TotalFiles"
    Write-Host "Would convert:         $WouldConvert" -ForegroundColor Green
    Write-Host "Would skip:            $WouldSkip" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Log:         $LogFile"
    Write-Host "Preview CSV: $CsvReport"
    Write-Host ""
    Write-Host "Run the script again and answer 'no' to the dry-run question to convert." -ForegroundColor Cyan

    Exit-Script -ExitCode 0
}


$Publisher = $null
$Publication = $null
$Word = $null
$WordDocument = $null

function Start-OfficeApplications {

    if ($null -ne $script:Word) {
        Release-ComObject -ComObject $script:Word
        $script:Word = $null
    }

    if ($null -ne $script:Publisher) {
        Release-ComObject -ComObject $script:Publisher
        $script:Publisher = $null
    }

    try {
        Write-ConversionLog -Message "Starting Microsoft Publisher."
        $script:Publisher = New-Object -ComObject Publisher.Application
        Write-ConversionLog -Level "SUCCESS" -Message "Microsoft Publisher started."
    }
    catch {
        Write-ConversionLog -Level "ERROR" -Message "Publisher could not be started: $($_.Exception.Message)"
        return $false
    }

    if ($OutputFormat -eq "WORD") {

        Disable-WordPdfWarning

        try {
            Write-ConversionLog -Message "Starting Microsoft Word."
            $script:Word = New-Object -ComObject Word.Application

            try {
                $script:Word.Visible = $ShowApplications
                $script:Word.DisplayAlerts = $WordAlertsNone
                $script:Word.AutomationSecurity = $WordDisableMacros
                $script:Word.Options.ConfirmConversions = $false
                $script:Word.Options.SaveInterval = 0
            }
            catch {
            }

            Write-ConversionLog -Level "SUCCESS" -Message "Microsoft Word started."
        }
        catch {
            Write-ConversionLog -Level "ERROR" -Message "Word could not be started: $($_.Exception.Message)"
            return $false
        }
    }

    Initialize-OfficeEnums

    return $true
}

if (-not (Start-OfficeApplications)) {

    Write-ConversionLog -Level "ERROR" -Message "Office could not be started. Stopping."

    Restore-WordPdfWarning

    Exit-Script -ExitCode 1
}



$SuccessCount = 0
$SkippedCount = 0
$FailedCount = 0
$TimedOutCount = 0
$ProcessedCount = 0

$Results = [System.Collections.Generic.List[object]]::new()
$OverallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    foreach ($File in $PubFiles) {

        $ProcessedCount++
        $Publication = $null
        $WordDocument = $null
        $TempPdfPath = $null

        $PercentComplete = [math]::Round((($ProcessedCount - 1) / $TotalFiles) * 100, 0)

        Write-Progress `
            -Activity "Converting Publisher files to $OutputDescription" `
            -Status "$ProcessedCount of $TotalFiles - $($File.Name)" `
            -PercentComplete $PercentComplete

        $OutputPath = [System.IO.Path]::ChangeExtension($File.FullName, $OutputExtension)

        if ((Test-Path -LiteralPath $OutputPath -PathType Leaf) -and (-not $OverwriteExistingOutput)) {

            Write-ConversionLog -Level "SKIPPED" -Message "Output already exists: $OutputPath"
            $SkippedCount++

            $Results.Add([PSCustomObject]@{
                SourcePUB       = $File.FullName
                OutputFile      = $OutputPath
                OutputFormat    = $OutputFormat
                Status          = "Skipped"
                Method          = ""
                AttemptCount    = 0
                DurationSeconds = 0
                OutputSizeBytes = (Get-Item -LiteralPath $OutputPath).Length
                Message         = "Matching output already exists"
            })

            continue
        }

        $FileStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $Converted = $false
        $LastErrorMessage = ""
        $MethodUsed = ""
        $Attempt = 0
        $WasTimedOut = $false

        while ((-not $Converted) -and ($Attempt -le $MaximumRetries)) {

            $Attempt++
            $Watchdog = $null

            try {
                Write-ConversionLog -Message "Attempt $Attempt`: $($File.FullName)"

                if ($null -ne $WordDocument) {
                    try { $WordDocument.Close($WordDoNotSaveChanges) } catch { }
                    Release-ComObject -ComObject $WordDocument
                    $WordDocument = $null
                }

                if ($null -ne $Publication) {
                    try { $Publication.Close() } catch { }
                    Release-ComObject -ComObject $Publication
                    $Publication = $null
                }

                if ($OverwriteExistingOutput -and (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
                    Remove-Item -LiteralPath $OutputPath -Force -ErrorAction Stop
                }

                $Watchdog = Start-Watchdog `
                    -TimeoutSeconds $FileTimeoutSeconds `
                    -MarkerPath $WatchdogMarker

                # --- Step 1 -------------------------------------------
                Write-ConversionLog -Level "STEP" -Message "  Step 1: opening publication in Publisher..."

                $Publication = Open-Publication -PublisherApp $Publisher -Path $File.FullName

                if ($null -eq $Publication) {
                    throw "Publisher returned no document object."
                }

                Write-ConversionLog -Level "STEP" -Message "  Step 1 complete: publication opened."

                if ($ShowApplications) {
                    try { $Publisher.ActiveWindow.Visible = $true } catch { }
                }

                if ($OutputFormat -eq "PDF") {

                    # --- Step 2 - direct PDF export -------------------
                    Write-ConversionLog -Level "STEP" -Message "  Step 2: exporting to PDF..."

                    $Publication.ExportAsFixedFormat($PublisherPdfFormat, $OutputPath)

                    Write-ConversionLog -Level "STEP" -Message "  Step 2 complete: PDF exported."

                    $MethodUsed = "Publisher direct PDF export"

                    Write-ConversionLog -Level "STEP" -Message "  Step 3: closing publication..."
                    $Publication.Close()
                    Release-ComObject -ComObject $Publication
                    $Publication = $null
                    Write-ConversionLog -Level "STEP" -Message "  Step 3 complete."
                }
                else {

                    # --- Word via a PDF bridge ------------------------
                    if ($KeepIntermediatePdf) {
                        $TempPdfPath = [System.IO.Path]::ChangeExtension($File.FullName, ".pdf")
                    }
                    else {
                        $TempPdfPath = Join-Path -Path $env:TEMP `
                            -ChildPath ("PubConvert_" + [guid]::NewGuid().ToString() + ".pdf")
                    }

                    Write-ConversionLog -Level "STEP" -Message "  Step 2: exporting to intermediate PDF..."

                    $Publication.ExportAsFixedFormat($PublisherPdfFormat, $TempPdfPath)

                    Write-ConversionLog -Level "STEP" -Message "  Step 2 complete: PDF created."

                    Write-ConversionLog -Level "STEP" -Message "  Step 3: closing publication..."
                    $Publication.Close()
                    Release-ComObject -ComObject $Publication
                    $Publication = $null
                    Write-ConversionLog -Level "STEP" -Message "  Step 3 complete."

                    if (-not (Test-Path -LiteralPath $TempPdfPath -PathType Leaf)) {
                        throw "Publisher did not create the intermediate PDF."
                    }

                    Write-ConversionLog -Level "STEP" -Message "  Step 4: opening PDF in Word (this can take a moment)..."

                    $WordDocument = $Word.Documents.Open(
                        $TempPdfPath,
                        $false,
                        $false,
                        $false
                    )

                    if ($null -eq $WordDocument) {
                        throw "Word could not open the intermediate PDF."
                    }

                    Write-ConversionLog -Level "STEP" -Message "  Step 4 complete: PDF opened in Word."

                    Write-ConversionLog -Level "STEP" -Message "  Step 5: saving as .docx..."

                    $SaveMethod = Save-WordDocumentAs `
                        -Document $WordDocument `
                        -DestinationPath $OutputPath `
                        -Format $WordDocxFormat

                    Write-ConversionLog -Level "STEP" -Message "  Step 5 complete: .docx saved via $SaveMethod."

                    $WordDocument.Close($WordDoNotSaveChanges)
                    Release-ComObject -ComObject $WordDocument
                    $WordDocument = $null

                    $MethodUsed = "PDF bridge - saved via $SaveMethod"

                    if (-not $KeepIntermediatePdf) {
                        Remove-Item -LiteralPath $TempPdfPath -Force -ErrorAction SilentlyContinue
                    }

                    $TempPdfPath = $null
                }

                Stop-Watchdog -WatchdogJob $Watchdog
                $Watchdog = $null

                if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
                    throw "The conversion finished without creating the output file."
                }

                $CreatedFile = Get-Item -LiteralPath $OutputPath -ErrorAction Stop

                if ($CreatedFile.Length -le 0) {
                    throw "The output file was created but is empty."
                }

                $Converted = $true
                $SuccessCount++
                $FileStopwatch.Stop()

                Write-ConversionLog -Level "SUCCESS" -Message "Created: $OutputPath"

                $Results.Add([PSCustomObject]@{
                    SourcePUB       = $File.FullName
                    OutputFile      = $OutputPath
                    OutputFormat    = $OutputFormat
                    Status          = "Success"
                    Method          = $MethodUsed
                    AttemptCount    = $Attempt
                    DurationSeconds = [math]::Round($FileStopwatch.Elapsed.TotalSeconds, 2)
                    OutputSizeBytes = $CreatedFile.Length
                    Message         = "Output created and verified"
                })
            }
            catch {
                $LastErrorMessage = $_.Exception.Message

                Stop-Watchdog -WatchdogJob $Watchdog
                $Watchdog = $null

                $KilledByWatchdog = Test-WatchdogFired -MarkerPath $WatchdogMarker

                if ($KilledByWatchdog) {

                    $WasTimedOut = $true

                    Write-ConversionLog -Level "ERROR" `
                        -Message "TIMEOUT after $FileTimeoutSeconds seconds. Office was closed for: $($File.FullName)"

                    $LastErrorMessage = "Timed out after $FileTimeoutSeconds seconds. See the last Step line for the stage that blocked."

                    Release-ComObject -ComObject $WordDocument
                    Release-ComObject -ComObject $Publication
                    $WordDocument = $null
                    $Publication = $null

                    if (-not (Start-OfficeApplications)) {
                        Write-ConversionLog -Level "ERROR" -Message "Office could not be restarted. Stopping."
                        break
                    }
                }
                else {

                    Write-ConversionLog -Level "WARNING" `
                        -Message "Attempt $Attempt failed for '$($File.FullName)': $LastErrorMessage"

                    if ($null -ne $WordDocument) {
                        try { $WordDocument.Close($WordDoNotSaveChanges) } catch { }
                        Release-ComObject -ComObject $WordDocument
                        $WordDocument = $null
                    }

                    if ($null -ne $Publication) {
                        try { $Publication.Close() } catch { }
                        Release-ComObject -ComObject $Publication
                        $Publication = $null
                    }
                }

                if ((-not $Converted) -and ($Attempt -le $MaximumRetries)) {
                    Start-Sleep -Seconds $RetryDelaySeconds
                }
            }
        }

        if (-not $Converted) {

            $FileStopwatch.Stop()
            $FailedCount++

            if ($WasTimedOut) {
                $TimedOutCount++
                $FinalStatus = "Timed out"
            }
            else {
                $FinalStatus = "Failed"
            }

            Write-ConversionLog -Level "ERROR" -Message "$FinalStatus`: $($File.FullName)"

            $Results.Add([PSCustomObject]@{
                SourcePUB       = $File.FullName
                OutputFile      = $OutputPath
                OutputFormat    = $OutputFormat
                Status          = $FinalStatus
                Method          = ""
                AttemptCount    = $Attempt
                DurationSeconds = [math]::Round($FileStopwatch.Elapsed.TotalSeconds, 2)
                OutputSizeBytes = 0
                Message         = $LastErrorMessage
            })
        }
    }
}
finally {

    Write-Progress -Activity "Converting Publisher files to $OutputDescription" -Completed

    if ($null -ne $WordDocument) {
        try { $WordDocument.Close($WordDoNotSaveChanges) } catch { }
        Release-ComObject -ComObject $WordDocument
        $WordDocument = $null
    }

    if ($null -ne $Publication) {
        try { $Publication.Close() } catch { }
        Release-ComObject -ComObject $Publication
        $Publication = $null
    }

    if ($null -ne $Word) {
        try { $Word.Quit() } catch { }
        Release-ComObject -ComObject $Word
        $Word = $null
    }

    if ($null -ne $Publisher) {
        try { $Publisher.Quit() } catch { }
        Release-ComObject -ComObject $Publisher
        $Publisher = $null
    }

    Get-Job -ErrorAction SilentlyContinue | Remove-Job -Force -ErrorAction SilentlyContinue

    if (Test-Path -LiteralPath $WatchdogMarker -PathType Leaf) {
        Remove-Item -LiteralPath $WatchdogMarker -Force -ErrorAction SilentlyContinue
    }

    Restore-WordPdfWarning

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}


# =====================================================================
# SUMMARY
# =====================================================================

$OverallStopwatch.Stop()

try {
    $Results | Export-Csv -LiteralPath $CsvReport -NoTypeInformation -Encoding UTF8
    Write-ConversionLog -Level "SUCCESS" -Message "CSV report created: $CsvReport"
}
catch {
    Write-ConversionLog -Level "ERROR" -Message "CSV report could not be saved: $($_.Exception.Message)"
}

Write-ConversionLog -Message "Conversion finished."
Write-ConversionLog -Message "Found: $TotalFiles"
Write-ConversionLog -Message "Converted: $SuccessCount"
Write-ConversionLog -Message "Skipped: $SkippedCount"
Write-ConversionLog -Message "Failed: $FailedCount (timed out: $TimedOutCount)"
Write-ConversionLog -Message "Duration: $($OverallStopwatch.Elapsed)"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " CONVERSION COMPLETE" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Output format: $OutputDescription ($OutputExtension)"
Write-Host "Found:         $TotalFiles"
Write-Host "Converted:     $SuccessCount" -ForegroundColor Green
Write-Host "Skipped:       $SkippedCount" -ForegroundColor Yellow
Write-Host "Failed:        $FailedCount" -ForegroundColor Red

if ($TimedOutCount -gt 0) {
    Write-Host "  Timed out:   $TimedOutCount" -ForegroundColor Red
}

Write-Host "Duration:      $($OverallStopwatch.Elapsed)"
Write-Host ""
Write-Host "Log:    $LogFile"
Write-Host "Report: $CsvReport"

Exit-Script -ExitCode 0
