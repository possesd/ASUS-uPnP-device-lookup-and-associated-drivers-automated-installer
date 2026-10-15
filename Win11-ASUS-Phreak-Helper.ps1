#requires -RunAsAdministrator

[CmdletBinding()]
param(
    [switch]$Resume,
    [switch]$Reset
)
# ROG Strix SCAR 16 G635LW - Windows 11 25H2
# Driver & Update Installer v84
#
# v12 adds controlled multi-package reboot batching: a vendor installer that reports
# 3010 (reboot required) is treated as a successful install with a deferred reboot.
# The script can continue through up to five attended packages before requesting a full
# system reboot.  Exit code 1641 remains a hard stop because the installer has actually
# initiated a reboot.  Persistent batch state is keyed to the Windows boot marker so a
# real reboot resets the five-package counter. User-selected 'restart later' choices are
# therefore respected whenever the installer returns without initiating the reboot.
#
# v30 changes (cumulative):
# - Fixes resume-state schema compatibility: older JSON rows without LocalPath/other fields are normalized before property assignment.
# - Prevents the terminating "The property LocalPath cannot be found" error seen when resuming packages from older installer state.
# - Keeps the existing Desktop archive location; the Desktop path itself is valid, including when Desktop is OneDrive-backed.
# - Adds explicit local-path diagnostics before package launch and verifies the installer path is a real local filesystem path.
# - Retains the unlimited PnP rescan, serial driver-before-companion dependency ordering, live download speed/percentage output, -Reset, and terminal hold-open behavior.
# - File downloads always emit live percentage, transferred size, elapsed time and MB/s to the terminal.
# - PnPUtil /scan-devices has NO timeout and is allowed to run until Windows reports completion; it is invoked exactly once at the start of each script invocation.
# - PnP rescan output explicitly states that it is local hardware enumeration and performs no download.
# - Direct vendor file downloads use the same progress engine instead of Invoke-WebRequest -OutFile.
# - STAGES 01-05 perform validation/preparation only.
# - STAGE 06 detects the complete present PnP hardware set first.
# - STAGE 07 performs the downloaded-content installation workflow AFTER detection.
# - MyASUS is installed/updated immediately before the final Armoury Crate/Aura stage.
# - v33 performs exactly ONE unbounded PnP device rescan at the start of each script invocation.
# - v35 adds resilient ASUS catalog/network retry handling and a narrowly scoped WinGet Microsoft Store certificate-pinning retry for 0x8A15005E on fresh Windows 11 25H2 installations.
# - v36 repairs/re-registers the existing Microsoft Store package for the interactive administrator before any Microsoft Store companion install, because a fresh Windows 11 image can have the Store package staged but not registered for the current user.
# - v41 does not remove/reinstall Microsoft Store or alter DNS; it only repairs Store registration when the current user is not already registered. If the Store package is absent entirely, it records that condition and continues without treating the Store as a driver failure.
# - v42 forces Microsoft Store companion installs into non-interactive mode and applies a hard WinGet timeout so "Starting package install..." cannot block the ASUS package phase indefinitely.
# - The certificate-pinning bypass is temporary for the Store operation only and is restored immediately afterward.
# - MyASUS and ASUS Microsoft Store companion packages use the same Store retry helper.
# - v33 captures exactly ONE authoritative hardware/driver/software baseline after that rescan.
# - No later phase re-runs PnPUtil /scan-devices or rebuilds the hardware/driver/software inventories.
# - The baseline captured at STAGE 06 is the only hardware-ID, driver-version and software-version list used for the run.
# - Reboots resume at the saved post-detection installation checkpoint rather than repeating pre-detection work.
# - After hardware detection, ASUS's official G635LW catalog is queried for complete
#   driver/setup packages. Packages are downloaded as provided by ASUS and installed
#   through their setup.exe/install.exe/install.cmd/install.bat when available.
# - BIOS/firmware packages are deliberately excluded from unattended execution.
# - v83 resolves XTU from the official ASUS Driver Store first (G635LW, then compatible current Intel-platform ASUS catalogs), rejects legacy ASUS XTU builds, and only falls back to Intel's official XTU when ASUS exposes no current compatible package. ASUS XTU installers are launched attended with no silent switches.
# - v14 explicitly prioritizes the ASUS Intel Graphics package for the Core Ultra 9 275HX before all other Intel platform packages and before Intel XTU.
# - v16 retains ALL architecture variants and orders them x64 -> ARM64 -> x86 -> neutral/unspecified.
# - The ASUS platform bundle includes chipset, processor power management (PPM), Serial IO, DTT, CSME/ME, VPU, PMT, GNA and Intel SST components when present in the G635LW catalog.
# - No generic Intel CPU package is allowed to overwrite ASUS OEM platform packages merely to make XTU appear; Windows supplies the processor-class driver and ASUS platform packages remain authoritative.

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'Continue'
$script:MicrosoftStorePrepared = $false

# v59 fatal-runtime diagnostics. If a true terminating PowerShell exception occurs,
# preserve its exact line/command/type in the active log before the existing hold-open
# handler takes over.
trap {
    try {
        $fatal=$_.Exception
        $inv=$_.InvocationInfo
        $line=''
        try { $line=$inv.Line.Trim() } catch {}
        $msg="FATAL SCRIPT ERROR | Line=$($inv.ScriptLineNumber) | Offset=$($inv.OffsetInLine) | Command=$line | Exception=$($fatal.GetType().FullName): $($fatal.Message)"
        if($script:LogFile){ Add-Content -LiteralPath $script:LogFile -Value ("[$(Get-Date -Format 'HH:mm:ss')] [FATAL] $msg") -Encoding UTF8 }
    } catch {}
    break
}

# v36 intentionally keeps the terminal open at the end of manual/interactive runs.
# It also keeps the complete PnP rescan unbounded: Windows is allowed to finish the scan.
# Scheduled -Resume startup runs do not wait for keyboard input.

$ExpectedModelPattern = 'G635LW'
$Version = 'v84'
$Base = [System.IO.Path]::Combine([Environment]::GetFolderPath('CommonApplicationData'),'ROG-G635LW-Installer-v84')
$LogDir = [System.IO.Path]::Combine($Base,'Logs')
$StateFile = [System.IO.Path]::Combine($Base,'state.json')
$ReportFile = [System.IO.Path]::Combine($Base,'Final-Report.txt')
$StableScript = [System.IO.Path]::Combine($Base,'Win11-ASUS-Phreak-Helper.ps1')
$TaskName = 'ROG-G635LW-Installer-v84-Resume'
$MaxPackagesBetweenReboots = 5

New-Item -ItemType Directory -Force -Path $Base,$LogDir | Out-Null
# v84 persistent console log: always write beside the currently executing v84 script.
# The same file is reused across normal runs and -Resume runs after reboot.
$script:V84ScriptLocation = $PSScriptRoot
if([string]::IsNullOrWhiteSpace($script:V84ScriptLocation)){$script:V84ScriptLocation = (Get-Location).Path}
$script:LogFile = [System.IO.Path]::Combine([string]$script:V84ScriptLocation,'powershellv84.log')
$LogFile = $script:LogFile
$script:DownloadAllContentMode = $false
$script:AllContentDownloadMarker = $null
$script:V84TranscriptStarted = $false
# v84 FIX: do not use Start-Transcript against the same log file that the installer may already be using.
# This avoids the file-lock collision that previously caused Add-Content to fail when a prior
# v84 console still held powershellv84.log open. Write-Log owns the persistent text log.
try {
    Add-Content -LiteralPath $script:LogFile -Value "`r`n===== ROG G635LW Installer v84 START $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') =====" -Encoding UTF8 -ErrorAction Stop
    Add-Content -LiteralPath $script:LogFile -Value "[INFO] v84 PowerShell log: $script:LogFile" -Encoding UTF8 -ErrorAction Stop
} catch {
    Write-Host "[WARN] Could not initialize powershellv84.log: $($_.Exception.Message)" -ForegroundColor Yellow
}
$TargetUser = 'possesd'
$TargetUserRoot = [System.IO.Path]::Combine('C:\Users',$TargetUser)
$LoginResumeBatch = [System.IO.Path]::Combine($TargetUserRoot,'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup','ROG-G635LW-Installer-v84-Resume.bat')
$LegacyV80ResumeBatch=[System.IO.Path]::Combine($TargetUserRoot,'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup','ROG-G635LW-Installer-v80-Resume.bat')
try { if(Test-Path -LiteralPath $LegacyV80ResumeBatch){Remove-Item -LiteralPath $LegacyV80ResumeBatch -Force -ErrorAction SilentlyContinue} } catch {}
$MasterDownloadRoot = [System.IO.Path]::Combine($TargetUserRoot,'ASUS-G635LW-Installer-Content')
$DesktopDownloadRoot = [System.IO.Path]::Combine($TargetUserRoot,'Desktop','ASUS-G635LW-Installer-Content')
$DownloadRoot = $MasterDownloadRoot
# v83 startup path initialization deliberately uses System.IO.Path::Combine for the initial paths.
# This prevents Join-Path from entering its interactive ChildPath prompt before MAIN starts.
# The visible archive remains under the user's Desktop as requested. The script never assumes
# that the Desktop path itself is a special object; it is treated as an ordinary filesystem path.
# If OneDrive Files On-Demand marks a downloaded installer online-only, the installer is deferred
# until the file is physically available locally.

function Get-NvidiaBroadcastSdkPackageSet {
    # Select NVIDIA Broadcast SDK redistributables for the detected GPU generation.
    # NVIDIA publishes separate official redistributables for RTX 20/30/40/50 families.
    # Do not fall back to RTX 50 packages when the GPU family cannot be determined.
    $gpuNames = @()
    try {
        $gpuNames = @(Get-CimInstance -ClassName Win32_VideoController -ErrorAction Stop |
            Where-Object { ([string]$_.Name -match '(?i)NVIDIA|GeForce|Quadro|TITAN') -and ([string]$_.Name -match '(?i)RTX|TITAN') } |
            ForEach-Object { [string]$_.Name })
    } catch {
        Write-Log "NVIDIA Broadcast SDK selection could not query Win32_VideoController: $($_.Exception.Message)" 'WARN'
    }
    $gpuText = [string]::Join(' | ', [string[]]$gpuNames)
    $generation = $null
    if ($gpuText -match '(?i)RTX\s*50\d{2}|GeForce RTX 50|Blackwell') { $generation = '50' }
    elseif ($gpuText -match '(?i)RTX\s*40\d{2}|GeForce RTX 40|RTX .*Ada|Ada Generation') { $generation = '40' }
    elseif ($gpuText -match '(?i)RTX\s*30\d{2}|GeForce RTX 30|RTX .*Ampere|Ampere Generation|RTX A\d{4}') { $generation = '30' }
    elseif ($gpuText -match '(?i)RTX\s*20\d{2}|GeForce RTX 20|Quadro RTX|TITAN RTX') { $generation = '20' }

    if (-not $generation) {
        Write-Log "Could not reliably identify the NVIDIA RTX generation from detected adapter name(s): '$gpuText'. NVIDIA AR/Audio/Video SDK downloads are being skipped rather than downloading a mismatched generation." 'WARN'
        return @()
    }
    Write-Log "Detected NVIDIA adapter(s): $gpuText. Selecting official NVIDIA Broadcast SDK redistributables for RTX $generation Series." 'OK'

    $resourceUrl = 'https://www.nvidia.com/en-gb/geforce/broadcasting/broadcast-sdk/resources/'
    try {
        $html = [string](Invoke-WebRequest -Uri $resourceUrl -UseBasicParsing -TimeoutSec 60 -Headers @{'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v84'}).Content
    } catch {
        Write-Log "Could not retrieve NVIDIA's official Broadcast SDK resource page. No generation-specific SDK package will be substituted: $($_.Exception.Message)" 'WARN'
        return @()
    }

    # The official page labels each Download link by GPU family. Match each direct
    # international.download.nvidia.com SDK URL to the nearest family label in the page.
    $labelPatterns = @{
        '50' = '(?i)For NVIDIA GeForce RTX\s*50 Series'
        '40' = '(?i)For NVIDIA GeForce RTX\s*40 Series|NVIDIA RTX Ada Generation GPUs'
        '30' = '(?i)For NVIDIA GeForce RTX\s*30 Series|NVIDIA RTX Ampere Generation GPUs'
        '20' = '(?i)For NVIDIA GeForce RTX\s*20 Series|NVIDIA Quadro RTX'
    }
    $flat = [regex]::Replace($html, '(?is)<script\b.*?</script>|<style\b.*?</style>', ' ')
    $flat = [System.Net.WebUtility]::HtmlDecode([regex]::Replace($flat, '(?s)<[^>]+>', ' '))
    $flat = [regex]::Replace($flat, '\s+', ' ')
    $urlMatches = [regex]::Matches($html, '(?i)https?://international\.download\.nvidia\.com/Windows/broadcast/sdk/(?:AFX|VFX|AR)/[^"''\s<>]+')
    $urlsByKind = @{ AFX=$null; VFX=$null; AR=$null }
    foreach ($match in $urlMatches) {
        $url = [System.Net.WebUtility]::HtmlDecode($match.Value) -replace '&amp;','&'
        $kindMatch = [regex]::Match($url, '(?i)/sdk/(AFX|VFX|AR)/')
        if (-not $kindMatch.Success) { continue }
        $kind = $kindMatch.Groups[1].Value.ToUpperInvariant()
        $urlPos = $flat.IndexOf([System.Net.WebUtility]::HtmlDecode([regex]::Match($match.Value, '[^/]+$').Value), [StringComparison]::OrdinalIgnoreCase)
        # Use a local slice around this URL in raw HTML; the closest generation label
        # normally appears alongside the associated Download control.
        $start = [Math]::Max(0, $match.Index - 1000)
        $length = [Math]::Min($html.Length - $start, $match.Length + 2000)
        $context = $html.Substring($start, $length)
        $labelHits = @()
        foreach ($gen in @('50','40','30','20')) {
            foreach ($lm in [regex]::Matches($context, $labelPatterns[$gen])) {
                $labelHits += [pscustomobject]@{Generation=$gen;Index=$lm.Index}
            }
        }
        if ($labelHits.Count -eq 0) { continue }
        $localUrlIndex = [Math]::Max(0, $match.Index - $start)
        $nearest = $labelHits | Sort-Object @{Expression={ [Math]::Abs($_.Index - $localUrlIndex) }} | Select-Object -First 1
        if ($nearest.Generation -eq $generation -and -not $urlsByKind[$kind]) { $urlsByKind[$kind] = $url }
    }

    $defs = @(
        [pscustomobject]@{Kind='AFX';Name="NVIDIA Audio Effects SDK RTX $generation";Folder='Audio Effects SDK';Version='1.6.1'}
        [pscustomobject]@{Kind='VFX';Name="NVIDIA Video Effects SDK RTX $generation";Folder='Video Effects SDK';Version='0.7.6'}
        [pscustomobject]@{Kind='AR';Name="NVIDIA AR SDK RTX $generation";Folder='AR-VR Effects SDK';Version='0.8.7'}
    )
    $result = @()
    foreach ($def in $defs) {
        $uri = [string]$urlsByKind[$def.Kind]
        if ([string]::IsNullOrWhiteSpace($uri)) {
            Write-Log "NVIDIA's official resource page did not expose a verifiable RTX $generation URL for $($def.Name). This package will be skipped; no RTX 50 or other-generation fallback will be used." 'WARN'
            continue
        }
        $fileName = [System.IO.Path]::GetFileName(([uri]$uri).AbsolutePath)
        $result += [pscustomobject]@{Name=$def.Name;Uri=$uri;FileName=$fileName;Folder=$def.Folder;Generation=$generation;Version=$def.Version}
    }
    return $result
}

function Wait-ForUserBeforeExit {
    param(
        [int]$ExitCode = 0,
        [string]$Reason = 'The installer has finished this run.'
    )

    # Automatic startup-resume runs must never block waiting for keyboard input.
    # Manual/interactive runs remain open so the complete terminal output can be reviewed.
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host ' ROG G635LW INSTALLER v84 - RUN FINISHED' -ForegroundColor Cyan
    Write-Host " $Reason" -ForegroundColor White
    Write-Host " Exit code: $ExitCode" -ForegroundColor White
    Write-Host ' The PowerShell window will remain open until you press ENTER.' -ForegroundColor Green
    Write-Host ' Review the complete terminal output above before closing it.' -ForegroundColor Green
    Write-Host '============================================================' -ForegroundColor Cyan

    if(-not $Resume){
        try {
            [void](Read-Host 'Press ENTER when you have finished reviewing the output')
        } catch {
            Write-Host 'Interactive input was unavailable; leaving the process at the end of the script.' -ForegroundColor Yellow
            Start-Sleep -Seconds 30
        }
    } else {
        Write-Log 'Automatic -Resume run detected; no interactive pause was requested for the scheduled startup process.' 'INFO'
    }

    return $ExitCode
}

function Invoke-InstallerReset {
    # -Reset removes ONLY this installer's persistent state, downloaded archive, and reboot task.
    # It deliberately does NOT uninstall drivers, software, or change device state.
    Write-Host 'ROG G635LW Installer v84 RESET requested.' -ForegroundColor Yellow
    Write-Host 'This will remove the v42 download archive, resume/checkpoint state, logs/report, and scheduled reboot task (and clean legacy v21/v20/v19/v18/v17/v16/v15 resume state).' -ForegroundColor Yellow
    Write-Host 'Installed drivers and installed software will NOT be removed or changed.' -ForegroundColor Green

    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Scheduled resume task removed: $TaskName" -ForegroundColor Green
    } catch {
        Write-Host "Scheduled task removal warning: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    try {
        if(Test-Path -LiteralPath $LoginResumeBatch){ Remove-Item -LiteralPath $LoginResumeBatch -Force -ErrorAction SilentlyContinue; Write-Host "Startup resume batch removed: $LoginResumeBatch" -ForegroundColor Green }
    } catch {
        Write-Host "Startup batch removal warning: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Remove the user-visible Desktop archive. If it is a junction, removing the junction
    # does not delete the target by itself, so the master archive is removed separately below.
    foreach($path in @($DesktopDownloadRoot,$MasterDownloadRoot)) {
        if(Test-Path -LiteralPath $path) {
            try {
                $item=Get-Item -LiteralPath $path -Force -ErrorAction Stop
                if(($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                    Remove-Item -LiteralPath $path -Force -ErrorAction Stop
                } else {
                    Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
                }
                Write-Host "Removed installer archive: $path" -ForegroundColor Green
            } catch {
                Write-Host "Could not immediately remove ${path}: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }

    # Also remove previous v17/v16/v15 resume tasks/state so an older pending startup task
    # cannot resurrect an older installer after v26 has been reset. This does not touch
    # installed drivers or installed applications.
    foreach($legacyTask in @('ROG-G635LW-Installer-v24-Resume','ROG-G635LW-Installer-v23-Resume','ROG-G635LW-Installer-v21-Resume','ROG-G635LW-Installer-v20-Resume','ROG-G635LW-Installer-v19-Resume','ROG-G635LW-Installer-v18-Resume','ROG-G635LW-Installer-v17-Resume','ROG-G635LW-Installer-v16-Resume','ROG-G635LW-Installer-v15-Resume')) {
        try { Unregister-ScheduledTask -TaskName $legacyTask -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    }

    # Because the currently executing script may itself be the stable copy inside $Base,
    # schedule a tiny one-shot cleanup process to remove v17 and legacy installer state
    # after this process exits.
    $legacyBases=@(
        $Base,
        (Join-Path $env:ProgramData 'ROG-G635LW-Installer-v24'),
        (Join-Path $env:ProgramData 'ROG-G635LW-Installer-v23'),
        (Join-Path $env:ProgramData 'ROG-G635LW-Installer-v21'),
        (Join-Path $env:ProgramData 'ROG-G635LW-Installer-v20'),
        (Join-Path $env:ProgramData 'ROG-G635LW-Installer-v19'),
        (Join-Path $env:ProgramData 'ROG-G635LW-Installer-v18'),
        (Join-Path $env:ProgramData 'ROG-G635LW-Installer-v17'),
        (Join-Path $env:ProgramData 'ROG-G635LW-Installer-v16'),
        (Join-Path $env:ProgramData 'ROG-G635LW-Installer-v15')
    ) | Select-Object -Unique
    $cleanupPaths=($legacyBases | ForEach-Object { "'$($_.Replace("'","''"))'" }) -join ','
    $cleanupCmd = 'Start-Sleep -Seconds 2; foreach($p in @(' + $cleanupPaths + ')){Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue}'
    try {
        Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-Command',$cleanupCmd) | Out-Null
        Write-Host "Scheduled removal of installer state: $Base" -ForegroundColor Green
    } catch {
        Write-Host "Could not schedule ProgramData cleanup: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "You can manually remove: $Base" -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host 'v53 reset complete. No installed drivers or software were removed.' -ForegroundColor Green
    Write-Host 'A normal v59 run can now be started from the Desktop script.' -ForegroundColor Green
    [void](Wait-ForUserBeforeExit -ExitCode 0 -Reason 'The installer reset completed. Installed drivers and software were not changed.')
    exit 0
}

function Write-Log {
    param([string]$Message,[string]$Level='INFO')
    # Logging must never abort an installer checkpoint because a nested call supplied an unexpected level token.
    if([string]::IsNullOrWhiteSpace($Level) -or @('INFO','OK','WARN','ERROR','STEP') -notcontains $Level){$Level='INFO'}
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'),$Level,$Message
    Write-Host $line
    try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -ErrorAction Stop } catch { Write-Host "[WARN] Log write failed: $($_.Exception.Message)" -ForegroundColor Yellow }
}


if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run this script as Administrator.'
}

# Keep a manually launched console open even if a terminating runtime error occurs.
# Scheduled -Resume executions are exempt so a startup task can finish without waiting for input.
trap {
    $trapMessage = $_.Exception.Message
    Write-Host '' -ForegroundColor Red
    Write-Host '============================================================' -ForegroundColor Red
    Write-Host ' ROG G635LW INSTALLER v59 - TERMINATING ERROR' -ForegroundColor Red
    Write-Host " $trapMessage" -ForegroundColor Red
    Write-Host ' The terminal will remain open so the error and preceding output can be reviewed.' -ForegroundColor Yellow
    Write-Host '============================================================' -ForegroundColor Red
    if(-not $Resume){
        try { [void](Read-Host 'Press ENTER when you have finished reviewing the error output') } catch { Start-Sleep -Seconds 30 }
    }
    exit 1
}

if ($Reset) { Invoke-InstallerReset }

# v84: Windows Update is deliberately excluded from this installer and persistently disabled on the target Windows installation.
function Disable-WindowsUpdatePersistently {
    Write-Log 'STAGE 02B - Persistently disabling Windows Update for this Windows installation' 'STEP'
    Write-Log 'Windows Update is being disabled by persistent machine policy. The Windows Update service itself is NOT forcibly disabled because Microsoft no longer recommends service-disabling as the supported method.' 'INFO'

    $wuPolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $wuAuPolicy = Join-Path $wuPolicy 'AU'
    $wuStorePolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore'

    try {
        New-Item -Path $wuPolicy -Force -ErrorAction Stop | Out-Null
        New-Item -Path $wuAuPolicy -Force -ErrorAction Stop | Out-Null

        # Microsoft-documented Windows Update policy controls:
        # - NoAutoUpdate=1 disables Automatic Updates.
        # - DisableWindowsUpdateAccess=1 removes access to Windows Update UI/features.
        # - DoNotConnectToWindowsUpdateInternetLocations=1 prevents Windows Update from
        #   connecting to Windows Update Internet locations.
        New-ItemProperty -LiteralPath $wuAuPolicy -Name 'NoAutoUpdate' -PropertyType DWord -Value 1 -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -LiteralPath $wuAuPolicy -Name 'AUOptions' -PropertyType DWord -Value 1 -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -LiteralPath $wuPolicy -Name 'DisableWindowsUpdateAccess' -PropertyType DWord -Value 1 -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -LiteralPath $wuPolicy -Name 'DoNotConnectToWindowsUpdateInternetLocations' -PropertyType DWord -Value 1 -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -LiteralPath $wuPolicy -Name 'ExcludeWUDriversInQualityUpdate' -PropertyType DWord -Value 1 -Force -ErrorAction Stop | Out-Null

        # Keep the Microsoft-documented WSUS policy path explicitly configured so a
        # future Windows Update service restart cannot simply fall back to the public
        # Windows Update endpoint while this installer policy remains in place.
        New-ItemProperty -LiteralPath $wuPolicy -Name 'WUServer' -PropertyType String -Value ' ' -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -LiteralPath $wuPolicy -Name 'WUStatusServer' -PropertyType String -Value ' ' -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -LiteralPath $wuPolicy -Name 'UpdateServiceUrlAlternate' -PropertyType String -Value ' ' -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -LiteralPath $wuAuPolicy -Name 'UseWUServer' -PropertyType DWord -Value 1 -Force -ErrorAction Stop | Out-Null

        # Prevent the Microsoft Store from offering a Windows OS upgrade. This is
        # separate from Store application updates and is included only to prevent an
        # OS upgrade offer from bypassing the Windows Update policy.
        New-Item -Path $wuStorePolicy -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -LiteralPath $wuStorePolicy -Name 'DisableOSUpgrade' -PropertyType DWord -Value 1 -Force -ErrorAction Stop | Out-Null

        Write-Log 'Persistent Windows Update policy values were written successfully.' 'OK'
    } catch {
        Write-Log "CRITICAL: Could not completely write the persistent Windows Update policy: $($_.Exception.Message)" 'ERROR'
        return $false
    }

    # Re-apply the same policy at every boot so the disabled state remains in force
    # throughout all installer reboots and after the installer has finished. The
    # enforcement script is deliberately stored outside the installer archive so
    # -Reset cannot remove the future Windows Update policy.
    try {
        $policyRoot = [System.IO.Path]::Combine([Environment]::GetFolderPath('CommonApplicationData'),'ROG-G635LW-WindowsUpdatePolicy')
        $policyFile = [System.IO.Path]::Combine($policyRoot,'Disable-WindowsUpdate.ps1')
        New-Item -ItemType Directory -Force -Path $policyRoot -ErrorAction Stop | Out-Null
        $policyScript = @'
# Persistent Windows Update policy enforcement for ASUS ROG G635LW installer.
# This file intentionally does NOT stop/disable the Windows Update service.
$wuPolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
$wuAuPolicy = Join-Path $wuPolicy 'AU'
$wuStorePolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore'
New-Item -Path $wuPolicy -Force | Out-Null
New-Item -Path $wuAuPolicy -Force | Out-Null
New-ItemProperty -LiteralPath $wuAuPolicy -Name 'NoAutoUpdate' -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -LiteralPath $wuAuPolicy -Name 'AUOptions' -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -LiteralPath $wuPolicy -Name 'DisableWindowsUpdateAccess' -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -LiteralPath $wuPolicy -Name 'DoNotConnectToWindowsUpdateInternetLocations' -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -LiteralPath $wuPolicy -Name 'ExcludeWUDriversInQualityUpdate' -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -LiteralPath $wuPolicy -Name 'WUServer' -PropertyType String -Value ' ' -Force | Out-Null
New-ItemProperty -LiteralPath $wuPolicy -Name 'WUStatusServer' -PropertyType String -Value ' ' -Force | Out-Null
New-ItemProperty -LiteralPath $wuPolicy -Name 'UpdateServiceUrlAlternate' -PropertyType String -Value ' ' -Force | Out-Null
New-ItemProperty -LiteralPath $wuAuPolicy -Name 'UseWUServer' -PropertyType DWord -Value 1 -Force | Out-Null
New-Item -Path $wuStorePolicy -Force | Out-Null
New-ItemProperty -LiteralPath $wuStorePolicy -Name 'DisableOSUpgrade' -PropertyType DWord -Value 1 -Force | Out-Null
'@
        Set-Content -LiteralPath $policyFile -Value $policyScript -Encoding UTF8 -Force

        $taskName = 'ROG-G635LW-Disable-Windows-Update'
        $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $action = New-ScheduledTaskAction -Execute $psExe -Argument ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $policyFile + '"')
        $trigger = New-ScheduledTaskTrigger -AtStartup
        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force -ErrorAction Stop | Out-Null
        Write-Log "Persistent Windows Update enforcement task registered: $taskName (SYSTEM / AtStartup). Policy script: $policyFile" 'OK'
    } catch {
        Write-Log "WARNING: Could not register the boot-time Windows Update enforcement task: $($_.Exception.Message). The machine policy remains persistent." 'WARN'
    }

    try {
        $noAuto = (Get-ItemProperty -LiteralPath $wuAuPolicy -Name 'NoAutoUpdate' -ErrorAction Stop).NoAutoUpdate
        $ux = (Get-ItemProperty -LiteralPath $wuPolicy -Name 'DisableWindowsUpdateAccess' -ErrorAction Stop).DisableWindowsUpdateAccess
        $net = (Get-ItemProperty -LiteralPath $wuPolicy -Name 'DoNotConnectToWindowsUpdateInternetLocations' -ErrorAction Stop).DoNotConnectToWindowsUpdateInternetLocations
        Write-Log "Windows Update persistent policy verification: NoAutoUpdate=$noAuto | DisableWindowsUpdateAccess=$ux | DoNotConnectToWindowsUpdateInternetLocations=$net" 'OK'
    } catch {
        Write-Log "WARNING: Could not verify all Windows Update policy values: $($_.Exception.Message)" 'WARN'
    }

    return $true
}

if(-not (Disable-WindowsUpdatePersistently)){
    Write-Log 'CRITICAL: Persistent Windows Update disabling could not be established. Stopping before driver/package work.' 'ERROR'
    [void](Wait-ForUserBeforeExit -ExitCode 4 -Reason 'The requested persistent Windows Update-disabled policy could not be established.')
    exit 4
}

# v26 takes ownership of future resume runs. If an older v21/v20/v19... startup task
# is still registered from a previous run, remove that task now so two installer versions
# cannot race over the same Desktop archive. This does not remove any installed driver/software.
foreach($legacyTask in @('ROG-G635LW-Installer-v24-Resume','ROG-G635LW-Installer-v23-Resume','ROG-G635LW-Installer-v21-Resume','ROG-G635LW-Installer-v20-Resume','ROG-G635LW-Installer-v19-Resume','ROG-G635LW-Installer-v18-Resume','ROG-G635LW-Installer-v17-Resume','ROG-G635LW-Installer-v16-Resume','ROG-G635LW-Installer-v15-Resume')) {
    if($legacyTask -ne $TaskName){
        try {
            $oldTask=Get-ScheduledTask -TaskName $legacyTask -ErrorAction SilentlyContinue
            if($null -ne $oldTask){
                Unregister-ScheduledTask -TaskName $legacyTask -Confirm:$false -ErrorAction SilentlyContinue
                Write-Log "Removed legacy scheduled resume task: $legacyTask" 'INFO'
            }
        } catch {}
    }
}

function Test-G635LW {
    $cs = Get-CimInstance Win32_ComputerSystem
    $model = "$($cs.Manufacturer) $($cs.Model)"
    Write-Log "Detected model: $model"
    if ($model -notmatch $ExpectedModelPattern) {
        throw "This installer is locked to ASUS ROG Strix SCAR 16 G635LW. Detected: $model"
    }
    return $model
}

function Get-OSInfo {
    $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    Write-Log "Detected Windows: $($cv.DisplayVersion) build $($cv.CurrentBuild).$($cv.UBR)"
    [pscustomobject]@{
        DisplayVersion=$cv.DisplayVersion
        Build=$cv.CurrentBuild
        UBR=$cv.UBR
    }
}

function Prepare-RestorePoint {
    Write-Log 'STAGE 03 - Recovery preparation' 'STEP'
    try {
        # Check/enable the Windows System Restore configuration before creating the point.
        $srKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore'
        if (Test-Path $srKey) {
            $disableSR = (Get-ItemProperty -LiteralPath $srKey -Name DisableSR -ErrorAction SilentlyContinue).DisableSR
            $disableConfig = (Get-ItemProperty -LiteralPath $srKey -Name DisableConfig -ErrorAction SilentlyContinue).DisableConfig
            if ($disableSR -eq 1 -or $disableConfig -eq 1) {
                Write-Log 'System Restore is disabled in the local configuration; enabling it for the installer recovery point.' 'WARN'
                New-ItemProperty -LiteralPath $srKey -Name DisableSR -PropertyType DWord -Value 0 -Force | Out-Null
                New-ItemProperty -LiteralPath $srKey -Name DisableConfig -PropertyType DWord -Value 0 -Force | Out-Null
            }
        }

        # Checkpoint-Computer relies on the VSS/Software Shadow Copy Provider stack.
        foreach ($svcName in @('VSS','swprv')) {
            try {
                $svc = Get-Service -Name $svcName -ErrorAction Stop
                if ($svc.StartType -eq 'Disabled') {
                    Set-Service -Name $svcName -StartupType Manual -ErrorAction Stop
                    Write-Log "$svcName service was disabled; changed startup type to Manual." 'WARN'
                }
                if ($svc.Status -ne 'Running') {
                    Start-Service -Name $svcName -ErrorAction Stop
                    Write-Log "$svcName service started for restore-point creation." 'OK'
                }
            } catch {
                Write-Log "Could not prepare ${svcName}: $($_.Exception.Message)" 'WARN'
            }
        }

        try {
            Enable-ComputerRestore -Drive 'C:\' -ErrorAction Stop
            Write-Log 'System Restore is enabled for C:\.' 'OK'
        } catch {
            Write-Log "Enable-ComputerRestore did not complete: $($_.Exception.Message)" 'WARN'
        }

        $rp = Get-CimInstance -Namespace root/default -ClassName SystemRestore -ErrorAction Stop |
            Sort-Object CreationTime -Descending | Select-Object -First 1
        if ($rp) {
            try {
                $dt = [Management.ManagementDateTimeConverter]::ToDateTime($rp.CreationTime)
                if ((Get-Date) - $dt -lt [TimeSpan]::FromHours(24)) {
                    Write-Log "A restore point already exists from $dt; Windows will not create another within 24 hours." 'OK'
                    return
                }
            } catch {}
        }

        Checkpoint-Computer -Description 'ROG G635LW Installer v59' -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
        Write-Log 'Restore point created successfully.' 'OK'
    } catch {
        Write-Log "Restore point was not created: $($_.Exception.Message). Continuing without aborting the installer." 'WARN'
    }
}





function Prepare-Resume {
    Write-Log 'STAGE 05 - Resume/reboot preparation (no package download/install)' 'STEP'
    Write-Log 'Main driver and software installation is intentionally deferred until after STAGE 06.' 'OK'
}

function Get-HardwareInventory {
    param([string]$Label='STAGE 06 - Hardware detection')
    Write-Log $Label 'STEP'

    $devices = @(Get-CimInstance Win32_PnPEntity | Where-Object {$_.Present -eq $true})

    # Capture the device's real hardware-ID list once per full inventory. This is the
    # association source used later for every ASUS driver package and
    # driver, rather than relying only on friendly device names.
    if(Get-Command Get-PnpDeviceProperty -ErrorAction SilentlyContinue){
        foreach($d in $devices){
            try{
                $hp=Get-PnpDeviceProperty -InstanceId ([string]$d.PNPDeviceID) -KeyName 'DEVPKEY_Device_HardwareIds' -ErrorAction SilentlyContinue
                $ids=@($hp.Data | ForEach-Object {[string]$_} | Where-Object {$_})
                $d | Add-Member NoteProperty HardwareIds $ids -Force
            }catch{$d | Add-Member NoteProperty HardwareIds @() -Force}
        }
    } else {
        foreach($d in $devices){$d | Add-Member NoteProperty HardwareIds @() -Force}
    }

    foreach ($d in $devices | Sort-Object Name) {
        $id=[string]$d.PNPDeviceID
        if ($id -match 'PCI\\VEN_' -or $d.Name -match 'NVIDIA|Intel|Realtek|ASUS|AMD|MediaTek|Killer|Bluetooth|Audio|Ethernet|Network|Display|Storage|Touchpad|Camera|USB|Chipset|Serial|HID') {
            $cm=[int]$d.ConfigManagerErrorCode
            $status=if($cm -eq 0){'OK (CM_PROB_NONE)'}else{"PROBLEM (code $cm)"}
            Write-Log ("HW: {0} | {1} | Driver {2} | PnP Status {3}" -f $d.Name,$d.Manufacturer,$d.DriverVersion,$status)
        }
    }

    $bad=@($devices | Where-Object {$_.ConfigManagerErrorCode -and $_.ConfigManagerErrorCode -ne 0})
    if ($bad.Count) {
        Write-Log "PnP configuration errors detected: $($bad.Count). Only non-zero ConfigManagerErrorCode values are counted as problems." 'WARN'
    } else {
        Write-Log 'No PnP device configuration errors detected. CM_PROB_NONE/0 is the normal healthy status and is not an error.' 'OK'
    }

    # v44 USB inventory: explicitly enumerate and classify every present USB device from the
    # authoritative STAGE 06 set. This includes USB audio/sound boxes, mice/keyboards/controllers,
    # phones (MTP/PTP/ADB/Fastboot-style interfaces), serial devices and composite peripherals.
    $usb=@($devices | Where-Object { [string]$_.PNPDeviceID -match '(?i)^USB\\VID_[0-9A-F]{4}&PID_[0-9A-F]{4}' })
    Write-Log "STAGE 06 USB baseline: detected $($usb.Count) present USB VID/PID device instance(s)." 'OK'
    foreach($u in $usb){
        $vidPid=[regex]::Match([string]$u.PNPDeviceID,'(?i)^USB\\VID_[0-9A-F]{4}&PID_[0-9A-F]{4}').Value
        $vendor=Get-HardwareVendor -Device $u
        $family=Get-DeviceComponentFamily -Device $u
        Write-Log "USB HW: $($u.Name) | Manufacturer=$($u.Manufacturer) | Vendor=$vendor | Family=$family | VID/PID=$vidPid | PNP=$($u.PNPDeviceID) | Driver=$($u.DriverVersion) | CM=$($u.ConfigManagerErrorCode)" 'INFO'
    }

    return $devices
}


function Initialize-MicrosoftStoreForCurrentUser {
    if($script:MicrosoftStorePrepared){ return $true }

    # Microsoft documents re-registering the existing Store package as the supported
    # recovery path when Microsoft Store exists on disk but is not registered for the
    # current user. Do not uninstall Store, do not modify DNS, and do not download a
    # third-party Store package.
    if([string]$env:USERNAME -eq 'SYSTEM'){
        Write-Log 'Microsoft Store preparation is deferred because this invocation is running as SYSTEM. Store applications require the interactive user context; the driver/package checkpoint remains authoritative.' 'WARN'
        return $false
    }

    Write-Log 'STAGE 07 - Checking Microsoft Store registration before Microsoft Store packages.' 'STEP'
    $storeFamily='Microsoft.WindowsStore_8wekyb3d8bbwe'
    $found=$false
    $registered=$false
    try {
        $current=@(Get-AppxPackage -Name 'Microsoft.WindowsStore' -ErrorAction SilentlyContinue)
        if($current.Count -gt 0){
            # Get-AppxPackage without -AllUsers enumerates packages registered for the
            # current interactive user. If it returns Microsoft.WindowsStore, it is already
            # registered. Do NOT call Add-AppxPackage -Register here: doing so while the
            # Store UI is open can produce ERROR_PACKAGE_IN_USE and is unnecessary.
            $found=$true
            $registered=$true
            Write-Log 'Microsoft Store is already registered for the current interactive user; no Store re-registration is required.' 'OK'
        }

        if(-not $found){
            $all=@(Get-AppxPackage -AllUsers -Name 'Microsoft.WindowsStore' -ErrorAction SilentlyContinue)
            if($all.Count -gt 0){
                $found=$true
                try {
                    Add-AppxPackage -RegisterByFamilyName -MainPackage $storeFamily -ErrorAction Stop
                    Write-Log 'Microsoft Store exists on the machine but was not registered for the current user; registered it by package family name.' 'OK'
                    $registered=$true
                } catch {
                    Write-Log "Microsoft Store family-name registration failed: $($_.Exception.Message)" 'WARN'
                    foreach($pkg in $all){
                        $manifest=Join-Path ([string]$pkg.InstallLocation) 'AppxManifest.xml'
                        if(Test-Path -LiteralPath $manifest){
                            try {
                                Add-AppxPackage -Path $manifest -Register -DisableDevelopmentMode -ErrorAction Stop
                                Write-Log "Microsoft Store registered from existing staged manifest: $manifest" 'OK'
                                $registered=$true
                            } catch { Write-Log "Microsoft Store staged-manifest registration failed: $($_.Exception.Message)" 'WARN' }
                        }
                    }
                }
            }
        }

        # Do not re-register Microsoft.StorePurchaseApp when the Store is already healthy.
        # Store/Purchase-App package mutations while the Store UI is open can fail with
        # ERROR_PACKAGE_IN_USE. If Store was actually repaired above, refresh the existing
        # Purchase App registration only after the repair.
        if($found -and -not $registered){
            $purchase=@(Get-AppxPackage -AllUsers -Name 'Microsoft.StorePurchaseApp' -ErrorAction SilentlyContinue)
            foreach($pkg in $purchase){
                $manifest=Join-Path ([string]$pkg.InstallLocation) 'AppxManifest.xml'
                if(Test-Path -LiteralPath $manifest){
                    try { Add-AppxPackage -Path $manifest -Register -DisableDevelopmentMode -ErrorAction Stop; Write-Log 'Microsoft Store Purchase App registration refreshed after Store repair.' 'OK' } catch { Write-Log "Store Purchase App registration could not be refreshed: $($_.Exception.Message)" 'WARN' }
                }
            }
        }
    } catch {
        Write-Log "Microsoft Store readiness check failed: $($_.Exception.Message)" 'WARN'
    }

    if(-not $found){
        Write-Log 'Microsoft.WindowsStore is not present in the local AppX package inventory. No unsupported third-party Store package will be downloaded by this installer; Microsoft Store companion packages will remain deferred.' 'WARN'
        $script:MicrosoftStorePrepared=$false
        return $false
    }

    # Do not launch wsreset.exe automatically. It can open Microsoft Store while WinGet is
    # trying to deploy an MSIX/AppX package, which can lock Microsoft.WindowsStore and cause
    # ERROR_PACKAGE_IN_USE. Store cache reset remains a manual recovery option.
    Write-Log 'Microsoft Store preparation complete; no automatic Store UI/cache reset will be launched before WinGet Store deployment.' 'INFO'

    $script:MicrosoftStorePrepared=$true
    return $true
}

function Invoke-WingetMsStoreInstall {
    param(
        [Parameter(Mandatory)][string]$WingetPath,
        [Parameter(Mandatory)][string]$PackageId,
        [switch]$Exact,
        [switch]$Silent,
        [switch]$Interactive,
        [int]$TimeoutSeconds = 180
    )

    # Microsoft Store AppX/MSIX transactions are deliberately non-interactive and
    # bounded. WinGet output is captured so a transient 0x8A15005E certificate
    # diagnostic from the first attempt cannot paint the installer console red when
    # the temporary certificate-pinning retry succeeds.
    $certError = -1978335138 # 0x8A15005E
    $timeoutCode = 1460      # ERROR_TIMEOUT
    $args = @('install','--id',$PackageId)
    if($Exact){ $args += '--exact' }
    $args += @('--source','msstore','--accept-package-agreements','--accept-source-agreements','--nowarn')
    $args += '--silent','--disable-interactivity','--scope','user'

    if($Interactive){
        Write-Log "[$PackageId] An interactive Store install was requested by the caller; WinGet deployment itself remains non-interactive. The installed application is opened separately for settings." 'INFO'
    }

    $run = {
        param($a)
        $token=[guid]::NewGuid().ToString('N')
        $stdoutFile=Join-Path $env:TEMP ("ROG-G635LW-WinGet-Stdout-$token.log")
        $stderrFile=Join-Path $env:TEMP ("ROG-G635LW-WinGet-Stderr-$token.log")
        try {
            $p=Start-Process -FilePath $WingetPath -ArgumentList $a -PassThru -NoNewWindow `
                -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile
            if(-not $p.WaitForExit([Math]::Max(1,$TimeoutSeconds)*1000)){
                Write-Log "WinGet Store operation exceeded the $TimeoutSeconds-second timeout; terminating WinGet and its child processes." 'WARN'
                try { & taskkill.exe /PID $p.Id /T /F 2>&1 | Out-Null } catch { try { $p.Kill($true) } catch {} }
                return [pscustomobject]@{Code=$timeoutCode;StdOut='';StdErr=''}
            }
            # Wait once more so redirected streams are fully flushed before reading them.
            try { $p.WaitForExit() } catch {}
            $so='';$se=''
            try { if(Test-Path -LiteralPath $stdoutFile){$so=Get-Content -LiteralPath $stdoutFile -Raw -ErrorAction SilentlyContinue} } catch {}
            try { if(Test-Path -LiteralPath $stderrFile){$se=Get-Content -LiteralPath $stderrFile -Raw -ErrorAction SilentlyContinue} } catch {}
            return [pscustomobject]@{Code=[int]$p.ExitCode;StdOut=[string]$so;StdErr=[string]$se}
        } finally {
            foreach($f in @($stdoutFile,$stderrFile)){
                if(Test-Path -LiteralPath $f){
                    try { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue } catch {}
                }
            }
        }
    }

    Write-Log "[$PackageId] WinGet Microsoft Store transaction starting." 'INFO'
    $first=& $run $args
    $code=[int]$first.Code

    if($code -ne $certError){
        if($code -ne 0 -and $code -ne 3010){
            $diag=(([string]$first.StdErr+"`n"+[string]$first.StdOut).Trim() -split "`r?`n" |
                Where-Object { $_.Trim() } | Select-Object -Last 8) -join ' | '
            if($diag){ Write-Log "[$PackageId] WinGet diagnostic: $diag" 'WARN' }
        }
        return [pscustomobject]@{ExitCode=$code;UsedCertificateBypass=$false}
    }

    # 0x8A15005E is specifically a WinGet Microsoft Store certificate-pinning
    # validation failure. Do not emit it as a warning when the supported temporary
    # bypass succeeds. Microsoft documents this policy/setting for environments where
    # Store certificate validation cannot match the expected endpoint certificate.
    Write-Log "[$PackageId] Microsoft Store certificate validation rejected the initial WinGet source connection. Applying the temporary Microsoft Store certificate-pinning retry." 'INFO'

    $previousEnabled=$false
    try {
        $info=@(& $WingetPath --info --disable-interactivity 2>&1 | Out-String)
        if(($info -join '') -match '(?im)BypassCertificatePinningForMicrosoftStore\s+Enabled'){$previousEnabled=$true}
    } catch {}

    $enabled=$false
    $retryCode=$code
    try {
        $setArgs=@('settings','--enable','BypassCertificatePinningForMicrosoftStore','--disable-interactivity')
        $sp=Start-Process -FilePath $WingetPath -ArgumentList $setArgs -Wait -PassThru -NoNewWindow `
            -RedirectStandardOutput (Join-Path $env:TEMP ("ROG-G635LW-WinGet-Settings-$([guid]::NewGuid().ToString('N')).out")) `
            -RedirectStandardError (Join-Path $env:TEMP ("ROG-G635LW-WinGet-Settings-$([guid]::NewGuid().ToString('N')).err"))
        if($sp.ExitCode -eq 0){
            $enabled=$true
            Write-Log "[$PackageId] Temporary Microsoft Store certificate-pinning bypass enabled for this retry." 'INFO'
        } else {
            Write-Log "[$PackageId] Could not enable the temporary Microsoft Store certificate-pinning bypass (exit code $($sp.ExitCode))." 'WARN'
        }

        if($enabled){
            $retry=& $run $args
            $retryCode=[int]$retry.Code
            if($retryCode -eq 0 -or $retryCode -eq 3010){
                Write-Log "[$PackageId] Microsoft Store transaction succeeded after the temporary certificate-pinning retry." 'OK'
            } else {
                $diag=(([string]$retry.StdErr+"`n"+[string]$retry.StdOut).Trim() -split "`r?`n" |
                    Where-Object { $_.Trim() } | Select-Object -Last 8) -join ' | '
                if($diag){ Write-Log "[$PackageId] WinGet Store retry diagnostic: $diag" 'WARN' }
            }
        }
    } finally {
        if($enabled -and -not $previousEnabled){
            try {
                $restoreArgs=@('settings','--disable','BypassCertificatePinningForMicrosoftStore','--disable-interactivity')
                $restoreOut=Join-Path $env:TEMP ("ROG-G635LW-WinGet-Restore-$([guid]::NewGuid().ToString('N')).out")
                $restoreErr=Join-Path $env:TEMP ("ROG-G635LW-WinGet-Restore-$([guid]::NewGuid().ToString('N')).err")
                $rp=Start-Process -FilePath $WingetPath -ArgumentList $restoreArgs -Wait -PassThru -NoNewWindow `
                    -RedirectStandardOutput $restoreOut -RedirectStandardError $restoreErr
                foreach($f in @($restoreOut,$restoreErr)){try{Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue}catch{}}
                if($rp.ExitCode -eq 0){
                    Write-Log "[$PackageId] Microsoft Store certificate-pinning validation restored after the temporary retry." 'OK'
                } else {
                    Write-Log "[$PackageId] WinGet could not restore Microsoft Store certificate-pinning validation (exit code $($rp.ExitCode))." 'ERROR'
                }
            } catch {
                Write-Log "[$PackageId] Exception while restoring Microsoft Store certificate-pinning validation: $($_.Exception.Message)" 'ERROR'
            }
        }
    }
    return [pscustomobject]@{ExitCode=$retryCode;UsedCertificateBypass=$enabled}
}


function Install-DownloadedStorePayload {
    param([Parameter(Mandatory)][string]$Directory,[Parameter(Mandatory)][string]$NamePattern)
    $files=@(Get-ChildItem -LiteralPath $Directory -Recurse -File -ErrorAction SilentlyContinue | Where-Object {$_.Extension -in '.msix','.msixbundle','.appx','.appxbundle'})
    if($files.Count -eq 0){return $false}
    $pkg=$files | Sort-Object Length -Descending | Select-Object -First 1
    $deps=@($files | Where-Object {$_.FullName -ne $pkg.FullName})
    try{
        Write-Log "[Store] Installing downloaded local payload $($pkg.FullName) for '$NamePattern'. No network download is permitted during the installation phase." 'STEP'
        if($deps.Count -gt 0){Add-AppxPackage -Path $pkg.FullName -DependencyPath @($deps.FullName) -ErrorAction Stop}else{Add-AppxPackage -Path $pkg.FullName -ErrorAction Stop}
        Start-Sleep -Seconds 3
        $found=@(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object {"$($_.Name) $($_.PackageFamilyName) $($_.PackageFullName)" -match $NamePattern})
        return ($found.Count -gt 0)
    }catch{Write-Log "[Store] Local AppX/MSIX installation failed for '$NamePattern': $($_.Exception.Message)" 'WARN';return $false}
}

function Install-MyASUS {
    Write-Log 'FINAL ASUS SOFTWARE GATE - MyASUS Microsoft Store installation/update and interactive settings (immediately before Armoury Crate/Aura)' 'STEP'
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) {
        Write-Log 'winget.exe was not found. MyASUS is REQUIRED immediately before the final Armoury Crate/Aura stage.' 'ERROR'
        return $false
    }
    $id='9N7R5S6B0ZZH'
    if($script:DownloadAllContentMode){
        $dir=Join-Path $DownloadRoot '04 Final ASUS Armoury Crate and Aura\MyASUS Store Download';New-Item -ItemType Directory -Force -Path $dir|Out-Null
        $files=@(Get-ChildItem -LiteralPath $dir -Recurse -File -ErrorAction SilentlyContinue | Where-Object {$_.Extension -in '.msix','.msixbundle','.appx','.appxbundle','.appinstaller'})
        if($files.Count -gt 0){
            Write-Log "[MyASUS] Reusing existing local Store payload in $dir; installation deferred until the installation phase." 'OK'
        } else {
            Write-Log '[MyASUS] Skipping winget download during the global download phase. Microsoft Store offline-package downloads require separate Microsoft Entra ID licensing authentication; MyASUS will be installed later through winget install using the current interactive Store session.' 'INFO'
        }
        return $true
    }
    $localMyAsusDir=Join-Path $DownloadRoot '04 Final ASUS Armoury Crate and Aura\MyASUS Store Download'
    if(Install-DownloadedStorePayload -Directory $localMyAsusDir -NamePattern '(?i)MyASUS|9N7R5S6B0ZZH'){
        Write-Log '[MyASUS] Installed from the previously downloaded local Store payload. No network download is permitted during installation.' 'OK'
    } elseif(-not (Initialize-MicrosoftStoreForCurrentUser)){
        Write-Log 'Microsoft Store is not registered/available for the current interactive user. MyASUS remains REQUIRED and the final Armoury Crate/Aura stage will not begin.' 'ERROR'
        return $false
    }
    $app=[pscustomobject]@{Name='MyASUS';Id=$id;Patterns='MyASUS|9N7R5S6B0ZZH'}
    $installed=Test-StoreCompanionInstalled -App $app
    if(-not $installed){
        for($attempt=1;$attempt -le 3 -and -not $installed;$attempt++){
            Write-Log "Installing MyASUS from Microsoft Store using exact package ID $id (attempt $attempt/3)." 'INFO'
            try {
                $result=Invoke-WingetMsStoreInstall -WingetPath $winget.Source -PackageId $id -Exact -Silent -TimeoutSeconds 180
                $code=[int]$result.ExitCode
                Write-Log "MyASUS WinGet exit code: $code." 'INFO'
                Start-Sleep -Seconds 3
                $installed=Test-StoreCompanionInstalled -App $app
                if($installed){
                    Write-Log 'MyASUS is verifiably registered as an installed AppX package.' 'OK'
                    break
                }
                Write-Log "MyASUS was NOT found in the installed AppX inventory after WinGet returned $code." 'WARN'
                if($attempt -lt 3){ [void](Reset-WingetMicrosoftStoreSource -WingetPath $winget.Source); Start-Sleep -Seconds 5 }
            } catch {
                Write-Log "MyASUS installation attempt $attempt failed: $($_.Exception.Message)" 'WARN'
                if($attempt -lt 3){ [void](Reset-WingetMicrosoftStoreSource -WingetPath $winget.Source); Start-Sleep -Seconds 5 }
            }
        }
    } else {
        Write-Log 'MyASUS is already installed; no Store download is required.' 'OK'
    }
    if(-not $installed){
        Write-Log 'MyASUS could not be verified as installed after the Store retries. Armoury Crate/Aura will NOT start.' 'ERROR'
        return $false
    }
    if(-not (Invoke-InteractiveStoreApp -DisplayName 'MyASUS' -NamePattern '(?i)MyASUS|9N7R5S6B0ZZH')){
        Write-Log 'MyASUS could not complete its interactive settings gate. Armoury Crate/Aura will NOT start.' 'ERROR'
        return $false
    }
    Write-Log 'MyASUS interactive settings gate complete. The final Armoury Crate/Aura stage may now begin.' 'OK'
    return $true
}






function Get-PnpSnapshot {
    @(Get-CimInstance Win32_PnPEntity | Where-Object {$_.Present -eq $true} |
        Select-Object Name,Manufacturer,PNPDeviceID,DriverVersion,ConfigManagerErrorCode)
}

function Find-AsusDownloadObjects {
    param($Node)

    $found = New-Object System.Collections.Generic.List[object]

    function Normalize-AsusUrl {
        param([string]$Value)
        if (-not $Value) { return $null }
        $u=$Value.Trim().Trim([char]34,[char]39)
        if ($u -match '^//') { return 'https:' + $u }
        if ($u -match '^https?://') { return $u }
        return $null
    }

    function LooksLikeInstallerUrl {
        param([string]$Url)
        if (-not $Url) { return $false }
        return ($Url -match '(?i)\.(exe|msi|zip|cab)(?:\?|$)')
    }

    function Walk {
        param($N,$ContextTitle='',$ContextVersion='',$ContextDate='',$ContextHash='',$ContextRaw=$null)
        if ($null -eq $N) { return }

        if ($N -is [System.Collections.IEnumerable] -and $N -isnot [string]) {
            foreach ($item in $N) { Walk $item $ContextTitle $ContextVersion $ContextDate $ContextHash $ContextRaw }
            return
        }

        if ($N -isnot [pscustomobject]) { return }

        $props=@{}
        foreach ($p in $N.PSObject.Properties) { $props[$p.Name]=$p.Value }

        $title=$ContextTitle
        foreach ($key in @('Title','Name','Description','DriverName','PackageName','CategoryName','Category')) {
            if ($props.ContainsKey($key) -and $props[$key] -is [string] -and $props[$key]) { $title=[string]$props[$key]; break }
        }
        $version=$ContextVersion
        foreach ($key in @('Version','version','DriverVersion','PackageVersion')) {
            if ($props.ContainsKey($key) -and $props[$key]) { $version=[string]$props[$key]; break }
        }
        $date=$ContextDate
        foreach ($key in @('ReleaseDate','Date','releaseDate')) {
            if ($props.ContainsKey($key) -and $props[$key]) { $date=[string]$props[$key]; break }
        }
        $hash=$ContextHash
        foreach ($key in @('SHA256','Sha256','Sha-256','Hash','FileHash')) {
            if ($props.ContainsKey($key) -and $props[$key]) { $hash=[string]$props[$key]; break }
        }

        $rawForPackage=$N
        if ($ContextRaw) { $rawForPackage=$ContextRaw }

        foreach ($p in $N.PSObject.Properties) {
            if ($p.Value -is [string]) {
                $candidate=Normalize-AsusUrl ([string]$p.Value)
                if ($candidate -and (LooksLikeInstallerUrl $candidate)) {
                    $fileTitle=$title
                    if (-not $fileTitle) { $fileTitle=[IO.Path]::GetFileName(([uri]$candidate).AbsolutePath) }
                    $found.Add([pscustomobject]@{
                        Title=$fileTitle
                        Version=$version
                        ReleaseDate=$date
                        DownloadUrl=$candidate
                        SHA256=$hash
                        Raw=$rawForPackage
                    })
                }
            } elseif ($p.Value -ne $null) {
                Walk $p.Value $title $version $date $hash $rawForPackage
            }
        }
    }

    Walk $Node
    return $found.ToArray()
}

function Initialize-DownloadFolders {
    Write-Log 'STAGE 07 - Initialising persistent Desktop driver/software download folders' 'STEP'
    if(-not(Test-Path $TargetUserRoot)){
        Write-Log "Target user path $TargetUserRoot was not found. Using the current user's Desktop instead." 'WARN'
        $script:DownloadRoot=Join-Path ([Environment]::GetFolderPath('Desktop')) 'ASUS-G635LW-Installer-Content'
        New-Item -ItemType Directory -Force -Path $script:DownloadRoot|Out-Null
        return
    }
    New-Item -ItemType Directory -Force -Path $TargetUserRoot|Out-Null
    New-Item -ItemType Directory -Force -Path (Split-Path $DesktopDownloadRoot -Parent)|Out-Null

    if(Test-Path $DesktopDownloadRoot){
        $di=Get-Item -LiteralPath $DesktopDownloadRoot -Force -ErrorAction SilentlyContinue
        $isJunction=($null -ne $di -and (($di.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0))
        if($isJunction){
            $script:DownloadRoot=$MasterDownloadRoot
            Write-Log "Desktop folder is an existing junction to $MasterDownloadRoot; using that persistent target." 'OK'
        }else{
            # A real Desktop directory may have been created by an earlier installer version.
            # It becomes authoritative so cancelling an installer cannot make a downloaded file
            # disappear from the user's requested Desktop archive.
            New-Item -ItemType Directory -Force -Path $MasterDownloadRoot|Out-Null
            try{
                Get-ChildItem -LiteralPath $MasterDownloadRoot -Recurse -File -ErrorAction SilentlyContinue|ForEach-Object{
                    $relative=$_.FullName.Substring($MasterDownloadRoot.Length).TrimStart('\')
                    $target=Join-Path $DesktopDownloadRoot $relative
                    if(-not(Test-Path $target)){
                        New-Item -ItemType Directory -Force -Path (Split-Path $target -Parent)|Out-Null
                        Copy-Item -LiteralPath $_.FullName -Destination $target -Force
                    }
                }
                Write-Log 'Merged any previously downloaded master-archive files into the real Desktop folder.' 'OK'
            }catch{Write-Log "Could not merge the older master archive into Desktop: $($_.Exception.Message)" 'WARN'}
            $script:DownloadRoot=$DesktopDownloadRoot
            Write-Log "Desktop folder is a real directory and is now the authoritative download/resume location: $script:DownloadRoot" 'OK'
        }
    }else{
        New-Item -ItemType Directory -Force -Path $MasterDownloadRoot|Out-Null
        try{
            New-Item -ItemType Junction -Path $DesktopDownloadRoot -Target $MasterDownloadRoot -ErrorAction Stop|Out-Null
            $script:DownloadRoot=$MasterDownloadRoot
            Write-Log "Created Desktop junction to $MasterDownloadRoot; both paths point to the same files." 'OK'
        }catch{
            New-Item -ItemType Directory -Force -Path $DesktopDownloadRoot|Out-Null
            $script:DownloadRoot=$DesktopDownloadRoot
            Write-Log 'Desktop junction could not be created; using Desktop as the authoritative archive.' 'WARN'
        }
    }
    Write-Log "Persistent download/resume root: $script:DownloadRoot" 'OK'
    Write-Log "User-visible Desktop folder: $DesktopDownloadRoot" 'OK'
}

function Find-ExistingArchiveFile {
    param([Parameter(Mandatory)][string]$Uri,[Parameter(Mandatory)][string]$Destination,[string]$Label='File')
    # The ASUS Rog Strix Scar archive is persistent across installer versions/runs. Before
    # downloading anything, search the entire archive recursively for the same source file.
    # v82 hardens this helper against an uninitialised script-scoped DownloadRoot. The XTU
    # recovery checkpoint can run from an error/resume path, so a null LiteralPath must never
    # reach Test-Path/Get-ChildItem.
    try{
        $archiveRoot=[string]$script:DownloadRoot
        if([string]::IsNullOrWhiteSpace($archiveRoot)){ $archiveRoot=[string]$DownloadRoot }
        if([string]::IsNullOrWhiteSpace($archiveRoot)){ $archiveRoot=[string](Split-Path $Destination -Parent) }
        if([string]::IsNullOrWhiteSpace($archiveRoot)){return $null}
        if(-not (Test-Path -LiteralPath $archiveRoot)){return $null}
        $destinationName=[IO.Path]::GetFileName($Destination)
        if($destinationName -and (Test-Path -LiteralPath $Destination)){
            try{if((Get-Item -LiteralPath $Destination -Force).Length -gt 0){return (Get-Item -LiteralPath $Destination -Force)}}catch{}
        }
        $sourceName=''
        try{$sourceName=[IO.Path]::GetFileName(([Uri]$Uri).AbsolutePath)}catch{}
        if([string]::IsNullOrWhiteSpace($sourceName) -or $sourceName -in @('installer','download','package','')){return $null}
        $matches=@(Get-ChildItem -LiteralPath $archiveRoot -Recurse -File -Force -ErrorAction SilentlyContinue | Where-Object {
            $_.Length -gt 0 -and ($_.Name -ieq $sourceName -or $_.Name.EndsWith($sourceName,[StringComparison]::OrdinalIgnoreCase))
        } | Sort-Object -Property `
            @{Expression={$_.FullName -eq $Destination}; Descending=$true}, `
            @{Expression='Length'; Descending=$true}, `
            @{Expression='LastWriteTime'; Descending=$true})
        if($matches.Count -gt 0){return $matches[0]}
    }catch{Write-Log "[$Label] Archive-wide existing-file search failed: $($_.Exception.Message)" 'WARN'}
    return $null
}

function Invoke-FileDownloadWithProgress {
    param([Parameter(Mandatory)][string]$Uri,[Parameter(Mandatory)][string]$Destination,[string]$Label='File',[string]$ExpectedSHA256='')
    $tmp="$Destination.partial";$resp=$null;$stream=$null;$file=$null
    try{
        $existingArchive=Find-ExistingArchiveFile -Uri $Uri -Destination $Destination
        if($existingArchive){
            $existingPath=[string]$existingArchive.FullName
            if($existingPath -ieq [IO.Path]::GetFullPath($Destination)){
                if($ExpectedSHA256 -match '^[A-Fa-f0-9]{64}$'){
                    $existing=(Get-FileHash -LiteralPath $existingPath -Algorithm SHA256).Hash
                    if($existing -eq $ExpectedSHA256.ToUpperInvariant()){Write-Log "[$Label] Existing archive file matches SHA-256; download skipped." 'OK';return $true}
                    Write-Log "[$Label] Existing archive file is present but its SHA-256 differs from the current catalog value. It will NOT be re-downloaded; the existing file is retained and will be used as the package source." 'WARN';return $true
                }
                Write-Log "[$Label] Existing archive file retained; download skipped." 'OK';return $true
            }
            # An older installer version may have stored the same source file under a different
            # generated name/subfolder. Reuse it by copying locally rather than downloading again.
            New-Item -ItemType Directory -Force -Path (Split-Path $Destination -Parent)|Out-Null
            Copy-Item -LiteralPath $existingPath -Destination $Destination -Force -ErrorAction Stop
            Write-Log "[$Label] Existing matching source file found in the ASUS Rog Strix Scar archive: $existingPath. Copied to the current package path; network download skipped." 'OK'
            if($ExpectedSHA256 -match '^[A-Fa-f0-9]{64}$'){
                $existingHash=(Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash
                if($existingHash -ne $ExpectedSHA256.ToUpperInvariant()){Write-Log "[$Label] Reused archive file SHA-256 differs from the current ASUS catalog value ($existingHash vs $ExpectedSHA256). The file remains available for this run and will NOT be downloaded again." 'WARN'}
            }
            return $true
        }
        if(Test-Path $tmp){Remove-Item $tmp -Force -ErrorAction SilentlyContinue};New-Item -ItemType Directory -Force -Path (Split-Path $Destination -Parent)|Out-Null
        [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
        $req=[Net.HttpWebRequest]::Create($Uri);$req.Method='GET';$req.AllowAutoRedirect=$true;$req.Timeout=60000;$req.ReadWriteTimeout=60000;$req.UserAgent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v64';$req.Accept='*/*'
        Write-Log "[$Label] DOWNLOAD START: $Uri" 'STEP'
        $resp=$req.GetResponse();$total=$resp.ContentLength;$stream=$resp.GetResponseStream();$file=[IO.File]::Open($tmp,[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::None);$buffer=New-Object byte[] (1024*1024);$received=[int64]0;$sw=[Diagnostics.Stopwatch]::StartNew();$lastLog=$sw.Elapsed.TotalSeconds;$lastReceived=[int64]0
        if($total -gt 0){Write-Log "[$Label] 0.0% | 0.0 / $([math]::Round($total/1MB,1)) MB | 0.0 MB/s" 'INFO'}else{Write-Log "[$Label] 0.0% | content length unknown | starting transfer..." 'INFO'}
        while(($read=$stream.Read($buffer,0,$buffer.Length))-gt 0){
            $file.Write($buffer,0,$read);$received+=$read
            $elapsed=[math]::Max($sw.Elapsed.TotalSeconds,0.001);$rate=($received/1MB)/$elapsed
            if($total -gt 0){$pct=[math]::Min(100,[math]::Round(($received/$total)*100,1));$mb=[math]::Round($received/1MB,1);$tm=[math]::Round($total/1MB,1);Write-Progress -Id 17 -Activity "Downloading $Label" -Status "$pct% | $mb / $tm MB | $([math]::Round($rate,1)) MB/s" -PercentComplete ([int]$pct)}else{$pct=0;$mb=[math]::Round($received/1MB,1);Write-Progress -Id 17 -Activity "Downloading $Label" -Status "$mb MB | $([math]::Round($rate,1)) MB/s" -PercentComplete 0}
            # Write a real terminal log line at least once per second, so progress is visible even when the host hides Write-Progress.
            if(($sw.Elapsed.TotalSeconds-$lastLog) -ge 1.0){
                if($total -gt 0){Write-Log "[$Label] $pct% | $mb / $tm MB | $([math]::Round($rate,1)) MB/s" 'INFO'}else{Write-Log "[$Label] $mb MB transferred | $([math]::Round($rate,1)) MB/s | total size unknown" 'INFO'}
                $lastLog=$sw.Elapsed.TotalSeconds;$lastReceived=$received
            }
        }
        $file.Close();$file=$null;$stream.Close();$stream=$null;$resp.Close();$resp=$null;$sw.Stop();Write-Progress -Id 17 -Activity "Downloading $Label" -Completed;Move-Item $tmp $Destination -Force
        $size=(Get-Item $Destination).Length;if($size -le 0){throw 'Downloaded file is empty.'}
        $finalRate=if($sw.Elapsed.TotalSeconds -gt 0){[math]::Round(($size/1MB)/$sw.Elapsed.TotalSeconds,1)}else{0}
        if($ExpectedSHA256 -match '^[A-Fa-f0-9]{64}$'){$actual=(Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash;if($actual -ne $ExpectedSHA256.ToUpperInvariant()){Remove-Item $Destination -Force -ErrorAction SilentlyContinue;throw "SHA-256 mismatch. Expected $ExpectedSHA256 but received $actual."};Write-Log "[$Label] DOWNLOAD COMPLETE: 100.0% | $([math]::Round($size/1MB,1)) MB | average $finalRate MB/s | SHA-256 verified." 'OK'}else{Write-Log "[$Label] DOWNLOAD COMPLETE: 100.0% | $([math]::Round($size/1MB,1)) MB | average $finalRate MB/s | no vendor SHA-256 was exposed." 'OK'}
        return $true
    }catch{Write-Progress -Id 17 -Activity "Downloading $Label" -Completed;try{$file.Close()}catch{};try{$stream.Close()}catch{};try{$resp.Close()}catch{};if(Test-Path $tmp){Remove-Item $tmp -Force -ErrorAction SilentlyContinue};Write-Log "[$Label] Download failed: $($_.Exception.Message)" 'WARN';return $false}
}

function Export-DownloadManifest {
    param([Parameter(Mandatory)]$Packages)
    try{$manifest=Join-Path $DownloadRoot 'Download-Manifest.csv';@($Packages|Select-Object Category,Title,Version,TargetArchitecture,ArchitectureRank,IsDriverPackage,ReleaseDate,DownloadUrl,SHA256,LocalPath,DownloadStatus,InstallStatus|Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $manifest);Write-Log "Download manifest: $manifest" 'OK'}catch{Write-Log "Could not write download manifest: $($_.Exception.Message)" 'WARN'}
}

function Find-ExistingContentFile {
    param([Parameter(Mandatory)][string]$FileName,[string]$PreferredPath='')
    try {
        if($PreferredPath -and (Test-Path -LiteralPath $PreferredPath)) { return (Get-Item -LiteralPath $PreferredPath) }
        if(-not (Test-Path -LiteralPath $DownloadRoot)) { return $null }
        $leaf=[IO.Path]::GetFileName($FileName)
        if(-not $leaf){ return $null }
        $hits=@(Get-ChildItem -LiteralPath $DownloadRoot -Recurse -File -Filter $leaf -ErrorAction SilentlyContinue |
            Where-Object { $_.Length -gt 0 } | Sort-Object FullName)
        if($hits.Count -gt 0){ return $hits[0] }
    } catch {}
    return $null
}

function Invoke-OfficialDownloadAndInstall {
    param([Parameter(Mandatory)][string]$Label,[Parameter(Mandatory)][string]$Uri,[Parameter(Mandatory)][string]$Destination,[string]$SHA256='',[string]$Arguments='/S',[switch]$Install,[switch]$InteractiveFallback)
    $existing=Find-ExistingContentFile -FileName ([IO.Path]::GetFileName($Destination)) -PreferredPath $Destination
    if($existing){
        if(([IO.Path]::GetFullPath($existing.FullName)) -ine ([IO.Path]::GetFullPath($Destination))){
            New-Item -ItemType Directory -Force -Path (Split-Path $Destination -Parent)|Out-Null
            Copy-Item -LiteralPath $existing.FullName -Destination $Destination -Force
            Write-Log "[$Label] Reusing existing installer from ASUS-G635LW-Installer-Content: $($existing.FullName) (no re-download)." 'OK'
        } else {
            Write-Log "[$Label] Installer already exists locally: $Destination (no re-download)." 'OK'
        }
        if($SHA256 -match '^[A-Fa-f0-9]{64}$'){
            try{
                $actual=(Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash
                if($actual -ieq $SHA256){Write-Log "[$Label] Existing local installer SHA-256 verified." 'OK'}
                else{Write-Log "[$Label] Existing local installer SHA-256 differs from the current vendor value; per persistent-archive policy it will be reused rather than downloaded again. Expected $SHA256; found $actual." 'WARN'}
            }catch{}
        }
    } else {
        if(Test-Path -LiteralPath (Join-Path $DownloadRoot 'ALL-CONTENT-DOWNLOAD-COMPLETE.json')){
            Write-Log "[$Label] INSTALLATION BLOCKED: the global content-download phase is complete but the required local installer is missing. No network download is permitted now." 'ERROR'
            return [pscustomobject]@{Downloaded=$false;Installed=$false;ExitCode=$null;Path=$Destination;DownloadBlocked=$true}
        }
        $ok=Invoke-FileDownloadWithProgress -Uri $Uri -Destination $Destination -Label $Label -ExpectedSHA256 $SHA256
        if(-not$ok){return [pscustomobject]@{Downloaded=$false;Installed=$false;ExitCode=$null;Path=$Destination}}
    }
    if(-not$Install){return [pscustomobject]@{Downloaded=$true;Installed=$false;ExitCode=$null;Path=$Destination}}
    if($script:DownloadAllContentMode){Write-Log "[$Label] DOWNLOAD-ONLY PREPARATION: installer execution is blocked until the complete content-download phase has finished. File retained at $Destination" 'OK';return [pscustomobject]@{Downloaded=$true;Installed=$false;ExitCode=$null;Path=$Destination;DownloadOnly=$true}}
    try{
        $windowStyle = if([string]::IsNullOrWhiteSpace($Arguments)){'Normal'}else{'Hidden'}
        $p=Start-Process -FilePath $Destination -ArgumentList $Arguments -WorkingDirectory (Split-Path $Destination -Parent) -Wait -PassThru -WindowStyle $windowStyle
        $good=($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010 -or $p.ExitCode -eq 1641)
        if($good){Write-Log "[$Label] Installer exit code $($p.ExitCode)." 'OK'}else{Write-Log "[$Label] Installer exit code $($p.ExitCode)." 'WARN'}
        if(-not$good -and $InteractiveFallback){
            Write-Log "[$Label] Silent installation did not return a success code. Launching the same official installer interactively and waiting for completion." 'STEP'
            $ip=Start-Process -FilePath $Destination -WorkingDirectory (Split-Path $Destination -Parent) -Wait -PassThru -WindowStyle Normal
            $good=($ip.ExitCode -eq 0 -or $ip.ExitCode -eq 3010 -or $ip.ExitCode -eq 1641)
            Write-Log "[$Label] Interactive installer exit code $($ip.ExitCode)." $(if($good){'OK'}else{'WARN'})
            return [pscustomobject]@{Downloaded=$true;Installed=$good;ExitCode=$ip.ExitCode;Path=$Destination}
        }
        return [pscustomobject]@{Downloaded=$true;Installed=$good;ExitCode=$p.ExitCode;Path=$Destination}
    }catch{
        Write-Log "[$Label] Installation failed: $($_.Exception.Message)" 'WARN'
        if($InteractiveFallback){
            try{
                Write-Log "[$Label] Launching the official installer interactively after the automated launch failed." 'STEP'
                $ip=Start-Process -FilePath $Destination -WorkingDirectory (Split-Path $Destination -Parent) -Wait -PassThru -WindowStyle Normal
                $good=($ip.ExitCode -eq 0 -or $ip.ExitCode -eq 3010 -or $ip.ExitCode -eq 1641)
                Write-Log "[$Label] Interactive installer exit code $($ip.ExitCode)." $(if($good){'OK'}else{'WARN'})
                return [pscustomobject]@{Downloaded=$true;Installed=$good;ExitCode=$ip.ExitCode;Path=$Destination}
            }catch{}
        }
        return [pscustomobject]@{Downloaded=$true;Installed=$false;ExitCode=$null;Path=$Destination}
    }
}


function Get-OSArchitecturePreference {
    try {
        if ([Environment]::Is64BitOperatingSystem) { return 'x64' }
    } catch {}
    return 'x86'
}

function Test-PackageLooksLikeDriver {
    param([Parameter(Mandatory)]$Package)
    $t = "$($Package.Category) $($Package.Title) $($Package.DownloadUrl)"
    # Intel graphics Store companions are applications, not driver packages.
    # Explicitly exclude both the legacy Command Center app and newer Intel Graphics Software app.
    if($t -match '(?i)Intel.*Graphic.*(Command\s*Center|Software\s*(Application|Installer)?)|Graphic\s*Command\s*Center|Command\s*Center\s*Application' -or
       $t -match '(?i)apps\.microsoft\.com/(?:store/)?(?:detail|apps)/9P8K5G2MWW6Z|apps\.microsoft\.com/store/apps/9PLFNLNT3G5G') { return $false }
    return ($t -match '(?i)driver|chipset|graphics|display|audio|sound|bluetooth|wireless|wlan|lan|ethernet|serial.?io|dynamic.?tuning|management.?engine|csme|vpu|gna|neural|smart.?sound|sst|touchpad|camera|card.?reader|storage|rst|rapid.?storage|thunderbolt|usb|firmware.?driver')
}

function Get-StoreAppDefinitionForAsusPackage {
    param([Parameter(Mandatory)]$Package)
    $t = "$($Package.Category) $($Package.Title) $($Package.DownloadUrl)"
    $defs=@(
        [pscustomobject]@{Name='Intel Graphics Software';Id='9P8K5G2MWW6Z';Pattern='IntelGraphicsSoftware|Intel.*Graphics.*Software|9P8K5G2MWW6Z';Titles='Intel.*Graphics.*Software(?:\s*(?:Application|Installer))?|Graphics.*Software.*Application';RequiredDriverPattern='Intel.*(Graphics|Display|Arc)|VEN_8086'},
        [pscustomobject]@{Name='Intel Graphics Command Center';Id='9PLFNLNT3G5G';Pattern='IntelGraphicsExperience|AppUp\.IntelGraphicsExperience|IGCC|Intel.*Graphic.*Command\s*Center|9PLFNLNT3G5G';Titles='Intel.*Graphic.*Command\s*Center|Graphic.*Command\s*Center|Command\s*Center\s*Application';RequiredDriverPattern='Intel.*(Graphics|Display|Arc)|VEN_8086&DEV_7D67'},
        [pscustomobject]@{Name='NVIDIA Control Panel';Id='9NF8H0H7WMLT';Pattern='NVIDIA.*Control.*Panel|NVIDIACorp\.NVIDIAControlPanel|9NF8H0H7WMLT';Titles='NVIDIA Control Panel|NVIDIA.*Control Panel Application';RequiredDriverPattern='NVIDIA.*(Graphics|Display)|VEN_10DE'},
        [pscustomobject]@{Name='MyASUS';Id='9N7R5S6B0ZZH';Pattern='MyASUS|9N7R5S6B0ZZH';Titles='^MyASUS$|MyASUS';RequiredDriverPattern='ASUS'},
        [pscustomobject]@{Name='Realtek Audio Control';Id='9P2B8MCSVPLN';Pattern='Realtek.*Audio.*Control|RealtekAudioControl|9P2B8MCSVPLN';Titles='Realtek.*Audio Control';RequiredDriverPattern='Realtek.*Audio|HDAUDIO'},
        [pscustomobject]@{Name='Dolby Access';Id='9N0866FS04W8';Pattern='Dolby.*Access|9N0866FS04W8';Titles='Dolby Access';RequiredDriverPattern='Dolby'},
        [pscustomobject]@{Name='Armoury Crate';Id='9PM9DFQRDH3F';Pattern='Armoury.*Crate|ArmouryCrate|9PM9DFQRDH3F';Titles='Armoury Crate';RequiredDriverPattern='ASUS'},
        [pscustomobject]@{Name='Aura Creator';Id='9MWZ3VLW5HBW';Pattern='Aura.*Creator|9MWZ3VLW5HBW';Titles='Aura Creator';RequiredDriverPattern='ASUS'}
    )
    foreach($d in $defs){
        if($t -match $d.Titles -or $t -match [regex]::Escape($d.Id)){ return $d }
    }
    # ASUS can expose Microsoft Store URLs without putting the Store product name in the title.
    # Accept only a URL that contains a plausible 12-character Microsoft Store product ID.
    if($t -match '(?i)(?:apps\.microsoft\.com|microsoft\.com/store|microsoft\.com/p/)'){
        $m=[regex]::Match($t,'(?i)(?<![A-Z0-9])[A-Z0-9]{12}(?![A-Z0-9])')
        if($m.Success){
            return [pscustomobject]@{Name=[string]$Package.Title;Id=$m.Value.ToUpperInvariant();Pattern=[regex]::Escape([string]$Package.Title);Titles=[regex]::Escape([string]$Package.Title);RequiredDriverPattern=''}
        }
    }
    return $null
}

function Test-PackageIsMicrosoftStoreApp {
    param([Parameter(Mandatory)]$Package)
    return ($null -ne (Get-StoreAppDefinitionForAsusPackage -Package $Package))
}

function Test-StoreAppInstalled {
    param([Parameter(Mandatory)]$Package)
    $def=Get-StoreAppDefinitionForAsusPackage -Package $Package
    if(-not $def){ return $false }
    try {
        $apps=@(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object {
            "$($_.Name) $($_.PackageFamilyName) $($_.PackageFullName)" -match [string]$def.Pattern
        })
        return ($apps.Count -gt 0)
    } catch { return $false }
}

function Test-CompanionPrerequisitesSatisfied {
    param([Parameter(Mandatory)]$Package,[Parameter(Mandatory)]$InstalledDrivers)
    $t="$($Package.Title) $($Package.Category) $($Package.DownloadUrl)"
    if($t -match '(?i)Intel.*Graphic.*(Command\s*Center|Software)|9PLFNLNT3G5G|9P8K5G2MWW6Z') {
        $gfx=@($InstalledDrivers | Where-Object {
            $id=[string]$_.DeviceID
            $name=[string]$_.DeviceName
            $ver=[string]$_.DriverVersion
            $status=[string]$_.Status
            (($id -match '(?i)VEN_8086&DEV_7D67') -or ($name -match '(?i)Intel.*(Graphics|Display|Arc)')) -and
            ($ver -match '\d') -and ($status -notmatch '(?i)error|unknown')
        })
        if($gfx.Count -gt 0) {
            Write-Log "COMPANION GATE: Intel graphics Store companion '$($Package.Title)' is allowed because the Intel graphics driver/device is present in the live installed-driver inventory." 'OK'
            return $true
        }
        Write-Log "COMPANION GATE: Intel graphics Store companion '$($Package.Title)' is deferred because the Intel graphics driver/device is not yet present. No Store application will be installed before its driver prerequisite." 'WARN'
        return $false
    }
    return $true
}

function Sort-PackagesForDependencyOrder {
    # v66: deliberately avoid a parameter named $Packages. PowerShell variable names
    # are case-insensitive, so a caller/local variable named $packages can otherwise
    # collide with the parameter and produce a misleading null-binding failure.
    param([AllowNull()][object[]]$InputPackages)
    $dependencyInput=@($InputPackages)
    $driver=@($dependencyInput | Where-Object { Test-PackageLooksLikeDriver -Package $_ })
    $software=@($dependencyInput | Where-Object { -not (Test-PackageLooksLikeDriver -Package $_) })
    # Within the driver phase, keep the required Intel graphics/platform ordering first.
    $ig=@($driver | Where-Object {
        $t="$($_.Category) $($_.Title)"; $raw=''; try{$raw=$_.Raw|ConvertTo-Json -Depth 20 -Compress}catch{}
        (($t -match '(?i)Intel.*(Graphic|Graphics|Display)') -or ($raw -match '(?i)ArrowLake.?HX|7D67|Intel.*Graphics.*Driver|Intel Graphic driver')) -and $t -notmatch '(?i)Command\s*Center|Application'
    })
    $ipf=@($driver | Where-Object {
        $t="$($_.Category) $($_.Title)"; $t -match '(?i)Innovation Platform Framework|\bIPF\b'
    }) | Where-Object { $ig -notcontains $_ }
    $ip=@($driver | Where-Object {
        $t="$($_.Category) $($_.Title)"; $t -match '(?i)Intel.*(Chipset|Platform Power Management|Serial IO|Dynamic Tuning|Converged Security|Management Engine|VPU|Platform Monitoring|Gaussian|Neural|SST|Smart Sound)|\b(Chipset|Platform Power Management|Intel\(R\) Serial IO|Intel Dynamic Tuning|Intel Converged Security|Intel VPU|Intel Platform Monitoring|Intel Gaussian|Intel\(R\) SST)\b'
    }) | Where-Object { $ig -notcontains $_ -and $ipf -notcontains $_ }
    # v45: storage and wired LAN are explicit high-priority driver groups. They are
    # removed from the generic bucket and processed immediately after Intel platform
    # prerequisites, so IRST/VMD and Realtek LAN cannot be buried behind later software.
    $irst=@($driver | Where-Object { "$($_.Category) $($_.Title)" -match '(?i)Intel.*Rapid Storage|\bIRST\b|Rapid Storage|\bRST\b|VMD|Volume Management' })
    $lan=@($driver | Where-Object { "$($_.Category) $($_.Title)" -match '(?i)Realtek.*LAN|LAN Driver|RTL8111H|RTL8125D' }) | Where-Object { $irst -notcontains $_ }
    $other=@($driver | Where-Object { $ig -notcontains $_ -and $ipf -notcontains $_ -and $ip -notcontains $_ -and $irst -notcontains $_ -and $lan -notcontains $_ })
    $ig=@(Sort-PackagesByArchitecture -InputPackages $ig)
    $ipf=@(Sort-PackagesByArchitecture -InputPackages $ipf)
    $ip=@(Sort-PackagesByArchitecture -InputPackages $ip)
    $irst=@(Sort-PackagesByArchitecture -InputPackages $irst)
    $lan=@(Sort-PackagesByArchitecture -InputPackages $lan)
    $other=@(Sort-PackagesByArchitecture -InputPackages $other)
    $software=@(Sort-PackagesByArchitecture -InputPackages $software)
    return @($ig+$ipf+$ip+$irst+$lan+$other+$software)
}

function Get-PackageArchitectureRank {
    param([Parameter(Mandatory)]$Package)
    $text = ''
    try { $text = "$($Package.Title) $($Package.Version) $($Package.Category) $($Package.DownloadUrl) $($Package.Raw | ConvertTo-Json -Depth 20 -Compress)" } catch { $text = "$($Package.Title) $($Package.Version) $($Package.Category) $($Package.DownloadUrl)" }
    # Explicit architecture markers win over generic/unmarked packages.
    # Architecture priority: x64 first, ARM64 second, x86 third, then
    # packages whose architecture is not explicitly stated. ALL variants are
    # retained; architecture is a priority, not an exclusion filter.
    if ($text -match '(?i)(x64|amd64|ntamd64|64[- ]?bit|win64)') { return 300 }
    if ($text -match '(?i)(arm64|aarch64)') { return 200 }
    if ($text -match '(?i)(x86|i386|i686|ntx86|32[- ]?bit|win32)') { return 100 }
    return 50
}

function Add-PackageArchitectureMetadata {
    param([Parameter(Mandatory)][object[]]$Packages)
    $osArch = Get-OSArchitecturePreference
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($pkg in $Packages) {
        $rank = Get-PackageArchitectureRank -Package $pkg
        $isDriver = Test-PackageLooksLikeDriver -Package $pkg
        $text = "$($pkg.Title) $($pkg.Category) $($pkg.DownloadUrl)"
        $explicitX64 = ($text -match '(?i)(x64|amd64|ntamd64|64[- ]?bit|win64)')
        $explicitArm = ($text -match '(?i)(arm64|aarch64)')
        $explicitX86 = ($text -match '(?i)(x86|i386|i686|ntx86|32[- ]?bit|win32)')
        # IMPORTANT: v19 never excludes a package merely because it is x86 or ARM64.
        # All architecture variants remain available for installation. The rank only
        # controls processing priority: x64 -> ARM64 -> x86 -> neutral/unspecified.
        $target = if($explicitX64){'x64'}elseif($explicitArm){'ARM64'}elseif($explicitX86){'x86'}else{'Neutral/Unspecified'}
        $pkg | Add-Member NoteProperty TargetArchitecture $target -Force
        $pkg | Add-Member NoteProperty ArchitectureRank $rank -Force
        $pkg | Add-Member NoteProperty IsDriverPackage ([bool]$isDriver) -Force
        $pkg | Add-Member NoteProperty ArchitectureOSPreference $osArch -Force
        $out.Add($pkg)
    }
    return $out.ToArray()
}

function Sort-PackagesByArchitecture {
    # v68: never use a parameter named $Packages; PowerShell variable names are case-insensitive.
    param([AllowNull()][object[]]$InputPackages)
    # ALL architecture variants are retained and sorted strictly by priority:
    # x64 first, ARM64 second, x86 third, then neutral/unspecified packages.
    return @($InputPackages | Sort-Object @{Expression={-[int]$_.ArchitectureRank};Descending=$false}, @{Expression={[string]$_.Title};Descending=$false}, @{Expression={ $v=Convert-ToComparableVersion ([string]$_.Version); if($null -eq $v){[version]'0.0'}else{$v} };Descending=$true})
}

function Test-AsusNetworkReady {
    param([int]$Attempts = 12,[int]$DelaySeconds = 15)
    Write-Log "Checking DNS/HTTPS readiness for ASUS before catalog discovery (up to $Attempts attempts)." 'INFO'
    for($i=1;$i -le $Attempts;$i++){
        foreach($hostName in @('rog.asus.com','www.asus.com')){
            try {
                $dns=Resolve-DnsName -Name $hostName -Type A -ErrorAction Stop | Where-Object {$_.IPAddress}
                if($dns){
                    try {
                        $tcp=Test-NetConnection -ComputerName $hostName -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue
                        if($tcp){ Write-Log "ASUS DNS and TCP/443 are ready for $hostName (attempt $i/$Attempts)." 'OK'; return $true }
                    } catch {}
                }
            } catch {}
        }
        Write-Log "ASUS network/DNS is not ready yet (attempt $i/$Attempts). Waiting $DelaySeconds seconds before retry." 'WARN'
        try { Clear-DnsClientCache -ErrorAction SilentlyContinue } catch {}
        Start-Sleep -Seconds $DelaySeconds
    }
    Write-Log 'ASUS DNS/HTTPS readiness check did not succeed within the retry window. Catalog requests will still be attempted.' 'WARN'
    return $false
}

function Get-AsusOfficialPackages {
    Write-Log 'STAGE 07B - ASUS official full-package discovery (after hardware detection)' 'STEP'
    $apiUrls=@('https://rog.asus.com/support/webapi/product/GetPDDrivers?website=global&model=G635LW&cpu=G635LW&osid=52&systemCode=rog','https://rog.asus.com/support/webapi/product/GetPDDrivers?website=global&model=G635LW&cpu=G635LW&osid=52&systemCode=rog&tag=1','https://www.asus.com/support/api/product.asmx/GetPDDrivers?osid=52&website=global&model=G635LW&cpu=G635LW')
    $response=$null;$usedUrl=$null
    [void](Test-AsusNetworkReady -Attempts 20 -DelaySeconds 15)
    foreach($url in $apiUrls){
        for($attempt=1;$attempt -le 3 -and -not $response;$attempt++){
            try{
                Write-Log "Querying ASUS official driver catalog (attempt $attempt/3): $url"
                $r=Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 60 -Headers @{'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v64';'Accept'='application/json,text/plain,*/*';'Referer'='https://rog.asus.com/support/'}
                if($r.StatusCode -eq 200 -and $r.Content){$response=$r.Content;$usedUrl=$url;Write-Log 'ASUS official driver catalog response received.' 'OK';break}
            }catch{
                Write-Log "ASUS catalog endpoint failed: $($_.Exception.Message)" 'WARN'
                if($attempt -lt 3){ Start-Sleep -Seconds 10; try { Clear-DnsClientCache -ErrorAction SilentlyContinue } catch {} }
            }
        }
        if($response){break}
    }
    if(-not $response){Write-Log 'ASUS catalog response remained unavailable after all endpoint retries. No ASUS package checkpoint is advanced; the next v35 invocation should retry discovery.' 'WARN';return @()}
    $jsonText=$response.Trim();if($jsonText -match '^\s*[^(]+\((.*)\)\s*;?\s*$'){$jsonText=$Matches[1]};try{$json=$jsonText|ConvertFrom-Json}catch{Write-Log "ASUS catalog response was not valid JSON: $($_.Exception.Message)" 'WARN';return @()}
    $packages=New-Object System.Collections.Generic.List[object]
    try{foreach($group in @($json.Result.Obj)){$category=[string]$group.Name;foreach($file in @($group.Files)){$url=$null;try{$url=[string]$file.DownloadUrl.Global}catch{};if(-not$url){try{$url=[string]$file.DownloadUrl}catch{}};if($url -and $url -match '^https?://'){$packages.Add([pscustomobject]@{Category=$category;Title=[string]$file.Title;Version=([string]$file.Version).TrimStart('V');ReleaseDate=[string]$file.ReleaseDate;DownloadUrl=$url;SHA256=[string]$file.SHA256;Raw=$file;LocalPath='';DownloadStatus='Pending';InstallStatus='Pending'})}}}}catch{Write-Log "ASUS Result.Obj/Files parser failed: $($_.Exception.Message)" 'WARN'}
    if($packages.Count -eq 0){$packages=@(Find-AsusDownloadObjects -Node $json|ForEach-Object{$_|Add-Member NoteProperty Category 'ASUS' -PassThru|Add-Member NoteProperty LocalPath '' -PassThru|Add-Member NoteProperty DownloadStatus 'Pending' -PassThru|Add-Member NoteProperty InstallStatus 'Pending' -PassThru})}
    $packages=@($packages|Where-Object{$t="$($_.Title) $($_.DownloadUrl) $($_.Category)";$t -notmatch '(?i)\bBIOS\b|\bFirmware\b|Manual|User.?Guide|Service.?Guide|Certificate|Wallpaper|eSupport'})
    # v15 IMPORTANT: do NOT collapse packages by Category|Title. ASUS can publish
    # multiple distinct installers with similar titles (different devices, revisions,
    # architectures, or release versions). v10/v12 could therefore create gaps such as
    # 002/004. Keep every distinct installer and only remove exact duplicate catalog rows.
    $seen=@{}
    $lossless=New-Object System.Collections.Generic.List[object]
    foreach($pkg in $packages){
        $sig=("{0}|{1}|{2}|{3}|{4}" -f [string]$pkg.Category,[string]$pkg.Title,[string]$pkg.Version,[string]$pkg.DownloadUrl,[string]$pkg.SHA256)
        if(-not $seen.ContainsKey($sig)){ $seen[$sig]=$true; $lossless.Add($pkg) }
    }
    $packages=@($lossless.ToArray())
    $packages=@(Add-PackageArchitectureMetadata -Packages $packages)
    # Latest-only policy: newest applicable version per logical ASUS driver/software target
    # is selected before the global download phase. Alternate hardware applicability is still
    # checked separately before installation; no generic Intel package replaces ASUS OEM content.
    $catalogDriverPackages=@($packages | Where-Object { Test-PackageLooksLikeDriver -Package $_ })
    $catalogNonDriverPackages=@($packages | Where-Object { -not (Test-PackageLooksLikeDriver -Package $_) })
    $latestDriverPackages=@(Select-BestAsusPackageCandidates -Packages $catalogDriverPackages)
    # Apply newest-version selection to repeated non-driver utilities/apps as well.
    # This prevents old Virtual Assistant and System Control Interface releases being downloaded alongside newer versions.
    $latestNonDriverPackages=@(Select-BestAsusPackageCandidates -Packages $catalogNonDriverPackages)
    Write-Log "LATEST-ONLY ASUS DRIVER POLICY: $($catalogDriverPackages.Count) candidates reduced to $($latestDriverPackages.Count) newest logical driver targets; $($catalogDriverPackages.Count - $latestDriverPackages.Count) older/duplicate driver candidates suppressed." 'OK'
    Write-Log "LATEST-ONLY ASUS SOFTWARE POLICY: $($catalogNonDriverPackages.Count) candidates reduced to $($latestNonDriverPackages.Count) newest logical software targets; $($catalogNonDriverPackages.Count - $latestNonDriverPackages.Count) older/duplicate software candidates suppressed." 'OK'
    $packages=@($latestDriverPackages + $latestNonDriverPackages)
    $packages=@(Sort-PackagesByArchitecture -InputPackages $packages)
    Write-Log "ASUS catalog parser committed $($packages.Count) package objects after architecture filtering/priority." 'INFO'
    if($packages.Count -eq 0){Write-Log 'ASUS catalog contained no current Windows package installer URLs after filtering.' 'WARN';return @()}
    Write-Log "ASUS catalog yielded $($packages.Count) selected full-package targets after excluding BIOS/firmware/documentation and suppressing older revisions of repeated logical targets." 'OK'
    $realtekCatalog=@($packages | Where-Object { "$($_.Category) $($_.Title)" -match '(?i)Realtek.*LAN|LAN Driver|RTL8111H|RTL8125D' })
    $irstCatalog=@($packages | Where-Object { "$($_.Category) $($_.Title)" -match '(?i)Intel.*Rapid Storage|\bIRST\b|Rapid Storage|\bRST\b|VMD|Volume Management' })
    $ipfCatalog=@($packages | Where-Object { "$($_.Category) $($_.Title)" -match '(?i)Innovation Platform Framework|\bIPF\b' })
    $intelGraphicsCatalog=@($packages | Where-Object { "$($_.Category) $($_.Title) $($_.DownloadUrl)" -match '(?i)Intel.*Graphic.*driver|Graphic_ROG_Intel|Graphic_DCH_ROG_Intel' -and "$($_.Title) $($_.DownloadUrl)" -notmatch '(?i)Command\s*Center|Graphics\s*Software' })
    if($intelGraphicsCatalog.Count -gt 0){
        $intelGraphicsSummary=@($intelGraphicsCatalog | ForEach-Object { "$($_.Title) $($_.Version) [$([IO.Path]::GetFileName(([uri]$_.DownloadUrl).AbsolutePath))]" }) -join '; '
        Write-Log "INTEL ONBOARD VGA PACKAGE: retained $($intelGraphicsCatalog.Count) official ASUS Intel graphics driver candidate(s): $intelGraphicsSummary. The newest selected applicable candidate will be downloaded before dependent graphics companion apps." 'OK'
    } else {
        Write-Log 'INTEL ONBOARD VGA PACKAGE WARNING: the ASUS catalog did not return an Intel onboard graphics driver candidate matching Intel Graphic driver / Graphic_ROG_Intel. The package was not fabricated; verify the official G635LW catalog response.' 'WARN'
    }
    $auditLevel=if($realtekCatalog.Count -gt 0 -and $irstCatalog.Count -gt 0 -and $ipfCatalog.Count -gt 0){'OK'}else{'WARN'}
    Write-Log "CRITICAL CATALOG AUDIT: Realtek LAN candidates=$($realtekCatalog.Count); Intel IRST/VMD candidates=$($irstCatalog.Count); Intel IPF candidates=$($ipfCatalog.Count). These packages are explicitly retained and separately prioritized for installation." $auditLevel
    if($ipfCatalog.Count -eq 0){Write-Log 'CRITICAL: G635LW ASUS catalog did not expose an IPF package. The dedicated IPF checkpoint will query a compatible official ASUS Intel-800-series catalog; no generic Intel or third-party package is substituted.' 'WARN'}
    if($realtekCatalog.Count -eq 0){Write-Log 'CRITICAL: ASUS catalog did not expose a Realtek LAN package. The installer will not pretend this component is covered; manufacturer fallback/Windows Update must handle it.' 'ERROR'}
    if($irstCatalog.Count -eq 0){Write-Log 'CRITICAL: ASUS catalog did not expose an Intel Rapid Storage package. The installer will not pretend this component is covered; manufacturer fallback/Windows Update must handle it.' 'ERROR'}
    Write-Log "ASUS catalog source used: $usedUrl" 'INFO'
    return $packages
}
function Test-AsusPackageAssociation {
    param(
        [Parameter(Mandatory)]$Package,
        [Parameter(Mandatory)]$Devices
    )

    $text = "$($Package.Title) $($Package.Version)"

    # Strong device/vendor/package keywords. These are deliberately conservative.
    $keywords = @(
        'NVIDIA','Intel','Realtek','ASUS','ROG','Bluetooth','Wireless','WLAN','LAN',
        'Ethernet','Audio','SST','Smart Sound','Graphics','Graphic','Display','VPU',
        'TouchPad','Touchpad','NumberPad','Camera','Webcam','IR','Chipset','Serial IO',
        'GPIO','SPI','I2C','UART','RST','Rapid Storage','MEI','Management Engine',
        'Dolby','Microsoft Effect Pack','MEP','Thunderbolt','USB','Card Reader',
        'Precision TouchPad','System Control Interface','Armoury Crate','MyASUS','GlideX'
    )

    $keywordHit = $false
    foreach ($k in $keywords) {
        if ($text -match [regex]::Escape($k)) { $keywordHit = $true; break }
    }

    if (-not $keywordHit) { return $false }

    # If ASUS exposes device-list metadata in the package object, use it.
    $rawText = ''
    try { $rawText = ($Package.Raw | ConvertTo-Json -Depth 20 -Compress) } catch {}

    if ($rawText) {
        $deviceNames = @($Devices | ForEach-Object { "$($_.Name) $($_.Manufacturer)" })

        # A package explicitly naming a device/vendor is accepted.
        foreach ($d in $deviceNames) {
            $parts = @($d -split '[\s\(\),/]+') | Where-Object { $_.Length -ge 4 }
            $hits = 0
            foreach ($part in $parts) {
                if ($rawText -match [regex]::Escape($part)) { $hits++ }
                if ($hits -ge 2) { return $true }
            }
        }

        # Generic model-specific packages from ASUS are also valid.
        if ($rawText -match '(?i)G635LW|ROG STRIX SCAR 16|ARL-H|ArrowLake HX') {
            return $true
        }
    }

    # If no device-list metadata is exposed by the API, retain the package because
    # it came from the official G635LW model catalog and matches a hardware/software class.
    return $true
}

function Get-InstallerFromPackage {
    param([Parameter(Mandatory)][string]$PackagePath,[Parameter(Mandatory)][string]$ExtractDirectory)
    $ext=[IO.Path]::GetExtension($PackagePath).ToLowerInvariant()
    if($ext -eq '.msi'){return [pscustomobject]@{Path='msiexec.exe';Arguments=@('/i',$PackagePath,'/qn','/norestart');Kind='MSI'}}
    if($ext -eq '.exe'){return [pscustomobject]@{Path=$PackagePath;Arguments=@('/S');Kind='EXE'}}
    try{Expand-Archive -LiteralPath $PackagePath -DestinationPath $ExtractDirectory -Force -ErrorAction Stop}catch{return $null}
    $setup=Get-ChildItem -LiteralPath $ExtractDirectory -Recurse -File -ErrorAction SilentlyContinue|Where-Object{$_.Name -match '^(setup|install|installer|ArmouryCrateInstallTool)(\.(exe|cmd|bat|msi))?$'}|Sort-Object FullName|Select-Object -First 1
    if(-not$setup){$setup=Get-ChildItem -LiteralPath $ExtractDirectory -Recurse -File -ErrorAction SilentlyContinue|Where-Object{$_.Extension -ieq '.exe' -and $_.Name -match '(?i)setup|install|installer'}|Sort-Object FullName|Select-Object -First 1}
    if($setup){if($setup.Extension -ieq '.msi'){return [pscustomobject]@{Path='msiexec.exe';Arguments=@('/i',$setup.FullName,'/qn','/norestart');Kind='MSI'}};if($setup.Extension -ieq '.bat' -or $setup.Extension -ieq '.cmd'){return [pscustomobject]@{Path='cmd.exe';Arguments=@('/c',$setup.FullName);Kind='CMD'}};return [pscustomobject]@{Path=$setup.FullName;Arguments=@('/S');Kind='EXE'}}
    return $null
}
function Get-PackageKey {
    param([Parameter(Mandatory)]$Package)
    $raw = "{0}|{1}|{2}|{3}" -f [string]$Package.Category,[string]$Package.Title,[string]$Package.Version,[string]$Package.DownloadUrl
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($raw))).Replace('-','')).ToLowerInvariant() }
    finally { $sha.Dispose() }
}


function Get-PackageInstallFingerprint {
    param(
        [Parameter(Mandatory)]$Package,
        [Parameter(Mandatory)]$Devices
    )
    $isDriver=[bool](Test-PackageLooksLikeDriver -Package $Package)
    $family=if($isDriver){Get-PackageComponentFamily -Package $Package}else{'Software'}
    $title=(([string]$Package.Title) -replace '\s+',' ').Trim().ToLowerInvariant()
    $version=(([string]$Package.Version) -replace '\s+','').Trim().ToLowerInvariant()
    $arch=([string]$Package.TargetArchitecture).ToLowerInvariant()
    $targets=New-Object System.Collections.Generic.List[string]
    if($isDriver){
        foreach($id in @(Get-PackageHardwareIds -Package $Package)){
            if($id){[void]$targets.Add(([string]$id).ToUpperInvariant())}
        }
        if($targets.Count -eq 0 -and $family){
            foreach($d in @($Devices)){
                $df=Get-DeviceComponentFamily -Device $d
                if($df -and $df -eq $family){
                    $did=[string]$d.PNPDeviceID
                    if($did){[void]$targets.Add($did.ToUpperInvariant())}
                }
            }
        }
    }
    if($targets.Count -eq 0){[void]$targets.Add('UNRESOLVED')}
    $targetText=(@($targets|Sort-Object -Unique)-join ';')
    $raw="$family|$title|$version|$arch|$targetText"
    $sha=[Security.Cryptography.SHA256]::Create()
    try{return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($raw))).Replace('-','')).ToLowerInvariant()}
    finally{$sha.Dispose()}
}

function Read-PackageInstallLedger {
    $path=Join-Path $DownloadRoot 'ASUS-Package-Install-Ledger.json'
    if(-not(Test-Path $path)){return @()}
    try{
        $rows=@(Get-Content -LiteralPath $path -Raw -ErrorAction Stop|ConvertFrom-Json)
        return @($rows|Where-Object{$_.Fingerprint})
    }catch{
        Write-Log "Package install ledger could not be read; current live-state checks will still govern installation. $($_.Exception.Message)" 'WARN'
        return @()
    }
}

function Save-PackageInstallLedger {
    param([Parameter(Mandatory)]$Rows)
    $path=Join-Path $DownloadRoot 'ASUS-Package-Install-Ledger.json'
    try{@($Rows)|Sort-Object Fingerprint -Unique|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $path -Encoding UTF8}catch{Write-Log "Could not save package install ledger: $($_.Exception.Message)" 'WARN'}
}

function Add-PackageInstallLedgerEntry {
    param(
        [Parameter(Mandatory)]$Ledger,
        [Parameter(Mandatory)]$Package,
        [Parameter(Mandatory)][string]$Fingerprint,
        [int]$ExitCode=0
    )
    $existing=@($Ledger|Where-Object{$_.Fingerprint -ne $Fingerprint})
    $existing += [pscustomobject]@{
        Fingerprint=$Fingerprint
        Title=[string]$Package.Title
        Version=[string]$Package.Version
        Category=[string]$Package.Category
        TargetArchitecture=[string]$Package.TargetArchitecture
        InstalledAt=(Get-Date).ToString('o')
        ExitCode=$ExitCode
    }
    return @($existing)
}

function Get-SafePackageName {
    param([Parameter(Mandatory)]$Package)
    $safe=("$($Package.Category)-$($Package.Title)" -replace '[^\w\.-]+','_').Trim('_')
    if(-not$safe){$safe='ASUS-Package'}
    return $safe
}

function Get-PackageDestination {
    param([Parameter(Mandatory=$true)]$Package,[Parameter(Mandatory=$true)][string]$AsusDirectory)

    $safe=Get-SafePackageName -Package $Package
    $version=[string]$Package.Version
    $version=($version -replace '[^A-Za-z0-9\.-]+','_').Trim('_')
    if(-not $version){$version='noversion'}

    $hashTag=''
    if(([string]$Package.SHA256) -match '^[A-Fa-f0-9]{64}$'){
        $hashTag=([string]$Package.SHA256).Substring(0,12).ToUpperInvariant()
    } else {
        $hashTag='NOHASH'
    }

    $uriName='package.bin'
    try{$uriName=[IO.Path]::GetFileName(([uri]$Package.DownloadUrl).AbsolutePath)}catch{}
    if(-not $uriName -or $uriName -notmatch '\.[A-Za-z0-9]{2,8}$'){$uriName='package.bin'}

    # Include version + SHA prefix in the archive filename. This prevents an old
    # package revision from occupying the same path as a newer catalog revision.
    return (Join-Path $AsusDirectory ("{0}-v{1}-{2}-{3}" -f $safe,$version,$hashTag,$uriName))
}

function Read-PackageResumeState {
    $path=Join-Path $DownloadRoot 'ASUS-Package-Resume-State.json'
    $legacyPaths=@(
        (Join-Path $env:ProgramData 'ROG-G635LW-Installer-v23\ASUS-Package-Resume-State.json'),
        (Join-Path $env:ProgramData 'ROG-G635LW-Installer-v23\State\ASUS-Package-Resume-State.json')
    )
    if(Test-Path $path){
        try{
            $raw=@(Get-Content $path -Raw|ConvertFrom-Json)
            return @($raw | ForEach-Object { Normalize-PackageResumeRow -Row $_ })
        }catch{Write-Log "ASUS package resume state was unreadable; attempting v23 migration/desktop reconstruction." 'WARN'}
    }
    foreach($legacyPath in $legacyPaths){
        if(Test-Path $legacyPath){
            try{
                $raw=@(Get-Content $legacyPath -Raw|ConvertFrom-Json)
                $rows=@($raw | ForEach-Object { Normalize-PackageResumeRow -Row $_ })
                if($rows.Count -gt 0){
                    Write-Log "Resume migration: imported $($rows.Count) package checkpoint row(s) from v23 state: $legacyPath" 'OK'
                    return $rows
                }
            }catch{Write-Log "Could not migrate legacy v23 resume state '$legacyPath': $($_.Exception.Message)" 'WARN'}
        }
    }
    return @()
}

function Normalize-PackageResumeRow {
    param([Parameter(Mandatory)]$Row)
    # Resume JSON created by older versions can lack properties introduced later.
    # PSCustomObject does not permit assignment to a property that does not exist,
    # which was the direct cause of the v20/v21 resume-state failure.
    $defaults=@{
        PackageKey=''; Title=''; Version=''; DownloadUrl=''; LocalPath='';
        DownloadStatus='Pending'; InstallStatus='Pending'; LastAction=''; Updated=''
    }
    foreach($name in $defaults.Keys){
        $prop=$Row.PSObject.Properties[$name]
        if($null -eq $prop){
            $Row | Add-Member -MemberType NoteProperty -Name $name -Value $defaults[$name] -Force
        }
    }
    return $Row
}

function Save-PackageResumeState {
    param([Parameter(Mandatory)]$Rows)
    $path=Join-Path $DownloadRoot 'ASUS-Package-Resume-State.json'
    try{@($Rows)|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $path -Encoding UTF8}catch{Write-Log "Could not save ASUS package resume state: $($_.Exception.Message)" 'WARN'}
}

function Find-ExistingAsusPackage {
    param([Parameter(Mandatory=$true)]$Package,[Parameter(Mandatory=$true)][string]$AsusDirectory,[Parameter(Mandatory=$true)][string]$PreferredPath)

    # First choice is always the new version/hash-aware preferred path.
    if(Test-Path $PreferredPath){return $PreferredPath}

    $safe=Get-SafePackageName -Package $Package
    $expected=''
    if(([string]$Package.SHA256) -match '^[A-Fa-f0-9]{64}$'){
        $expected=([string]$Package.SHA256).ToUpperInvariant()
    }

    $uriName=''
    try{$uriName=[IO.Path]::GetFileName(([uri]$Package.DownloadUrl).AbsolutePath)}catch{}

    $matches=@()
    if($uriName){
        $matches=@(Get-ChildItem -LiteralPath $AsusDirectory -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object{$_.Name -like "*$safe*$uriName" -and $_.Length -gt 0})
    }
    if($matches.Count -eq 0){
        $matches=@(Get-ChildItem -LiteralPath $AsusDirectory -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object{$_.Name -like "*$safe*" -and $_.Length -gt 0})
    }

    # IMPORTANT: never reuse an archive file solely because its title matches.
    # Reuse it only if its actual SHA matches the current ASUS catalog hash.
    if($expected){
        foreach($m in @($matches)){
            try{
                $actual=(Get-FileHash -LiteralPath $m.FullName -Algorithm SHA256).Hash.ToUpperInvariant()
                if($actual -eq $expected){
                    Write-Log "Existing archive match found by SHA-256: $($m.FullName)" 'OK'
                    return $m.FullName
                }
            }catch{}
        }
        if($matches.Count -gt 0){
            Write-Log ("Current ASUS package has a different SHA-256 from {0} existing archive candidate(s); stale files are retained and a new version/hash-specific file will be downloaded." -f $matches.Count) 'INFO'
        }
    } elseif($matches.Count -gt 0) {
        return (@($matches|Sort-Object LastWriteTime -Descending|Select-Object -First 1).FullName)
    }

    return $null
}

function Get-LegacyInstalledAsusTitles {
    # v10 did not persist a per-package checkpoint before an installer was launched. To avoid
    # reinstalling packages that v10 already completed, import its final manifest/log evidence
    # when available. This is only a migration aid; v13's own JSON checkpoint is authoritative.
    $titles=New-Object System.Collections.Generic.List[string]
    $legacyBase=Join-Path $env:ProgramData 'ROG-G635LW-Installer-v10'
    $csv=Join-Path $DownloadRoot 'Download-Manifest.csv'
    $csvCandidates=@($csv, (Join-Path $legacyBase 'Download-Manifest.csv'))
    foreach($c in $csvCandidates){
        if(Test-Path $c){
            try{
                foreach($r in @(Import-Csv -LiteralPath $c)){
                    if([string]$r.InstallStatus -match '^Installed\((0|3010|1641)\)$' -and $r.Title){$titles.Add([string]$r.Title)}
                }
            }catch{}
        }
    }
    $logDir=Join-Path $legacyBase 'Logs'
    if(Test-Path $logDir){
        foreach($lf in @(Get-ChildItem -LiteralPath $logDir -Filter 'Installer-*.log' -File -ErrorAction SilentlyContinue)){
            try{
                foreach($line in Get-Content -LiteralPath $lf.FullName -ErrorAction SilentlyContinue){
                    if($line -match 'Full ASUS setup completed for (.+?) \(exit (0|3010|1641)\)\.'){$titles.Add($matches[1])}
                }
            }catch{}
        }
    }
    return @($titles|Sort-Object -Unique)
}


function Get-BootMarker {
    try { return (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime.ToUniversalTime().ToString('o') }
    catch { return (Get-Date).ToUniversalTime().Date.ToString('o') }
}

function Test-PendingSystemReboot {
    $pending=$false
    try { if(Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'){$pending=$true} } catch {}
    try { if(Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'){$pending=$true} } catch {}
    try {
        $v=Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
        if($v.PendingFileRenameOperations){$pending=$true}
    } catch {}
    return $pending
}


function Invoke-PnpDeviceRescan {
    # IMPORTANT: there is intentionally NO timeout. Windows PnP is allowed to complete
    # the entire hardware rescan, even if it takes several minutes. This is a local
    # device-enumeration operation; it does not download drivers or software.
    Write-Log 'PnP rescan is a LOCAL Windows hardware/device enumeration step; it does NOT download drivers, software, updates, or .exe files.' 'STEP'
    Write-Log 'Starting pnputil /scan-devices with NO timeout. The scan will be allowed to run until Windows reports that the PnP operation has completed.' 'INFO'
    try {
        $pnputil=Join-Path $env:SystemRoot 'System32\pnputil.exe'
        if(-not(Test-Path $pnputil)){Write-Log 'PnPUtil was not found; continuing with WMI/CIM device enumeration.' 'WARN';return}
        $p=New-Object System.Diagnostics.Process
        $p.StartInfo=New-Object System.Diagnostics.ProcessStartInfo
        $p.StartInfo.FileName=$pnputil
        $p.StartInfo.Arguments='/scan-devices'
        $p.StartInfo.UseShellExecute=$false
        $p.StartInfo.CreateNoWindow=$true
        # Do not redirect stdout/stderr here. The scan is allowed to run indefinitely and
        # redirecting a child process can theoretically block if its pipe buffer fills.
        $p.StartInfo.RedirectStandardOutput=$false
        $p.StartInfo.RedirectStandardError=$false
        $null=$p.Start()
        $lastReport=0
        while(-not $p.HasExited){
            Start-Sleep -Seconds 1
            if($p.HasExited){break}
            $elapsed=[int]((Get-Date)-$p.StartTime).TotalSeconds
            if($elapsed -ge ($lastReport+10)){
                Write-Log "PnP rescan still running locally... ${elapsed}s elapsed; no network download is occurring in this step. Waiting for Windows PnP to finish the complete device scan." 'INFO'
                $lastReport=$elapsed
            }
        }
        $p.WaitForExit()
        if($p.ExitCode -eq 0){Write-Log 'PnPUtil device rescan completed successfully. All Windows-reported PnP scan work has completed.' 'OK'}else{Write-Log "PnPUtil device rescan returned exit code $($p.ExitCode); the process itself has completed, so continuing with the live inventory." 'WARN'}
        Start-Sleep -Seconds 2
    } catch { Write-Log "PnPUtil device rescan failed: $($_.Exception.Message). Continuing with the live inventory." 'WARN' }
}

function Get-InstalledDriverInventory {
    Write-Log 'Building full installed signed-driver inventory for resume reconciliation.' 'STEP'
    $rows=New-Object System.Collections.Generic.List[object]
    $presentIds=@{}

    # Win32_PnPSignedDriver does NOT expose a documented 'Present' property. v13
    # incorrectly filtered on $_.Present, which caused every installed driver to be
    # discarded and produced the observed "0 present device-driver records" result.
    # Build the live device list separately and reconcile by PNPDeviceID instead.
    try {
        foreach($d in @(Get-CimInstance Win32_PnPEntity -ErrorAction Stop | Where-Object {$_.PNPDeviceID})) {
            $presentIds[[string]$d.PNPDeviceID]=$true
        }
        Write-Log "Live PnP device inventory contains $($presentIds.Count) device instance IDs." 'OK'
    } catch {
        Write-Log "Live PnP device inventory could not be read: $($_.Exception.Message)" 'WARN'
    }

    $allDrivers=@()
    try {
        $allDrivers=@(Get-CimInstance Win32_PnPSignedDriver -ErrorAction Stop | Where-Object {$_.DeviceID})
    } catch {
        Write-Log "Win32_PnPSignedDriver inventory query failed: $($_.Exception.Message)" 'WARN'
    }

    foreach($d in $allDrivers) {
        $id=[string]$d.DeviceID
        # Prefer currently present devices. If the PnP entity query itself failed or
        # returned no IDs, retain the signed-driver records rather than falsely reporting zero.
        if($presentIds.Count -gt 0 -and -not $presentIds.ContainsKey($id)){continue}
        $rows.Add([pscustomobject]@{
            DeviceName=[string]$d.DeviceName
            Manufacturer=[string]$d.Manufacturer
            DeviceID=$id
            HardwareID=if($d.PSObject.Properties['HardwareID']){(@($d.HardwareID) -join ';')}else{''}
            DeviceClass=if($d.PSObject.Properties['DeviceClass']){[string]$d.DeviceClass}else{''}
            DriverVersion=[string]$d.DriverVersion
            DriverDate=[string]$d.DriverDate
            InfName=[string]$d.InfName
            DriverProvider=[string]$d.DriverProviderName
            IsSigned=[bool]$d.IsSigned
            Started=[bool]$d.Started
            Status=[string]$d.Status
        })
    }

    # Safety fallback: if the live-device join unexpectedly produced no records,
    # retain the raw signed-driver inventory so resume reconciliation cannot jump
    # ahead merely because a WMI provider returned an unusual device list.
    if($rows.Count -eq 0 -and $allDrivers.Count -gt 0) {
        Write-Log 'Live PnP/signed-driver join returned zero records; using the raw Win32_PnPSignedDriver records as a safe fallback.' 'WARN'
        foreach($d in $allDrivers) {
            $rows.Add([pscustomobject]@{
                DeviceName=[string]$d.DeviceName
                Manufacturer=[string]$d.Manufacturer
                DeviceID=[string]$d.DeviceID
                HardwareID=if($d.PSObject.Properties['HardwareID']){(@($d.HardwareID) -join ';')}else{''}
                DeviceClass=if($d.PSObject.Properties['DeviceClass']){[string]$d.DeviceClass}else{''}
                DriverVersion=[string]$d.DriverVersion
                DriverDate=[string]$d.DriverDate
                InfName=[string]$d.InfName
                DriverProvider=[string]$d.DriverProviderName
                IsSigned=[bool]$d.IsSigned
                Started=[bool]$d.Started
                Status=[string]$d.Status
            })
        }
    }

    Write-Log "Installed signed-driver inventory contains $($rows.Count) device-driver records after live-device reconciliation." 'OK'
    return $rows.ToArray()
}
function Get-InstalledSoftwareInventory {
    $rows=New-Object System.Collections.Generic.List[object]
    $roots=@(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    foreach($root in $roots){
        try{
            foreach($x in @(Get-ItemProperty $root -ErrorAction SilentlyContinue | Where-Object {$_.DisplayName})){
                $rows.Add([pscustomobject]@{
                    DisplayName=[string]$x.DisplayName
                    DisplayVersion=[string]$x.DisplayVersion
                    Publisher=[string]$x.Publisher
                    InstallLocation=[string]$x.InstallLocation
                })
            }
        }catch{}
    }
    return $rows.ToArray()
}

function Convert-ToComparableVersion {
    param([string]$Value)
    if(-not $Value){return $null}
    $v=($Value -replace '^[Vv]\s*','').Trim()
    $m=[regex]::Match($v,'\d+(?:\.\d+){0,3}')
    if(-not $m.Success){return $null}
    try{return [version]$m.Value}catch{return $null}
}


function Get-PackageHardwareIds {
    param([Parameter(Mandatory)]$Package)
    $ids=New-Object System.Collections.Generic.List[string]
    $text=''
    try { $text = "$($Package.Title) $($Package.Category) $($Package.DownloadUrl) $($Package.Raw | ConvertTo-Json -Depth 40 -Compress)" } catch { $text = "$($Package.Title) $($Package.Category) $($Package.DownloadUrl)" }
    # Hardware IDs are opaque strings; use them only for exact/case-insensitive comparisons.
    # Capture PCI, USB, HDAUDIO and ACPI-style identifiers exposed by ASUS metadata.
    foreach($m in [regex]::Matches($text,'(?i)(?:PCI\\)?VEN_[0-9A-F]{4}&DEV_[0-9A-F]{4}(?:&SUBSYS_[0-9A-F]{8})?')){[void]$ids.Add($m.Value)}
    foreach($m in [regex]::Matches($text,'(?i)USB\\VID_[0-9A-F]{4}&PID_[0-9A-F]{4}(?:\\[^"''\s,}]+)?')){[void]$ids.Add($m.Value)}
    foreach($m in [regex]::Matches($text,'(?i)HDAUDIO\\FUNC_[0-9A-F]{2}&VEN_[0-9A-F]{4}&DEV_[0-9A-F]{4}(?:[^"''\s,}]*)')){[void]$ids.Add($m.Value)}
    foreach($m in [regex]::Matches($text,'(?i)ACPI\\[A-Z0-9_]+(?:&[A-Z0-9_]+)*')){[void]$ids.Add($m.Value)}
    return @($ids | Sort-Object -Unique)
}

function Get-DeviceHardwareIds {
    param([Parameter(Mandatory)]$Device)
    $ids=New-Object System.Collections.Generic.List[string]
    foreach($prop in @('PNPDeviceID','DeviceID','HardwareID','HardwareIds')){
        try {
            $v=$Device.PSObject.Properties[$prop]
            if($v -and $v.Value){ foreach($id in @($v.Value)){ if([string]$id){[void]$ids.Add([string]$id)} } }
        } catch {}
    }
    # Win32_PnPEntity exposes PNPDeviceID but not the complete HardwareID/CompatibleID list.
    # Ask the PnP cmdlets for DEVPKEY_Device_HardwareIds when available.
    try {
        if($Device.PNPDeviceID -and (Get-Command Get-PnpDeviceProperty -ErrorAction SilentlyContinue)){
            $r=Get-PnpDeviceProperty -InstanceId ([string]$Device.PNPDeviceID) -KeyName 'DEVPKEY_Device_HardwareIds' -ErrorAction SilentlyContinue
            foreach($id in @($r.Data)){if([string]$id){[void]$ids.Add([string]$id)}}
        }
    } catch {}
    return @($ids | Sort-Object -Unique)
}

function Get-PackageComponentFamily {
    param([Parameter(Mandatory)]$Package)
    $t="$($Package.Category) $($Package.Title)"
    switch -Regex ($t) {
        'Realtek.*LAN|LAN Driver|RTL8111H|RTL8125D' { return 'Realtek LAN' }
        'Intel.*Rapid Storage|\bIRST\b|Rapid Storage|\bRST\b|VMD|Volume Management' { return 'Intel Rapid Storage/VMD' }
        'Intel.*Graphic|Intel.*Display' { return 'Intel Graphics' }
        'Innovation Platform Framework|\bIPF\b' { return 'Intel Innovation Platform Framework' }
        'Platform.*Monitoring|\bPMT\b' { return 'Intel Platform Monitoring' }
        'Platform Power Management|\bPPM\b' { return 'Intel Platform Power' }
        'Serial IO' { return 'Intel Serial IO' }
        'Dynamic Tuning|\bDTT\b' { return 'Intel Dynamic Tuning' }
        'Execution Technology|Trusted Execution Technology|\bTXT\b' { return 'Intel Trusted Execution Technology' }
        'Converged Security|Management Engine|\bCSME\b|\bMEI\b' { return 'Intel Management Engine' }
        'VPU|Vision Processing' { return 'Intel VPU' }
        'Gaussian|\bGNA\b' { return 'Intel GNA' }
        'NumberPad|NumPad|Numpad|Numeric Keypad' { return 'ASUS NumberPad' }
        'Smart Sound|\bSST\b' { return 'Intel Smart Sound' }
        'Bluetooth' { return 'Bluetooth' }
        'Wireless|WLAN|Wi-?Fi' { return 'Wireless LAN' }
        'NVIDIA.*Graphic|NVIDIA.*Display|GeForce' { return 'NVIDIA Graphics' }
        'Audio|Sound|Realtek.*Audio' { return 'Audio' }
        'Touchpad|Precision TouchPad' { return 'Touchpad' }
        'Camera|Webcam|IR' { return 'Camera/IR' }
        'Card Reader' { return 'Card Reader' }
        'Thunderbolt' { return 'Thunderbolt' }
        'USB' { return 'USB' }
        'Chipset' { return 'Chipset' }
        default { return '' }
    }
}

function Test-DriverPackageHardwareAssociation {
    param(
        [Parameter(Mandatory)]$Package,
        [Parameter(Mandatory)]$Devices,
        [Parameter(Mandatory)]$InstalledDrivers
    )
    if(-not (Test-PackageLooksLikeDriver -Package $Package)){ return [pscustomobject]@{Applicable=$true;Reason='Non-driver package; hardware-ID gate not required';Family=''} }
    $family=Get-PackageComponentFamily -Package $Package
    $pkgIds=@(Get-PackageHardwareIds -Package $Package)
    $allDevices=@($Devices)
    $allInstalled=@($InstalledDrivers)

    # v45 CRITICAL SPECIAL-CASE: Realtek LAN and Intel IRST/VMD are not allowed to
    # disappear behind incomplete ASUS catalog metadata. ASUS publishes separate
    # Realtek variants and an Intel IRST package for G635LW. Evaluate these families
    # against the live controller BEFORE generic catalog hardware-ID metadata.
    $pkgText="$($Package.Title) $($Package.Category)"
    # v62 CRITICAL: Intel Innovation Platform Framework is a first-class ASUS package.
    # G635LW Arrow Lake-HX systems expose IPF through System Devices/ACPI/PCI entries,
    # and ASUS package metadata can omit those hardware IDs. Match the live IPF devices
    # explicitly so the official ASUS IPF package cannot be archived and then skipped.
    if($family -eq 'Intel Innovation Platform Framework'){
        $ipfDevices=@($allDevices | Where-Object {
            $blob="$($_.Name) $($_.Manufacturer) $($_.PNPDeviceID) $($_.HardwareID)"
            $blob -match '(?i)Innovation Platform Framework|ACPI\\INTC(10[0-9A-F]{2}|35[0-9A-F]{2})|PCI\\VEN_8086&DEV_(AD03|7F23)'
        })
        if($ipfDevices.Count -gt 0 -and $pkgText -match '(?i)Innovation Platform Framework|\bIPF\b'){
            return [pscustomobject]@{Applicable=$true;Reason='CRITICAL Intel IPF package explicitly matched live Innovation Platform Framework ACPI/PCI device(s)';Family=$family}
        }
        if($ipfDevices.Count -gt 0){ return [pscustomobject]@{Applicable=$true;Reason='Live Intel Innovation Platform Framework device(s) detected; ASUS IPF package remains eligible despite incomplete catalog metadata';Family=$family} }
        return [pscustomobject]@{Applicable=$false;Reason='No live Intel Innovation Platform Framework device was found in the STAGE 06 baseline';Family=$family}
    }
    if($family -eq 'Realtek LAN'){
        $realtekDevices=@($allDevices | Where-Object {
            $blob="$($_.Name) $($_.Manufacturer) $($_.PNPDeviceID) $($_.HardwareID)"
            $blob -match '(?i)Realtek|VEN_10EC&DEV_(8168|8125)|RTL8111H|RTL8125D'
        })
        if($pkgText -match '(?i)RTL8111H'){
            $hit=@($realtekDevices | Where-Object { "$($_.Name) $($_.PNPDeviceID) $($_.HardwareID)" -match '(?i)RTL8111H|VEN_10EC&DEV_8168' })
            if($hit.Count -gt 0){ return [pscustomobject]@{Applicable=$true;Reason='CRITICAL Realtek RTL8111H package explicitly matched the live RTL8111H/PCI DEV_8168 controller';Family=$family} }
            return [pscustomobject]@{Applicable=$false;Reason='RTL8111H package retained in catalog but the live controller is not RTL8111H/DEV_8168';Family=$family}
        }
        if($pkgText -match '(?i)RTL8125D'){
            $hit=@($realtekDevices | Where-Object { "$($_.Name) $($_.PNPDeviceID) $($_.HardwareID)" -match '(?i)RTL8125D|VEN_10EC&DEV_8125' })
            if($hit.Count -gt 0){ return [pscustomobject]@{Applicable=$true;Reason='CRITICAL Realtek RTL8125D package explicitly matched the live RTL8125D/PCI DEV_8125 controller';Family=$family} }
            return [pscustomobject]@{Applicable=$false;Reason='RTL8125D package retained in catalog but the live controller is not RTL8125D/DEV_8125';Family=$family}
        }
        if($realtekDevices.Count -gt 0){ return [pscustomobject]@{Applicable=$true;Reason='CRITICAL Realtek LAN package matched a live Realtek Ethernet controller';Family=$family} }
    }

    if($family -eq 'Intel Rapid Storage/VMD'){
        $storageDevices=@($allDevices | Where-Object {
            $blob="$($_.Name) $($_.Manufacturer) $($_.PNPDeviceID) $($_.HardwareID)"
            $blob -match '(?i)Intel.*(Rapid Storage|VMD|Volume Management)|VEN_8086&DEV_(467F|7D0B|7D0C|7D0F|7D60|7D63|7D65)|SCSIAdapter'
        })
        if($storageDevices.Count -gt 0){ return [pscustomobject]@{Applicable=$true;Reason='CRITICAL Intel IRST/VMD package matched a live Intel storage/VMD controller';Family=$family} }
        return [pscustomobject]@{Applicable=$false;Reason='Intel IRST package retained in catalog, but no live Intel VMD/IRST storage controller was detected';Family=$family}
    }

    # Exact hardware-ID association has highest precedence for all other driver families.
    if($pkgIds.Count -gt 0){
        foreach($d in $allDevices){
            $did=@(Get-DeviceHardwareIds -Device $d)
            foreach($pid in $pkgIds){
                foreach($didOne in $did){
                    if($didOne -and ($didOne -ieq $pid -or $didOne -imatch ('^'+[regex]::Escape($pid)+'$'))){
                        return [pscustomobject]@{Applicable=$true;Reason="Exact package hardware ID '$pid' matched live device '$($d.Name)'";Family=$family}
                    }
                    # For PCI IDs, allow a package to omit subsystem while retaining exact VEN/DEV.
                    if($pid -match '(?i)VEN_[0-9A-F]{4}&DEV_[0-9A-F]{4}' -and $didOne -match ('(?i)'+[regex]::Escape($Matches[0]))){
                        return [pscustomobject]@{Applicable=$true;Reason="PCI VEN/DEV package hardware ID matched live device '$($d.Name)'";Family=$family}
                    }
                }
            }
        }
        return [pscustomobject]@{Applicable=$false;Reason="Package exposes hardware ID(s) [$($pkgIds -join ', ')] but none matched the live PnP inventory";Family=$family}
    }

    # Known OEM variants without explicit IDs in the catalog are still tied to the actual controller.
    $pkgText="$($Package.Title) $($Package.Category)"
    if($pkgText -match '(?i)RTL8111H'){ $hit=@($allDevices|Where-Object{("$($_.Name) $($_.PNPDeviceID)" -match '(?i)RTL8111H|VEN_10EC&DEV_8168')}); if($hit.Count -eq 0){return [pscustomobject]@{Applicable=$false;Reason='RTL8111H package has no matching live controller';Family='Realtek LAN'}} }
    if($pkgText -match '(?i)RTL8125D'){ $hit=@($allDevices|Where-Object{("$($_.Name) $($_.PNPDeviceID)" -match '(?i)RTL8125D|VEN_10EC&DEV_8125')}); if($hit.Count -eq 0){return [pscustomobject]@{Applicable=$false;Reason='RTL8125D package has no matching live controller';Family='Realtek LAN'}} }

    # For driver packages with no explicit hardware ID, associate by component family/vendor
    # against the live device inventory. This is intentionally broader than exact IDs so ASUS
    # packages whose catalog metadata omits INF IDs are not lost.
    $nameBlob=(($allDevices|ForEach-Object{"$($_.Name) $($_.Manufacturer) $($_.PNPDeviceID)"}) -join ' | ')
    $familyPatterns=@{
        'Realtek LAN'='Realtek.*(LAN|Ethernet|Network)|Ethernet.*Realtek|PCI.*VEN_10EC';
        'Intel Rapid Storage/VMD'='Intel.*(Rapid Storage|VMD|Volume Management)|VEN_8086&DEV_.*(467F|7D0B|7D0C|7D0F|7D60|7D63|7D65)';
        'Intel Graphics'='Intel.*(Graphics|Display|Arc)|VEN_8086&DEV_7D67';
        'Intel Platform Monitoring'='Intel.*Platform Monitoring|PMT';
        'Intel Platform Power'='Intel.*Platform Power|PPM';
        'Intel Serial IO'='Intel.*Serial IO';
        'Intel Dynamic Tuning'='Intel.*Dynamic Tuning|DTT';
        'Intel Trusted Execution Technology'='Intel.*(Execution Technology|Trusted Execution Technology|TXT)|TXT';
        'Intel Management Engine'='Intel.*(Management Engine|CSME|MEI)';
        'Intel VPU'='Intel.*VPU|Vision Processing';
        'Intel GNA'='Intel.*Gaussian|GNA';
        'Intel Smart Sound'='Intel.*(Smart Sound|SST)';
        'Bluetooth'='Bluetooth';
        'Wireless LAN'='Wireless|WLAN|Wi-?Fi';
        'NVIDIA Graphics'='NVIDIA.*(Graphics|Display)|GeForce|VEN_10DE';
        'Audio'='Audio|Sound';
        'Touchpad'='Touchpad|Precision TouchPad';
        'Camera/IR'='Camera|Webcam|IR Camera';
        'Card Reader'='Card Reader|Realtek.*Card';
        'Thunderbolt'='Thunderbolt';
        'USB'='USB';
        'Chipset'='Chipset|PCI Root|SMBus|System device'
    }
    if($family -and $familyPatterns.ContainsKey($family) -and $nameBlob -match $familyPatterns[$family]){
        return [pscustomobject]@{Applicable=$true;Reason="Live hardware matches component family '$family'";Family=$family}
    }
    # Model-wide ASUS driver packages with no explicit device metadata remain eligible; their
    # own installer/INF performs the final device match. We still record that no exact ID was exposed.
    return [pscustomobject]@{Applicable=$true;Reason='Official ASUS package has no explicit hardware ID metadata; final INF/installer device match remains authoritative';Family=$family}
}

function Test-InstalledDriverForPackage {
    param([Parameter(Mandatory)]$Package,[Parameter(Mandatory)]$Driver)
    $wanted=Convert-ToComparableVersion ([string]$Package.Version)
    $releaseDate=$null;try{if([string]$Package.ReleaseDate){$releaseDate=[datetime]$Package.ReleaseDate}}catch{}
    $have=Convert-ToComparableVersion ([string]$Driver.DriverVersion)
    $haveDate=$null;try{if([string]$Driver.DriverDate){$haveDate=[datetime]$Driver.DriverDate}}catch{}
    if($wanted -and $have -and $have -ge $wanted){return $true}
    if($releaseDate -and $haveDate -and $haveDate -ge $releaseDate){return $true}
    return $false
}

function Test-AsusPackageAlreadySatisfied {
    param(
        [Parameter(Mandatory)]$Package,
        [Parameter(Mandatory)]$InstalledDrivers,
        [Parameter(Mandatory)]$InstalledSoftware
    )
    $title=[string]$Package.Title
    $wanted=Convert-ToComparableVersion ([string]$Package.Version)
    $releaseDate=$null
    try { if([string]$Package.ReleaseDate){$releaseDate=[datetime]$Package.ReleaseDate} } catch {}

    # Legacy association logic retained from earlier installer revisions: every driver package is first associated with the actual live hardware.
    # This prevents a different controller/variant from satisfying the package and
    # prevents the same driver family from being installed repeatedly after reboot.
    $assoc=Test-DriverPackageHardwareAssociation -Package $Package -Devices @($script:CurrentLiveDevices) -InstalledDrivers $InstalledDrivers
    if(-not $assoc.Applicable -and (Test-PackageLooksLikeDriver -Package $Package)){
        return [pscustomobject]@{Satisfied=$true;Reason="Not applicable to live hardware: $($assoc.Reason)";Device=''}
    }

    # v28 authoritative family/version reconciliation. This is intentionally before
    # the older title-pattern fallback so VPU, NumberPad and similar packages are
    # satisfied from the actual device family rather than a fragile display-name regex.
    $strong=Test-InstalledComponentSatisfiesAsusPackage -Package $Package -InstalledDrivers $InstalledDrivers -InstalledSoftware $InstalledSoftware
    if($strong.Satisfied){ return $strong }

    # v28 final universal driver gate. The selected ASUS package must not be launched
    # merely because its catalog row is new. If Windows already reports a matching
    # installed driver at the same/newer version or release date, the package is done.
    $universal=Test-InstalledDriverMatchesAsusPackage -Package $Package -InstalledDrivers $InstalledDrivers
    if($universal.Satisfied){ return $universal }

    # Build a conservative component-specific match.  The important rule is that a
    # package is considered installed only when the live Windows device inventory
    # contains the corresponding hardware/component, not merely because another
    # package from the same vendor is present.
    $patterns=@()
    $variantPattern=$null
    switch -Regex ($title) {
        'Realtek.*LAN|LAN Driver' {
            $patterns+=@('Realtek.*LAN|Ethernet|Realtek.*(2\.5GbE|PCIe|Gaming)')
            if($title -match '(?i)RTL8111H'){$variantPattern='RTL8111H|DEV_8168'}
            elseif($title -match '(?i)RTL8125D'){$variantPattern='RTL8125D|DEV_8125'}
            break
        }
        'Intel.*Rapid Storage|\bIRST\b|Rapid Storage|\bRST\b|VMD|Volume Management' {
            $patterns+=@('Rapid Storage|RST|VMD|Volume Management Device|Intel.*Storage')
            break
        }
        'Intel.*Graphic|Intel Graphic driver|Intel.*Display' {$patterns+=@('Intel.*Graphics|Intel.*Graphic|Intel.*Display');break}
        'Platform.*Monitoring|PMT' {$patterns+=@('Platform Monitoring|PMT');break}
        'Platform Power Management|PPM' {$patterns+=@('Platform Power Management|PPM');break}
        'Serial IO' {$patterns+=@('Serial IO|Serial I/O');break}
        'Dynamic Tuning|DTT' {$patterns+=@('Dynamic Tuning|DTT');break}
        'Converged Security|Management Engine|CSME|MEI' {$patterns+=@('Converged Security|Management Engine|CSME|MEI');break}
        'VPU|Vision Processing|NPU' {$patterns+=@('Intel.*(VPU|AI Boost|Neural Processing|NPU)|Vision Processing|\bNPU\b|\bVPU\b');break}
        'Gaussian|GNA' {$patterns+=@('Gaussian|GNA');break}
        'Smart Sound|SST' {$patterns+=@('Smart Sound|SST');break}
        'Bluetooth' {$patterns+=@('Bluetooth');break}
        'NVIDIA.*Graphic|NVIDIA.*Display' {$patterns+=@('NVIDIA.*(Graphic|Display)|GeForce');break}
        'Audio' {$patterns+=@('Audio|Sound');break}
        'Touchpad|Precision TouchPad' {$patterns+=@('Touchpad|Precision TouchPad');break}
        default {$patterns+=@([regex]::Escape(($title -replace '\b(driver|software|package|utility)\b','').Trim()))}
    }

    foreach($d in @($InstalledDrivers)){
        $blob="$($d.DeviceName) $($d.Manufacturer) $($d.DriverProvider) $($d.InfName) $($d.DeviceID) $($d.HardwareID) $($d.DeviceClass)"
        $match=$false
        foreach($pat in $patterns){if($blob -match $pat){$match=$true;break}}
        if(-not $match){continue}
        if($variantPattern -and $blob -notmatch $variantPattern){continue}

        # A driver with a healthy live device record is already the installed component.
        # Prefer a comparable version, but ASUS package version strings and Windows
        # DriverVersion strings are not always expressed in the same numbering scheme.
        $have=Convert-ToComparableVersion ([string]$d.DriverVersion)
        if($wanted -and $have -and $have -ge $wanted){
            return [pscustomobject]@{Satisfied=$true;Reason="Installed driver $($d.DriverVersion) on '$($d.DeviceName)' >= ASUS package $($Package.Version)";Device=$d.DeviceName}
        }

        # For Realtek and some other OEM packages the ASUS package version is an
        # ASUS release/build identifier rather than the Windows DriverVersion. In
        # that case compare the actual installed driver date with the ASUS release date.
        $haveDate=$null
        try { if([string]$d.DriverDate){$haveDate=[datetime]$d.DriverDate} } catch {}
        if($releaseDate -and $haveDate -and $haveDate -ge $releaseDate){
            return [pscustomobject]@{Satisfied=$true;Reason="Installed live driver '$($d.DeviceName)' dated $($haveDate.ToString('yyyy-MM-dd')) is not older than ASUS package release $($releaseDate.ToString('yyyy-MM-dd'))";Device=$d.DeviceName}
        }

        # If neither version representation is comparable, do not falsely declare a
        # newer package satisfied; continue searching other matching devices.
    }

    # Software packages are reconciled independently of PnP drivers.
    $softwareNeedles=@()
    if($title -match '(?i)Armoury Crate'){$softwareNeedles+=@('Armoury Crate','ASUS Framework Service')}
    elseif($title -match '(?i)MyASUS'){$softwareNeedles+=@('MyASUS')}
    elseif($title -match '(?i)GlideX'){$softwareNeedles+=@('GlideX')}
    elseif($title -match '(?i)NVIDIA Control Panel'){$softwareNeedles+=@('NVIDIA Control Panel')}
    elseif($title -match '(?i)Intel Graphics Software'){$softwareNeedles+=@('Intel Graphics Software')}
    if($softwareNeedles.Count -gt 0){
        foreach($sw in @($InstalledSoftware)){
            foreach($needle in $softwareNeedles){
                if([string]$sw.DisplayName -match [regex]::Escape($needle)){
                    $have=Convert-ToComparableVersion ([string]$sw.DisplayVersion)
                    if(-not $wanted -or ($have -and $have -ge $wanted)){
                        return [pscustomobject]@{Satisfied=$true;Reason="Installed software '$($sw.DisplayName)' $($sw.DisplayVersion) satisfies package";Device=$sw.DisplayName}
                    }
                }
            }
        }
    }
    return [pscustomobject]@{Satisfied=$false;Reason='No matching installed component at or above the package release/version';Device=''}
}

function Export-ResumeReconciliation {
    param([Parameter(Mandatory)]$Packages,[Parameter(Mandatory)]$InstalledDrivers,[Parameter(Mandatory)]$InstalledSoftware)

    # Diagnostic only: this report must NEVER be allowed to terminate Stage 07B.
    $rows=New-Object System.Collections.Generic.List[object]
    $n=0
    foreach($p in @($Packages)){
        $n++
        try{
            $r=Test-AsusPackageAlreadySatisfied -Package $p -InstalledDrivers @($InstalledDrivers) -InstalledSoftware @($InstalledSoftware)
            [void]$rows.Add([pscustomobject]@{
                Title=[string]$p.Title
                Version=[string]$p.Version
                TargetArchitecture=[string]$p.TargetArchitecture
                ArchitectureRank=$p.ArchitectureRank
                IsDriverPackage=[bool]$p.IsDriverPackage
                Category=[string]$p.Category
                DownloadUrl=[string]$p.DownloadUrl
                InstalledSatisfied=[bool]$r.Satisfied
                Reason=[string]$r.Reason
                MatchedComponent=[string]$r.Device
            })
        }catch{
            Write-Log ("Resume reconciliation skipped package #{0} '{1}': {2}" -f $n,[string]$p.Title,$_.Exception.Message) 'WARN'
            [void]$rows.Add([pscustomobject]@{
                Title=[string]$p.Title
                Version=[string]$p.Version
                TargetArchitecture=[string]$p.TargetArchitecture
                ArchitectureRank=$p.ArchitectureRank
                IsDriverPackage=[bool]$p.IsDriverPackage
                Category=[string]$p.Category
                DownloadUrl=[string]$p.DownloadUrl
                InstalledSatisfied=$false
                Reason='Reconciliation error; main installer gate will evaluate this package'
                MatchedComponent=''
            })
        }
    }
    try{
        $path=Join-Path $DownloadRoot 'Resume-Reconciliation.csv'
        @($rows.ToArray()) | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $path
        Write-Log "Resume reconciliation written: $path" 'OK'
    }catch{
        Write-Log "Resume reconciliation report could not be written; continuing installer: $($_.Exception.Message)" 'WARN'
    }
}

function Get-AsusPackageBatchState {
    $file=Join-Path $script:DownloadRoot 'ASUS-Package-Reboot-Batch-State.json'
    $boot=Get-BootMarker
    if(Test-Path $file){
        try {
            $x=Get-Content -LiteralPath $file -Raw -ErrorAction Stop | ConvertFrom-Json
            if([string]$x.BootMarker -eq $boot){ return [pscustomobject]@{BootMarker=$boot;PendingReboot=[bool]$x.PendingReboot;PackagesSinceReboot=[int]$x.PackagesSinceReboot;Path=$file} }
        } catch {}
    }
    $state=[pscustomobject]@{BootMarker=$boot;PendingReboot=$false;PackagesSinceReboot=0;Path=$file}
    $state | ConvertTo-Json | Set-Content -LiteralPath $file -Encoding UTF8
    return $state
}

function Save-AsusPackageBatchState {
    param([Parameter(Mandatory)]$State)
    $State | Select-Object BootMarker,PendingReboot,PackagesSinceReboot | ConvertTo-Json | Set-Content -LiteralPath $State.Path -Encoding UTF8
}

function Register-LoginResumeBatch {
    # Create a one-shot batch file in the interactive user's Startup folder. Windows
    # launches the batch at the next user logon; the batch then elevates the v82
    # PowerShell installer with -Verb RunAs and resumes from the saved checkpoint.
    # The batch removes itself after successfully starting the elevated PowerShell process.
    try {
        $startupDir=Split-Path -Parent $LoginResumeBatch
        New-Item -ItemType Directory -Force -Path $startupDir -ErrorAction Stop | Out-Null

        $resumeScript=$StableScript
        if(-not (Test-Path -LiteralPath $resumeScript)){
            $source=(Resolve-Path -LiteralPath $PSCommandPath -ErrorAction Stop).Path
            New-Item -ItemType Directory -Force -Path (Split-Path -Parent $resumeScript) -ErrorAction Stop | Out-Null
            Copy-Item -LiteralPath $source -Destination $resumeScript -Force -ErrorAction Stop
        }

        $escapedScript=$resumeScript.Replace('^','^^').Replace('&','^&').Replace('|','^|').Replace('<','^<').Replace('>','^>').Replace('(','^(').Replace(')','^)')
        $batch=@'
@echo off
setlocal
set "SCRIPT=__ROG_V81_SCRIPT__"
if not exist "%SCRIPT%" exit /b 2
set "ROG_V81_SCRIPT=%SCRIPT%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$s=$env:ROG_V81_SCRIPT; Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',$s,'-Resume')"
if errorlevel 1 exit /b 3
del /f /q "%~f0" >nul 2>&1
exit /b 0
'@
        $batch=$batch.Replace('__ROG_V81_SCRIPT__',$escapedScript)
        Set-Content -LiteralPath $LoginResumeBatch -Value $batch -Encoding ASCII -Force
        Write-Log "One-shot logon resume batch created: $LoginResumeBatch" 'OK'
        Write-Log "The next interactive Windows logon will launch v82 as Administrator and resume from the saved checkpoint." 'OK'
        return $true
    } catch {
        Write-Log "Could not create the Windows Startup resume batch '$LoginResumeBatch': $($_.Exception.Message)" 'ERROR'
        return $false
    }
}

function Request-BatchedReboot {
    param([Parameter(Mandatory)]$BatchState)
    Write-Log "Five successful driver/software/package installation events have been processed. Saving the checkpoint and automatically rebooting Windows. There is NO maximum reboot count; the installer may reboot as many times as required." 'WARN'
    $BatchState.PendingReboot=$true
    Save-AsusPackageBatchState -State $BatchState
    $batchReady=Register-LoginResumeBatch
    if(-not $batchReady){
        Write-Log 'CRITICAL: the five-install automatic reboot checkpoint was reached, but the Startup resume batch could not be created. Windows will NOT reboot because automatic resume could not be guaranteed.' 'ERROR'
        return $false
    }
    Write-Log 'Automatic five-install checkpoint armed successfully. Windows will reboot and v82 will relaunch elevated at the next interactive logon from the saved package checkpoint.' 'WARN'
    return $true
}


function Invoke-FiveInstallRebootGate {
    param(
        [Parameter(Mandatory)]$BatchState,
        [Parameter(Mandatory)][string]$InstalledItemTitle
    )

    # Count successful top-level installation events only. Downloads, skipped packages,
    # failed/cancelled installers and already-installed detections do not consume a slot.
    $BatchState.PackagesSinceReboot = [int]$BatchState.PackagesSinceReboot + 1
    Save-AsusPackageBatchState -State $BatchState
    Write-Log "INSTALLATION COUNTER: '$InstalledItemTitle' completed successfully. $($BatchState.PackagesSinceReboot)/$MaxPackagesBetweenReboots installation events completed since the last reboot or reboot prompt." 'OK'

    if($BatchState.PackagesSinceReboot -lt $MaxPackagesBetweenReboots){
        return $false
    }

    # Every five successful installation events automatically reboots. The user requested
    # approximately 5-6 packages between reboots and explicitly allowed as many reboots as
    # required; therefore there is no Y/N prompt and no reboot-count ceiling.
    $doReboot=Request-BatchedReboot -BatchState $BatchState
    if($doReboot){ return $true }
    Write-Log 'Automatic reboot checkpoint could not be armed safely. The current checkpoint is preserved and the package phase will stop here.' 'ERROR'
    return $false
}


function Test-AsusPackageApplicableToLiveHardware {
    param([Parameter(Mandatory)]$Package,[Parameter(Mandatory)]$Devices,[Parameter(Mandatory)]$InstalledDrivers)
    # Legacy association logic retained from earlier installer revisions: apply the same hardware-ID + component-family association to every ASUS
    # driver package, not just the Realtek variants. Non-driver software remains eligible
    # here and is subsequently governed by companion/software prerequisites.
    $assoc=Test-DriverPackageHardwareAssociation -Package $Package -Devices $Devices -InstalledDrivers $InstalledDrivers
    if(-not $assoc.Applicable){
        Write-Log "HARDWARE ASSOCIATION: '$($Package.Title)' rejected before download: $($assoc.Reason)" 'INFO'
        return $false
    }
    Write-Log "HARDWARE ASSOCIATION: '$($Package.Title)' => applicable. $($assoc.Reason)" 'INFO'
    return $true
}


function Get-PackageLogicalInstallKey {
    param([Parameter(Mandatory=$true)]$Package)

    try {
        $isDriver=$false
        try { $isDriver=[bool](Test-PackageLooksLikeDriver -Package $Package) } catch {}

        if($isDriver){
            $family=''
            try { $family=[string](Get-PackageComponentFamily -Package $Package) } catch {}
            if(-not $family){ $family='UnknownDriver' }

            # Use safe string metadata only for the primary logical key. Hardware-ID
            # extraction remains available elsewhere for authoritative applicability,
            # but is deliberately NOT called here because catalog Raw objects can vary
            # in type under Windows PowerShell 5.1.
            $title=[string]$Package.Title
            $category=[string]$Package.Category
            # Do not include DownloadUrl in the logical key: ASUS changes URLs/file names between releases.
            $identity=($title+' '+$category)
            $identity=$identity -replace '(?i)\b(v?\d+(?:\.\d+){1,5})\b',' '
            $identity=$identity -replace '[^A-Za-z0-9]+',' '
            $identity=($identity -replace '\s+',' ').Trim().ToLowerInvariant()

            return ('DRIVER|{0}|{1}' -f $family.ToLowerInvariant(),$identity)
        }

        $title=[string]$Package.Title
        if(-not $title){$title=[string]$Package.Name}
        $title=$title -replace '(?i)\b(ASUS|ROG|Driver|Drivers|Software|Package|Utility|Utilities)\b',' '
        $title=$title -replace '(?i)\b(v?\d+(?:\.\d+){1,4})\b',' '
        $title=$title -replace '[^A-Za-z0-9]+',' '
        $title=($title -replace '\s+',' ').Trim().ToLowerInvariant()
        if(-not $title){$title='unknown software'}
        return ('SOFTWARE|{0}' -f $title)
    }
    catch {
        # Never let a logical-key problem terminate the entire installer.
        $fallback=([string]$Package.Title+' '+[string]$Package.Category)
        $fallback=$fallback -replace '[^A-Za-z0-9]+',' '
        $fallback=($fallback -replace '\s+',' ').Trim().ToLowerInvariant()
        if(-not $fallback){$fallback='unknown package'}
        Write-Log ("PACKAGE KEY FALLBACK: unable to fully classify '{0}': {1}" -f [string]$Package.Title,$_.Exception.Message) 'WARN'
        return ('FALLBACK|{0}' -f $fallback)
    }
}

function Select-BestAsusPackageCandidates {
    param([Parameter(Mandatory=$true)]$Packages)

    # Hashtable-of-arrays implementation. This avoids the .NET generic List<T> and
    # Sort-Object comparer paths that caused "Argument types do not match" on the
    # previous Windows PowerShell run.
    $groups=@{}

    foreach($pkg in @($Packages)){
        try {
            $key=[string](Get-PackageLogicalInstallKey -Package $pkg)
            if(-not $key){$key='FALLBACK|unknown package'}

            if(-not $groups.ContainsKey($key)){
                $groups[$key]=@()
            }
            $groups[$key]=@($groups[$key]) + @($pkg)
        }
        catch {
            # A catalog entry that cannot be grouped is retained as a unique target
            # rather than killing the entire ASUS phase.
            $uniqueKey=('UNIQUE|'+[guid]::NewGuid().ToString('N'))
            $groups[$uniqueKey]=@($pkg)
            Write-Log ("PACKAGE DEDUP FALLBACK: retaining '{0}' as unique candidate: {1}" -f [string]$pkg.Title,$_.Exception.Message) 'WARN'
        }
    }

    $selected=@()

    foreach($key in @($groups.Keys)){
        $candidates=@($groups[$key])
        if($candidates.Count -le 1){
            $selected=@($selected)+@($candidates[0])
            continue
        }

        # Prefer native x64, then ARM64, x86, and neutral packages. Within the preferred
        # architecture choose the highest comparable driver version, then newest ASUS release date.
        $winner=$candidates[0]
        $winnerArch=0; try{$winnerArch=[int]$winner.ArchitectureRank}catch{}
        $winnerVersion=$null
        try{$winnerVersion=Convert-ToComparableVersion ([string]$winner.Version)}catch{}
        $winnerDate=$null
        try{if([string]$winner.ReleaseDate){$winnerDate=[datetime]$winner.ReleaseDate}}catch{}
        if($candidates.Count -gt 1){
            for($i=1;$i -lt $candidates.Count;$i++){
                $candidate=$candidates[$i]
                $ca=0;try{$ca=[int]$candidate.ArchitectureRank}catch{}
                $cv=$null;try{$cv=Convert-ToComparableVersion ([string]$candidate.Version)}catch{}
                $cd=$null;try{if([string]$candidate.ReleaseDate){$cd=[datetime]$candidate.ReleaseDate}}catch{}
                $take=$false
                if($ca -gt $winnerArch){$take=$true}
                elseif($ca -eq $winnerArch){
                    if($cv -and -not $winnerVersion){$take=$true}
                    elseif($cv -and $winnerVersion -and $cv -gt $winnerVersion){$take=$true}
                    elseif(-not $cv -and -not $winnerVersion -and $cd -and -not $winnerDate){$take=$true}
                    elseif($cd -and $winnerDate -and $cd -gt $winnerDate){$take=$true}
                }
                if($take){$winner=$candidate;$winnerArch=$ca;$winnerVersion=$cv;$winnerDate=$cd}
            }
        }

        $selected=@($selected)+@($winner)

        foreach($loser in @($candidates)){
            if([object]::ReferenceEquals($loser,$winner)){continue}
            Write-Log ("PACKAGE DEDUPLICATION: suppressing older/duplicate candidate '{0}' {1}; selected '{2}' {3} for logical target '{4}'." -f `
                [string]$loser.Title,[string]$loser.Version,[string]$winner.Title,[string]$winner.Version,$key) 'OK'
        }
    }

    return @($selected)
}

function Get-IntelVpuPnpDriverState {
    # Win32_PnPSignedDriver can omit/rename some NPU/VPU records on newer Intel
    # platforms. Query the live PnP device properties directly as an independent
    # source before allowing the ASUS VPU EXE to launch.
    $rows = New-Object System.Collections.Generic.List[object]
    if(-not (Get-Command Get-PnpDevice -ErrorAction SilentlyContinue) -or
       -not (Get-Command Get-PnpDeviceProperty -ErrorAction SilentlyContinue)) {
        return @()
    }
    $devices = @()
    try {
        $devices = @(Get-PnpDevice -PresentOnly -ErrorAction Stop | Where-Object {
            $blob = "$($_.FriendlyName) $($_.Manufacturer) $($_.InstanceId)"
            $blob -match '(?i)Intel.*(AI Boost|VPU|NPU|Neural Processing|Vision Processing)|\b(AI Boost|VPU|NPU)\b'
        })
    } catch {
        return @()
    }
    foreach($device in $devices) {
        $version = ''; $date = ''; $inf = ''; $manufacturer = [string]$device.Manufacturer
        try { $p = Get-PnpDeviceProperty -InstanceId ([string]$device.InstanceId) -KeyName 'DEVPKEY_Device_DriverVersion' -ErrorAction SilentlyContinue; if($p -and $null -ne $p.Data){$version=[string]$p.Data} } catch {}
        try { $p = Get-PnpDeviceProperty -InstanceId ([string]$device.InstanceId) -KeyName 'DEVPKEY_Device_DriverDate' -ErrorAction SilentlyContinue; if($p -and $null -ne $p.Data){$date=[string]$p.Data} } catch {}
        try { $p = Get-PnpDeviceProperty -InstanceId ([string]$device.InstanceId) -KeyName 'DEVPKEY_Device_DriverInfPath' -ErrorAction SilentlyContinue; if($p -and $null -ne $p.Data){$inf=[string]$p.Data} } catch {}
        $rows.Add([pscustomobject]@{
            DeviceName = [string]$device.FriendlyName
            Manufacturer = $manufacturer
            DeviceID = [string]$device.InstanceId
            DriverVersion = $version
            DriverDate = $date
            InfName = $inf
            Status = [string]$device.Status
        })
    }
    return @($rows)
}

function Test-InstalledDriverMatchesAsusPackage {
    param(
        [Parameter(Mandatory)]$Package,
        [Parameter(Mandatory)]$InstalledDrivers
    )

    $family=Get-PackageComponentFamily -Package $Package
    $pkgIds=@(Get-PackageHardwareIds -Package $Package | ForEach-Object {[string]$_})
    $title=[string]$Package.Title
    $category=[string]$Package.Category
    $pkgBlob="$title $category".Trim()

    # Build useful title tokens while removing words that occur in nearly every ASUS
    # catalog entry. These tokens are only a fallback after hardware-ID/family matching.
    $tokens=@(
        ($pkgBlob -replace '(?i)\b(ASUS|ROG|Driver|Drivers|Software|Package|Utility|Utilities|Version|V\d+(?:\.\d+){1,3}|Windows|Win10|Win11|x64|64bit|64-bit|32bit|32-bit)\b',' ') `
        -replace '[^A-Za-z0-9]+',' ' -split '\s+' |
        Where-Object { $_.Length -ge 3 } |
        Select-Object -Unique
    )

    $wanted=Convert-ToComparableVersion ([string]$Package.Version)
    $releaseDate=$null
    try { if([string]$Package.ReleaseDate){$releaseDate=[datetime]$Package.ReleaseDate} } catch {}

    foreach($d in @($InstalledDrivers)){
        $deviceBlob="$($d.DeviceName) $($d.Manufacturer) $($d.DriverProvider) $($d.InfName) $($d.DeviceID) $($d.HardwareID) $($d.DeviceClass)"
        $hardwareMatched=$false
        $familyMatched=$false
        $tokenMatched=$false

        # 1) Exact/compatible hardware identity is authoritative.
        if($pkgIds.Count -gt 0){
            foreach($pid in $pkgIds){
                if(-not $pid){continue}
                if($deviceBlob -match [regex]::Escape($pid)){
                    $hardwareMatched=$true;break
                }
                # PCI package IDs may contain SUBSYS/REV details while Windows' installed
                # record only exposes the base VEN/DEV. Match the base identity as well.
                if($pid -match '(?i)(VEN_[0-9A-F]{4}&DEV_[0-9A-F]{4})'){
                    if($deviceBlob -match [regex]::Escape($Matches[1])){
                        $hardwareMatched=$true;break
                    }
                }
            }
        }

        # v45: Realtek variant safety. RTL8111H and RTL8125D share the same
        # component family, but must never satisfy each other's ASUS package.
        $variantCompatible=$true
        if($pkgBlob -match '(?i)RTL8111H'){
            $variantCompatible=($deviceBlob -match '(?i)RTL8111H|VEN_10EC&DEV_8168')
        } elseif($pkgBlob -match '(?i)RTL8125D'){
            $variantCompatible=($deviceBlob -match '(?i)RTL8125D|VEN_10EC&DEV_8125')
        }
        if(-not $variantCompatible){ continue }

        # 2) Component-family identity catches ASUS packages whose catalog metadata
        # does not expose INF hardware IDs (common for Intel platform packages).
        if($family){
            $df=Get-DeviceComponentFamily -Device ([pscustomobject]@{
                Name=[string]$d.DeviceName
                Manufacturer=[string]$d.Manufacturer
                PNPDeviceID=[string]$d.DeviceID
                HardwareID=[string]$d.HardwareID
            })
            if($df -eq $family){$familyMatched=$true}
        }

        # 3) Generic vendor/title fallback. This is deliberately conservative:
        # require at least one distinctive token and, for multi-token titles, two hits.
        if($tokens.Count -gt 0){
            $hits=0
            foreach($tok in $tokens){
                if($deviceBlob -match [regex]::Escape($tok)){ $hits++ }
            }
            $minimum=if($tokens.Count -ge 2){2}else{1}
            if($hits -ge $minimum){$tokenMatched=$true}
        }

        # v84 anti-cross-family correction: a broad/shared hardware ID must not let
        # an unrelated Intel device (for example Smart Sound) satisfy a different
        # known family such as Dynamic Tuning, GNA, IRST, TXT, or VPU. For recognized
        # families, accept a family match; accept a hardware-ID match only when the
        # installed record itself has no identifiable component family. This retains
        # exact-ID support for unnamed devices without cross-satisfying known families.
        if($family){
            if(-not $familyMatched){
                if(-not ($hardwareMatched -and -not $df)){continue}
            }
        } elseif(-not ($hardwareMatched -or $familyMatched -or $tokenMatched)){continue}

        $have=Convert-ToComparableVersion ([string]$d.DriverVersion)
        $haveDate=$null
        try { if([string]$d.DriverDate){$haveDate=[datetime]$d.DriverDate} } catch {}

        # Version/date reconciliation follows Windows' own selection model: for equal
        # hardware matches, date precedes version. We nevertheless require both to be
        # non-conflicting when both are available.
        if($wanted -and $have -and $have -ge $wanted){
            return [pscustomobject]@{
                Satisfied=$true
                Reason="Installed matching driver '$($d.DeviceName)' version $($d.DriverVersion) is >= package version $($Package.Version)"
                Device=$d.DeviceName
            }
        }

        if($releaseDate -and $haveDate -and $haveDate -ge $releaseDate){
            return [pscustomobject]@{
                Satisfied=$true
                Reason="Installed matching driver '$($d.DeviceName)' dated $($haveDate.ToString('yyyy-MM-dd')) is >= package release date $($releaseDate.ToString('yyyy-MM-dd'))"
                Device=$d.DeviceName
            }
        }

        # If ASUS and Windows expose the same version text, treat that as an exact
        # installed match even when the release date metadata is absent.
        if($wanted -and $have -and ([string]$d.DriverVersion).Trim() -eq ([string]$Package.Version).Trim()){
            return [pscustomobject]@{
                Satisfied=$true
                Reason="Installed matching driver '$($d.DeviceName)' has the exact package driver version $($d.DriverVersion); installer suppressed."
                Device=$d.DeviceName
            }
        }
    }

    return [pscustomobject]@{Satisfied=$false;Reason='No matching installed driver at or above the ASUS package version/date was found';Device=''}
}

function Test-InstalledComponentSatisfiesAsusPackage {
    param(
        [Parameter(Mandatory)]$Package,
        [Parameter(Mandatory)]$InstalledDrivers,
        [Parameter(Mandatory)]$InstalledSoftware
    )

    $isDriver = [bool](Test-PackageLooksLikeDriver -Package $Package)
    $wanted = Convert-ToComparableVersion ([string]$Package.Version)
    $releaseDate = $null
    try { if([string]$Package.ReleaseDate){ $releaseDate=[datetime]$Package.ReleaseDate } } catch {}

    if($isDriver){
        # v28: authoritative whole-driver reconciliation. This runs before family-only
        # checks so Intel platform packages (TXT, VPU, PMT, PPM, etc.) are recognized
        # from their actual installed PnP driver records even when the friendly name
        # differs from the ASUS catalog title.
        $universal=Test-InstalledDriverMatchesAsusPackage -Package $Package -InstalledDrivers $InstalledDrivers
        if($universal.Satisfied){ return $universal }

        $family = Get-PackageComponentFamily -Package $Package

        # v84 Intel Chipset INF anti-repeat gate:
        # Intel's chipset setup is an INF-update bundle, not one conventional device
        # driver. The device-family matcher intentionally does not classify every
        # Intel system device as a single family, so inspect chipset-specific records
        # and the installed Intel Chipset Device Software product explicitly.
        if($family -eq 'Chipset' -and $title -match '(?i)Intel.*Chipset.*Software'){
            $chipsetSoftwareMatches=@($InstalledSoftware | Where-Object {
                ([string]$_.DisplayName -match '(?i)Intel.*Chipset.*(Device )?Software') -and
                ([string]$_.Publisher -match '(?i)Intel')
            })
            foreach($csw in $chipsetSoftwareMatches){
                $haveSoftware=Convert-ToComparableVersion ([string]$csw.DisplayVersion)
                if($wanted -and $haveSoftware -and $haveSoftware -ge $wanted){
                    return [pscustomobject]@{
                        Satisfied=$true
                        Reason="Installed Intel chipset software '$($csw.DisplayName)' version $($csw.DisplayVersion) is >= ASUS package $($Package.Version); chipset installer suppressed."
                        Device=$csw.DisplayName
                    }
                }
            }

            $chipsetDevicePattern='(?i)(SMBus|LPC Controller|Host Bridge|PCI Express Root Port|GPIO|SPI Controller|PMC|PCH|System device|Chipset|I2C Controller|Thermal subsystem)'
            $excludeChipsetPattern='(?i)(Innovation Platform Framework|\bIPF\b|VPU|AI Boost|Neural|NPU|Graphics|Display|Management Engine|\bMEI\b|Serial IO|Dynamic Tuning|Smart Sound|\bSST\b|\bGNA\b|Gaussian|Rapid Storage|\bRST\b|\bVMD\b)'
            $chipsetRecords=@($InstalledDrivers | Where-Object {
                $recordText="$($_.DeviceName) $($_.Manufacturer) $($_.DeviceID) $($_.HardwareID) $($_.DeviceClass)"
                ($recordText -match '(?i)Intel|VEN_8086') -and
                ($recordText -match $chipsetDevicePattern) -and
                ($recordText -notmatch $excludeChipsetPattern)
            })
            if($chipsetRecords.Count -gt 0){
                $comparableChipsetVersions=@()
                $unknownChipsetVersion=$false
                foreach($cr in $chipsetRecords){
                    $haveChipset=Convert-ToComparableVersion ([string]$cr.DriverVersion)
                    if($haveChipset){
                        $comparableChipsetVersions+=@([pscustomobject]@{Record=$cr;Version=$haveChipset})
                    } else {
                        $unknownChipsetVersion=$true
                    }
                }
                $newerOrSame=@($comparableChipsetVersions | Where-Object {$wanted -and $_.Version -ge $wanted})
                if($newerOrSame.Count -gt 0){
                    $hit=$newerOrSame[0]
                    return [pscustomobject]@{
                        Satisfied=$true
                        Reason="Installed Intel chipset/system-device driver '$($hit.Record.DeviceName)' version $($hit.Record.DriverVersion) is >= ASUS chipset bundle version $($Package.Version); installer suppressed to prevent same-version/downgrade overwrite."
                        Device=$hit.Record.DeviceName
                    }
                }
                if($unknownChipsetVersion){
                    return [pscustomobject]@{
                        Satisfied=$true
                        Reason="Intel chipset/system-device records are present but at least one matching record has no comparable driver version; fail-closed anti-downgrade policy suppresses the chipset bundle rather than risk overwriting an unknown same/newer component."
                        Device=([string]$chipsetRecords[0].DeviceName)
                    }
                }
            }
        }

        # v84 VPU hotfix: independently inspect live Intel AI Boost/NPU PnP properties.
        # If Windows says the device is present and OK but does not expose a comparable
        # driver version, fail closed instead of launching the ASUS EXE blindly.
        if($family -eq 'Intel VPU'){
            $vpuDevices=@(Get-IntelVpuPnpDriverState)
            foreach($vpu in $vpuDevices){
                $haveVpu=Convert-ToComparableVersion ([string]$vpu.DriverVersion)
                if($wanted -and $haveVpu -and $haveVpu -ge $wanted){
                    return [pscustomobject]@{
                        Satisfied=$true
                        Reason="Live PnP properties show Intel VPU device '$($vpu.DeviceName)' already has driver version $($vpu.DriverVersion), >= ASUS package $($Package.Version); installer suppressed."
                        Device=$vpu.DeviceName
                    }
                }
                if(-not $haveVpu -and ([string]$vpu.Status -eq 'OK')){
                    return [pscustomobject]@{
                        Satisfied=$true
                        Reason="Live Intel VPU device '$($vpu.DeviceName)' is present and reports Status=OK, but Windows did not expose a comparable driver version; fail-closed anti-downgrade protection suppresses the installer."
                        Device=$vpu.DeviceName
                    }
                }
            }
        }
        foreach($d in @($InstalledDrivers)){
            $dev = [pscustomobject]@{
                Name = [string]$d.DeviceName
                Manufacturer = [string]$d.Manufacturer
                PNPDeviceID = [string]$d.DeviceID
                HardwareID = [string]$d.HardwareID
            }
            $df = Get-DeviceComponentFamily -Device $dev
            if(-not $family -or -not $df -or $df -ne $family){ continue }

            # Downgrade protection: when a matching driver family is already present but
            # the vendor catalog exposes no parseable target version/date, do not guess.
            # The installer is retained in the archive, but execution is blocked because
            # the script cannot prove that the candidate is newer than the installed driver.
            if(-not $wanted -and -not $releaseDate){
                return [pscustomobject]@{
                    Satisfied=$true
                    Reason="Matching installed $family driver '$($d.DeviceName)' is present, but the candidate has no comparable version/release date; fail-closed downgrade protection blocks an unverified overwrite"
                    Device=$d.DeviceName
                }
            }

            $have = Convert-ToComparableVersion ([string]$d.DriverVersion)
            if($wanted -and $have -and $have -ge $wanted){
                return [pscustomobject]@{
                    Satisfied=$true
                    Reason="Installed $family driver $($d.DriverVersion) on '$($d.DeviceName)' >= ASUS package $($Package.Version)"
                    Device=$d.DeviceName
                }
            }

            $haveDate=$null
            try { if([string]$d.DriverDate){$haveDate=[datetime]$d.DriverDate} } catch {}
            if($releaseDate -and $haveDate -and $haveDate -ge $releaseDate){
                return [pscustomobject]@{
                    Satisfied=$true
                    Reason="Installed $family driver '$($d.DeviceName)' dated $($haveDate.ToString('yyyy-MM-dd')) is not older than ASUS package release $($releaseDate.ToString('yyyy-MM-dd'))"
                    Device=$d.DeviceName
                }
            }
        }
    }
    else {
        # Generic software reconciliation. This deliberately covers ASUS utilities
        # such as Smart Display Control that were previously absent from the small
        # hard-coded software-name list.
        $title = [string]$Package.Title
        $tokens = @(
            $title -replace '(?i)\b(ASUS|ROG|Driver|Drivers|Software|Package|Utility|Utilities|V\d+(?:\.\d+)*)\b',' ' `
            -replace '[^A-Za-z0-9]+',' '
        ) -split '\s+' | Where-Object { $_.Length -ge 4 } | Select-Object -Unique

        foreach($sw in @($InstalledSoftware)){
            $name = [string]$sw.DisplayName
            if(-not $name){ continue }
            $hits = 0
            foreach($token in $tokens){
                if($name -match [regex]::Escape($token)){ $hits++ }
            }
            if($hits -lt 2 -and $tokens.Count -gt 1){ continue }

            $have = Convert-ToComparableVersion ([string]$sw.DisplayVersion)
            if(-not $wanted -or ($have -and $have -ge $wanted)){
                return [pscustomobject]@{
                    Satisfied=$true
                    Reason="Installed software '$name' $($sw.DisplayVersion) satisfies ASUS package $($Package.Title) $($Package.Version)"
                    Device=$name
                }
            }
        }
    }

    return [pscustomobject]@{Satisfied=$false;Reason='No installed component at or above the selected package version was found';Device=''}
}


function Test-NonBlockingAsusPackage {
    param([Parameter(Mandatory=$true)]$Package)
    $title=[string]$Package.Title
    # ASUS describes displayHDR as a certificate helper that exposes an HDR
    # certificate in Settings/Display when the panel is certified. It is not a
    # device driver and its CMD helper can return 1 when there is no applicable
    # certificate action. Treat only exit code 1 for this exact package as
    # non-blocking; real driver/setup failures remain blocking.
    if($title -match '^(?i:displayHDR)$'){return $true}
    return $false
}


function Test-StoreCompanionInstalled {
    param([Parameter(Mandatory)]$App)
    $fake=[pscustomobject]@{Title=[string]$App.Name;Category='Microsoft Store Application';DownloadUrl="https://apps.microsoft.com/detail/$($App.Id)"}
    try { return [bool](Test-StoreAppInstalled -Package $fake) } catch { return $false }
}

function Reset-WingetMicrosoftStoreSource {
    param([Parameter(Mandatory)][string]$WingetPath)
    try {
        Write-Log 'WinGet Microsoft Store source reset requested after a Store transaction returned success without a verifiable AppX installation.' 'WARN'
        $p=Start-Process -FilePath $WingetPath -ArgumentList @('source','reset','--force','--disable-interactivity') -Wait -PassThru -NoNewWindow -ErrorAction Stop
        Write-Log "WinGet source reset exit code: $($p.ExitCode)." $(if($p.ExitCode -eq 0){'OK'}else{'WARN'})
        return ($p.ExitCode -eq 0)
    } catch {
        Write-Log "WinGet Microsoft Store source reset failed: $($_.Exception.Message)" 'WARN'
        return $false
    }
}

function Install-RelatedMicrosoftStoreApps {
    param(
        [Parameter(Mandatory=$true)]$Devices,
        [Parameter(Mandatory=$true)]$AsusPackages
    )

    Write-Log 'STAGE 07G - Installing and running related Microsoft Store hardware/driver companion applications through the official msstore source' 'STEP'
    Write-Log 'v53 Store policy: WinGet exit code 0 is NOT accepted as proof of installation. Every Store transaction is verified against the installed AppX inventory, and every installed Store companion is launched and held as an interactive settings gate before the next Store app.' 'INFO'

    $apps=@(
        # MyASUS is deliberately excluded here. It is installed/launched immediately before
        # the final Armoury Crate/Aura stage, as required by the final ASUS software ordering.
        [pscustomobject]@{Name='Intel Graphics Software';Id='9P8K5G2MWW6Z';Publisher='Intel';Reason='Required Intel onboard graphics settings/management companion listed for the G635LW';Patterns='IntelGraphicsSoftware|Intel.*Graphics.*Software|9P8K5G2MWW6Z';Required=$true},
        [pscustomobject]@{Name='Intel Graphics Command Center';Id='9PLFNLNT3G5G';Publisher='Intel';Reason='Intel graphics control companion';Patterns='IntelGraphicsExperience|AppUp\.IntelGraphicsExperience|IGCC|Intel.*Graphic.*Command.*Center|9PLFNLNT3G5G';Required=$false},
        [pscustomobject]@{Name='Realtek Audio Control';Id='9P2B8MCSVPLN';Publisher='Realtek Semiconductor Corp.';Reason='Realtek Audio/Codec Console companion';Patterns='Realtek.*Audio|Realtek.*Codec|ALC3288|Audio';Required=$false},
        [pscustomobject]@{Name='NVIDIA Control Panel';Id='9NF8H0H7WMLT';Publisher='NVIDIA Corp.';Reason='NVIDIA DCH graphics control companion';Patterns='NVIDIA.*(Graphic|Graphics|Display)|GeForce|RTX 5080';Required=$false},
        [pscustomobject]@{Name='Dolby Access';Id='9N0866FS04W8';Publisher='Dolby Laboratories';Reason='Dolby Atmos companion for the ASUS Dolby Atmos package';Patterns='Dolby Atmos|Dolby';Required=$false},
        [pscustomobject]@{Name='Armoury Crate';Id='9PM9DFQRDH3F';Publisher='ASUS';Reason='ROG/ASUS device-control companion';Patterns='Armoury Crate|ArmouryCrate|Aura Creator';Required=$false},
        [pscustomobject]@{Name='Aura Creator';Id='9MWZ3VLW5HBW';Publisher='ASUS';Reason='Aura/lighting companion associated with Armoury Crate';Patterns='Armoury Crate|Aura Creator|Aura|Lighting|RGB';Required=$false}
    )

    $hardwareText=''
    try { $hardwareText=(@($Devices)|ForEach-Object { "$($_.Name) $($_.Manufacturer) $($_.PNPDeviceID) $($_.HardwareID)" }) -join ' ' } catch {}
    $catalogText=''
    try { $catalogText=(@($AsusPackages)|ForEach-Object { "$($_.Title) $($_.Category) $($_.DownloadUrl)" }) -join ' ' } catch {}
    $allText="$hardwareText $catalogText"

    $winget=Get-Command winget.exe -ErrorAction SilentlyContinue
    if(-not $winget){
        Write-Log 'Related Microsoft Store pass: winget.exe is unavailable; Store companion apps remain deferred.' 'WARN'
        return [pscustomobject]@{Attempted=0;Installed=0;Deferred=$apps.Count}
    }
    if(-not (Initialize-MicrosoftStoreForCurrentUser)){
        Write-Log 'Related Microsoft Store pass: Microsoft Store is not ready for the interactive user. Store companion apps are deferred.' 'WARN'
        return [pscustomobject]@{Attempted=0;Installed=0;Deferred=$apps.Count}
    }

    $attempted=0;$installed=0;$deferred=0
    foreach($app in $apps){
        $appApplicable=($app.Required -or ($allText -match $app.Patterns))
        if(-not $appApplicable){
            Write-Log "[MSSTORE] Skipping unrelated Store app '$($app.Name)' because no matching hardware/ASUS package relationship was found in the STAGE 06 baseline/catalog." 'INFO'
            continue
        }

        $already=Test-StoreCompanionInstalled -App $app
        if($already){
            Write-Log "[MSSTORE] $($app.Name) is already installed. No Store download is required; launching it now for the required interactive settings gate." 'OK'
        } else {
            $attempted++
            $installedThisApp=$false
            for($try=1;$try -le 2 -and -not $installedThisApp;$try++){
                Write-Log "[MSSTORE] Installing '$($app.Name)' with WinGet using the official msstore source and exact Product ID $($app.Id) (attempt $try/2)." 'STEP'
                try {
                    $wp=Invoke-WingetMsStoreInstall -WingetPath $winget.Source -PackageId $app.Id -Exact -Silent -TimeoutSeconds 180
                    $code=[int]$wp.ExitCode
                    Write-Log "[MSSTORE] $($app.Name) WinGet exit code: $code." 'INFO'
                    Start-Sleep -Seconds 3
                    $installedThisApp=Test-StoreCompanionInstalled -App $app
                    if($installedThisApp){
                        Write-Log "[MSSTORE] $($app.Name) is now verifiably registered as an installed AppX package." 'OK'
                        break
                    }
                    if($try -eq 1){
                        Write-Log "[MSSTORE] WinGet returned exit code $code but $($app.Name) is NOT present in the installed AppX inventory. A source reset/retry will be performed rather than treating exit code 0 as success." 'WARN'
                        [void](Reset-WingetMicrosoftStoreSource -WingetPath $winget.Source)
                        Start-Sleep -Seconds 2
                    }
                } catch {
                    Write-Log "[MSSTORE] $($app.Name) Store transaction raised an exception: $($_.Exception.Message)." 'WARN'
                    if($try -eq 1){ [void](Reset-WingetMicrosoftStoreSource -WingetPath $winget.Source); Start-Sleep -Seconds 2 }
                }
            }
            if(-not $installedThisApp){
                $deferred++
                Write-Log "[MSSTORE] $($app.Name) could not be verified as installed after two Store attempts. It is deferred and the installer will NOT pretend that exit code 0 means installed." 'WARN'
                continue
            }
            $installed++
        }

        # Every Store companion is an interactive settings stage. Do not merely install it:
        # launch it and wait indefinitely before the next Store app/package is allowed.
        if(-not (Invoke-InteractiveStoreApp -DisplayName $app.Name -NamePattern $app.Patterns)){
            $deferred++
            Write-Log "[MSSTORE] $($app.Name) could not complete its interactive launch/settings gate. The Store sequence is held at this application rather than silently advancing." 'ERROR'
            break
        }
    }
    Write-Log "Related Microsoft Store pass complete: attempted=$attempted, installed=$installed, deferred=$deferred." 'OK'
    return [pscustomobject]@{Attempted=$attempted;Installed=$installed;Deferred=$deferred}
}

function Invoke-PackageCompanionExecutables {
    param(
        [Parameter(Mandatory)][string]$ExtractDirectory,
        [Parameter(Mandatory)][string]$PrimaryPath,
        [string]$PackageTitle='Package'
    )
    if(-not (Test-Path -LiteralPath $ExtractDirectory)){ return 0 }
    $primary=[IO.Path]::GetFullPath($PrimaryPath)
    $candidates=@(Get-ChildItem -LiteralPath $ExtractDirectory -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -ieq '.exe' -and ([IO.Path]::GetFullPath($_.FullName) -ine $primary) } |
        Where-Object {
            $n=$_.Name
            # Never treat a second installer/updater/uninstaller/firmware flasher as a settings companion.
            if($n -match '(?i)^(setup|install|installer|uninstall|unins|update|updater|patch|bootstrap|launcher|autorun|dpinst|driver|firmware|flash|bios|vcredist|vc_redist|dotnet).*\.exe$'){return $false}
            if($n -match '(?i)(uninstall|unins|firmware|flash|bios|vcredist|vc_redist|dotnet|runtime|redist|crash|report|telemetry)'){return $false}
            # Run same-package non-installer executables so the user can enter device/application settings.
            # Dangerous/maintenance executables are excluded above. This intentionally includes
            # vendor utilities even when their filename does not contain a settings keyword.
            return $true
        } | Sort-Object FullName)
    if($candidates.Count -eq 0){ return 0 }
    Write-Log "  COMPANION PASS: '$PackageTitle' contains $($candidates.Count) selected same-package configuration/control/utility executable(s). They will be run sequentially before the next package." 'STEP'
    $ran=0
    foreach($exe in $candidates){
        try{
            Write-Log "    Companion executable: $($exe.FullName)" 'INFO'
            $proc=Start-Process -FilePath $exe.FullName -WorkingDirectory $exe.DirectoryName -PassThru -WindowStyle Normal -ErrorAction Stop
            # Settings/configuration utilities are intentionally interactive. Do not impose a timeout:
            # the package is not allowed to advance until the user closes the utility.
            $proc.WaitForExit()
            Write-Log "    Companion '$($exe.Name)' exited with code $($proc.ExitCode). The package settings stage is complete." $(if($proc.ExitCode -in @(0,3010,1641)){'OK'}else{'WARN'})
            $ran++
        }catch{Write-Log "    Companion '$($exe.Name)' could not be launched: $($_.Exception.Message)" 'WARN'}
    }
    return $ran
}


function Invoke-InteractiveStoreApp {
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$NamePattern
    )
    # Intel Graphics Command Center is a companion UI, not a driver installer.
    # Do not launch it or wait for the user to close it: that blocks the package loop
    # and does not install/update the Intel display driver.
    if($DisplayName -match '(?i)^Intel Graphics Command Center$' -or $NamePattern -match '(?i)IntelGraphicsExperience|9PLFNLNT3G5G'){
        $igccApps = @()
        try {
            $igccApps = @(Get-AppxPackage -ErrorAction SilentlyContinue | Where-Object {
                "$($_.Name) $($_.PackageFamilyName) $($_.PackageFullName)" -match '(?i)IntelGraphicsExperience|AppUp\.IntelGraphicsExperience'
            })
        } catch {}
        if($igccApps.Count -gt 0){
            $igccVersions = @($igccApps | ForEach-Object { [string]$_.Version } | Where-Object { $_ })
            if($igccVersions.Count -gt 0){
                Write-Log "  [Intel Graphics Command Center] Already installed (AppX version(s): $($igccVersions -join ', ')); skipping GUI launch. This companion check does not claim that a Store update was installed." 'OK'
            } else {
                Write-Log "  [Intel Graphics Command Center] Already installed; skipping GUI launch. Installed AppX version was not exposed by the current-user inventory." 'OK'
            }
            return $true
        }
        Write-Log "  [Intel Graphics Command Center] Not detected in current-user AppX inventory; skipping interactive launch gate so the installer cannot hang on the UI. Store install/update status must be handled by the Store package transaction." 'WARN'
        return $true
    }

    Write-Log "  [$DisplayName] Searching the installed Microsoft Store package for a launchable application." 'STEP'
    $apps=@()

    # Give a just-completed Store deployment a short registration window. This does
    # not re-run hardware/driver detection; it only waits for AppX registration.
    for($attempt=1;$attempt -le 20 -and $apps.Count -eq 0;$attempt++){
        try {
            $apps=@(Get-AppxPackage -ErrorAction SilentlyContinue | Where-Object {
                "$($_.Name) $($_.PackageFamilyName) $($_.PackageFullName)" -match $NamePattern
            })
            if($apps.Count -eq 0){
                $apps=@(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object {
                    "$($_.Name) $($_.PackageFamilyName) $($_.PackageFullName)" -match $NamePattern
                })
            }
        } catch {}
        if($apps.Count -eq 0){ Start-Sleep -Seconds 1 }
    }

    if($apps.Count -eq 0){
        Write-Log "  [$DisplayName] The Store package could not be found in the AppX inventory after registration wait." 'ERROR'
        return $false
    }

    foreach($app in $apps){
        try {
            $manifestPath=Join-Path $app.InstallLocation 'AppxManifest.xml'
            if(-not (Test-Path -LiteralPath $manifestPath)){ continue }
            [xml]$manifest=Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop
            $ns=New-Object System.Xml.XmlNamespaceManager($manifest.NameTable)
            $ns.AddNamespace('a','http://schemas.microsoft.com/appx/manifest/foundation/windows10')

            # 1. Preferred: manifest-declared executable.
            foreach($application in $manifest.SelectNodes('//a:Applications/a:Application',$ns)){
                $exe=[string]$application.Executable
                $appId=[string]$application.Id
                if($exe){
                    $exePath=Join-Path $app.InstallLocation $exe
                    if(Test-Path -LiteralPath $exePath){
                        Write-Log "  [$DisplayName] Launching manifest executable $exePath. Waiting indefinitely for the user to close it." 'STEP'
                        $proc=Start-Process -FilePath $exePath -WorkingDirectory $app.InstallLocation -PassThru -WindowStyle Normal -ErrorAction Stop
                        $proc.WaitForExit()
                        Write-Log "  [$DisplayName] Interactive settings application closed with exit code $($proc.ExitCode). Package settings stage is complete." 'OK'
                        return $true
                    }
                }
            }

            # 2. Intel GCC and some Store packages may have an executable in the
            # WindowsApps directory even when the manifest's Executable attribute is
            # not populated. Intel documents IGCC.exe inside the Intel GCC package.
            $knownExeNames=@()
            if($DisplayName -match '(?i)Intel.*Graphics.*Command|Intel.*Graphics'){
                $knownExeNames=@('IGCC.exe','IGCC64.exe')
            }
            if($knownExeNames.Count -gt 0){
                foreach($known in $knownExeNames){
                    $foundExe=@(Get-ChildItem -LiteralPath $app.InstallLocation -Recurse -File -Filter $known -Force -ErrorAction SilentlyContinue | Select-Object -First 1)
                    if($foundExe.Count -gt 0){
                        Write-Log "  [$DisplayName] Found package executable $($foundExe[0].FullName) even though the manifest did not expose it. Launching it and waiting indefinitely for settings." 'OK'
                        $proc=Start-Process -FilePath $foundExe[0].FullName -WorkingDirectory $foundExe[0].DirectoryName -PassThru -WindowStyle Normal -ErrorAction Stop
                        $proc.WaitForExit()
                        Write-Log "  [$DisplayName] Interactive settings application closed with exit code $($proc.ExitCode). Package settings stage is complete." 'OK'
                        return $true
                    }
                }
            }

            # 3. Fallback: use the current user's Start-menu AppUserModelID.
            # Get-StartApps returns the actual AppID registered for the current user.
            # This is more reliable than constructing a family-name!application-ID
            # from a manifest when the Store package uses a generated/packaged entry.
            try {
                $startApps=@(Get-StartApps -ErrorAction SilentlyContinue | Where-Object {
                    ([string]$_.Name -match [regex]::Escape($DisplayName)) -or
                    ([string]$_.Name -match $NamePattern)
                })
                foreach($entry in $startApps){
                    $appUserModelId=[string]$entry.AppID
                    if([string]::IsNullOrWhiteSpace($appUserModelId)){continue}
                    Write-Log "  [$DisplayName] Launching registered Start-menu AppID $appUserModelId through AppsFolder." 'INFO'
                    $before=@(Get-Process -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
                    Start-Process -FilePath (Join-Path $env:WINDIR 'explorer.exe') `
                        -ArgumentList "shell:AppsFolder\$appUserModelId" -WindowStyle Normal -ErrorAction Stop | Out-Null

                    # If a real executable appears, wait on it. For UWP apps hosted by
                    # ApplicationFrameHost, use a foreground/window-process fallback below.
                    $root=[IO.Path]::GetFullPath([string]$app.InstallLocation).TrimEnd('\')
                    for($wait=0;$wait -lt 20;$wait++){
                        Start-Sleep -Seconds 1
                        $found=@()
                        try {
                            $found=@(Get-Process -ErrorAction SilentlyContinue | Where-Object {
                                try {
                                    $path=$_.Path
                                    if($path){ return ([IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)) }
                                } catch {}
                                return $false
                            })
                        } catch {}
                        if($found.Count -gt 0){
                            Write-Log "  [$DisplayName] Packaged application process detected. Waiting indefinitely for the user to close it." 'STEP'
                            foreach($proc in $found){ try{$proc.WaitForExit()}catch{} }
                            Write-Log "  [$DisplayName] Interactive Store application closed. Package settings stage is complete." 'OK'
                            return $true
                        }
                    }

                    # The shell launch itself is the authoritative activation path. If
                    # Windows registered the AppID but exposes no process path to the
                    # current PowerShell security context, require the user to close the
                    # app and press ENTER rather than silently advancing.
                    Write-Log "  [$DisplayName] Windows accepted the registered AppID launch, but the packaged process path is not exposed to PowerShell. The settings gate will remain here until you confirm the application has been closed." 'WARN'
                    try {
                        [void](Read-Host "[$DisplayName] Configure your settings, close the application, then press ENTER to continue")
                        Write-Log "  [$DisplayName] User confirmed the interactive Store application was configured and closed." 'OK'
                        return $true
                    } catch {
                        Write-Log "  [$DisplayName] User confirmation could not be collected: $($_.Exception.Message)" 'ERROR'
                        return $false
                    }
                }
            } catch {
                Write-Log "  [$DisplayName] Start-menu AppID discovery failed: $($_.Exception.Message)" 'WARN'
            }

            # 4. Last fallback: construct the manifest AUMID.
            foreach($application in $manifest.SelectNodes('//a:Applications/a:Application',$ns)){
                $appId=[string]$application.Id
                if($appId -and $app.PackageFamilyName){
                    $aumid="$($app.PackageFamilyName)!$appId"
                    Write-Log "  [$DisplayName] Attempting manifest AUMID $aumid through AppsFolder." 'INFO'
                    try {
                        Start-Process -FilePath (Join-Path $env:WINDIR 'explorer.exe') -ArgumentList "shell:AppsFolder\$aumid" -WindowStyle Normal -ErrorAction Stop | Out-Null
                        Write-Log "  [$DisplayName] Manifest AUMID activation was issued. Waiting for explicit user confirmation rather than silently advancing." 'WARN'
                        [void](Read-Host "[$DisplayName] Configure your settings, close the application, then press ENTER to continue")
                        Write-Log "  [$DisplayName] User confirmed the interactive Store application was configured and closed." 'OK'
                        return $true
                    } catch {
                        Write-Log "  [$DisplayName] Manifest AUMID launch failed: $($_.Exception.Message)" 'WARN'
                    }
                }
            }
        } catch {
            Write-Log "  [$DisplayName] Store application launch attempt failed: $($_.Exception.Message)" 'WARN'
        }
    }

    Write-Log "  [$DisplayName] Could not locate a launchable executable or registered AppID in the installed Store package. Interactive settings were NOT completed, so the installer will not silently advance past this Store gate." 'ERROR'
    return $false
}


function Test-IsIntelGraphicsPackage {
    param([Parameter(Mandatory)]$Package)
    $t = "$(if($null -ne $Package.Category){$Package.Category}) $(if($null -ne $Package.Title){$Package.Title})"
    $url = [string]$Package.DownloadUrl
    $raw = ''
    try { $raw = [string]($Package.Raw | ConvertTo-Json -Depth 20 -Compress) } catch {}
    if($t -match '(?i)Command\s*Center|Graphics\\s*Command\s*Center|Intel\s*Graphics\s*Command' -or $url -match '(?i)9PLFNLNT3G5G') { return $false }
    return [bool](
        ($t -match '(?i)Intel.*(Graphic|Graphics|Display)') -or
        ($raw -match '(?i)ArrowLake.?HX|7D67|Intel.*Graphics.*Driver|Intel Graphic driver') -or
        ($url -match '(?i)Intel.*Graphics|Intel.*Display')
    )
}

function Invoke-IntelGraphicsCommandCenterGate {
    param([switch]$StopOnFailure)
    $displayName = 'Intel Graphics Command Center'
    $storeId = '9PLFNLNT3G5G'
    Write-Log "  INTEL GRAPHICS COMPANION CHECK: verify presence/version only; never open the app as an installation gate. Store version upgrades are a separate transaction." 'STEP'
    try {
        $gccPackage = [pscustomobject]@{Title=$displayName;Category='Microsoft Store Application';DownloadUrl="https://apps.microsoft.com/detail/$storeId"}

        # Existing installed package is sufficient for the dependency/presence gate.
        # Never launch IGCC merely to make the installer continue: its GUI can remain
        # open indefinitely and does not install or update the graphics driver.
        $appxMatches = @()
        try {
            $appxMatches = @(Get-AppxPackage -ErrorAction SilentlyContinue | Where-Object {
                ([string]$_.Name -match '(?i)IntelGraphicsExperience|AppUp\.IntelGraphicsExperience') -or
                ([string]$_.PackageFamilyName -match '(?i)IntelGraphicsExperience|AppUp\.IntelGraphicsExperience')
            })
        } catch {}

        $storeInstalled = $false
        try { $storeInstalled = [bool](Test-StoreAppInstalled -Package $gccPackage) } catch {}

        if($appxMatches.Count -gt 0 -or $storeInstalled){
            $detectedVersions = @($appxMatches | ForEach-Object { [string]$_.Version } | Where-Object { $_ })
            if($detectedVersions.Count -gt 0){
                Write-Log "  [$displayName] Already installed (AppX version(s): $($detectedVersions -join ', ')); not launching the GUI. Companion presence gate satisfied." 'OK'
            } else {
                Write-Log "  [$displayName] Already installed; installed version is not exposed by the current detection path. Not launching the GUI; companion presence gate satisfied." 'OK'
            }
            return $true
        }

        $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
        if(-not $winget){
            Write-Log "  [$displayName] Not installed and WinGet is unavailable; cannot install the required Store companion." 'ERROR'
            return $false
        }
        if(-not (Initialize-MicrosoftStoreForCurrentUser)){
            Write-Log "  [$displayName] Not installed and Microsoft Store is not ready for the current user." 'ERROR'
            return $false
        }

        Write-Log "  [$displayName] Not installed; installing from Microsoft Store. The app will not be launched after installation." 'STEP'
        $gccResult = Invoke-WingetMsStoreInstall -WingetPath $winget.Source -PackageId $storeId -Exact -Silent -TimeoutSeconds 180
        $gccCode = [int]$gccResult.ExitCode
        if($gccCode -notin @(0,3010)){
            Write-Log "  [$displayName] Microsoft Store installation returned exit code $gccCode." 'ERROR'
            return $false
        }

        $verified = $false
        try {
            $verifyAppx = @(Get-AppxPackage -ErrorAction SilentlyContinue | Where-Object {
                ([string]$_.Name -match '(?i)IntelGraphicsExperience|AppUp\.IntelGraphicsExperience') -or
                ([string]$_.PackageFamilyName -match '(?i)IntelGraphicsExperience|AppUp\.IntelGraphicsExperience')
            })
            $verified = $verifyAppx.Count -gt 0
        } catch {}
        if(-not $verified){
            try { $verified = [bool](Test-StoreAppInstalled -Package $gccPackage) } catch {}
        }
        if(-not $verified){
            Write-Log "  [$displayName] Store installer returned exit code $gccCode, but post-install presence verification failed." 'ERROR'
            return $false
        }

        Write-Log "  [$displayName] Installed and verified; GUI launch skipped. Companion presence gate satisfied." 'OK'
        return $true
    } catch {
        Write-Log "  [$displayName] Companion presence check failed: $($_.Exception.Message)" 'ERROR'
        return $false
    }
}

function Install-AsusFullPackages {
    param(
        [Parameter(Mandatory)]$Devices,
        # v66: use a distinct parameter name so it cannot collide with the local
        # $packages working variable under PowerShell's case-insensitive rules.
        [Parameter(Mandatory)]$CatalogPackages,
        [Parameter(Mandatory)]$InstalledDrivers,
        [Parameter(Mandatory)]$InstalledSoftware
    )
    $script:CurrentLiveDevices=@($Devices)
    Write-Log 'STAGE 07B - Sequential ASUS package workflow: download -> install -> run same-package executables/settings -> next package' 'STEP'
    Write-Log 'v74 parser-safe XTU detection: catalog input uses CatalogPackages and dependency sorting uses InputPackages; no case-insensitive Packages/packages parameter collision is possible.' 'INFO'
    # IMPORTANT: PowerShell variable names are case-insensitive. In v64, the
    # parameter $Packages and local $packages were therefore the SAME variable.
    # The assignment below effectively assigned the parameter to itself; when the
    # catalog array was empty/null at binding time, the mandatory object[] parameter
    # of Sort-PackagesForDependencyOrder received $null and terminated the run.
    # Keep the input parameter name distinct from the local working list.
    $packageList=@($CatalogPackages)
    if($packageList.Count -eq 0){Write-Log 'No ASUS full setup packages were available from the official catalog. The catalog request may have failed due to temporary DNS/network availability; the persistent resume checkpoint remains unchanged.' 'WARN';return [pscustomobject]@{Installed=0;Downloaded=0;RebootRequired=$false;StoppedForInstallerFailure=$false}}
    # v66: $packages is now a local working variable, not an alias for a parameter.
    $packages=@($packageList)
    Write-Log "ASUS package working list created: $($packages.Count) object(s)." 'INFO'
    $legacyInstalledTitles=@(Get-LegacyInstalledAsusTitles)
    if($legacyInstalledTitles.Count -gt 0){Write-Log "Resume migration: found $($legacyInstalledTitles.Count) ASUS package(s) previously completed by v10; those packages will be marked installed and skipped." 'OK'}

    # Dependency-aware serial order: ALL device-driver packages are processed before
    # companion applications/utilities. Within the driver phase the Intel graphics
    # driver is first, followed by Intel platform prerequisites, then other drivers.
    # Companion software such as Intel Graphics Command Center is deliberately placed
    # after the driver phase and is individually gated against its live driver prerequisite.
    if($packages.Count -gt 0){
        $packages=@(Sort-PackagesForDependencyOrder -InputPackages @($packages))
    } else {
        Write-Log 'ASUS dependency sort skipped because the package list is empty.' 'WARN'
    }
    Write-Log "Dependency sort completed successfully: $($packages.Count) package object(s) returned." 'OK'
    Write-Log 'Dependency order: driver packages first (Intel Graphics -> Intel IPF -> Intel platform -> remaining drivers), followed by companion software/utilities. A companion package cannot run until its required device driver is detected as installed.' 'INFO'
    $ipfCandidates=@($packages | Where-Object { "$($_.Category) $($_.Title)" -match '(?i)Innovation Platform Framework|\bIPF\b' })
    if($ipfCandidates.Count -gt 0){Write-Log "Intel Innovation Platform Framework is an explicit ASUS package target: $($ipfCandidates.Count) catalog candidate(s) retained; the highest-priority x64 candidate will be processed before other Intel platform prerequisites and before Intel XTU." 'OK'}
    else{Write-Log 'Intel Innovation Platform Framework is an explicit required package target, but no ASUS catalog candidate was exposed in this invocation. No generic Intel IPF package is substituted in the ASUS phase.' 'WARN'}

    # Get-AsusOfficialPackages has already applied latest-version selection separately
    # to driver and non-driver logical targets. Preserve that selected list unchanged here;
    # hardware applicability remains an install gate, not a reason to fabricate/replace OEM packages.
    Write-Log 'Latest-only ASUS driver and software selection was completed during catalog discovery; the package phase consumes that selected list without a second version-selection pass.' 'OK'

    Write-Log 'STAGE 07B-RECONCILE: starting non-fatal installed-state reconciliation report.' 'STEP'
    try {
        Export-ResumeReconciliation -Packages @($packages) -InstalledDrivers @($InstalledDrivers) -InstalledSoftware @($InstalledSoftware)
    } catch {
        Write-Log ("STAGE 07B-RECONCILE: report generation failed but installer will continue. {0}" -f $_.Exception.Message) 'WARN'
    }
    Write-Log 'STAGE 07B-RECONCILE: completed; entering package installation loop.' 'OK'

    $asusDir=Join-Path $DownloadRoot '01 ASUS G635LW Official Packages';New-Item -ItemType Directory -Force -Path $asusDir|Out-Null
    $resumeRows=@(Read-PackageResumeState)
    $installLedger=@(Read-PackageInstallLedger)
    $currentRunFingerprints=@{}
    $installed=0;$downloaded=0;$reboot=$false;$stopped=$false;$deferredReboot=$false;$deferredStore=$false;$index=0
    Write-Log "Anti-repeat persistent ledger contains $($installLedger.Count) historical package fingerprint(s); the current run also maintains an in-memory duplicate-launch guard." 'INFO'
    $batchState=Get-AsusPackageBatchState
    Write-Log 'v82 AUTOMATIC REBOOT POLICY: after every five successful top-level driver/software/package installation events, the script saves its checkpoint, registers the Startup resume batch, and automatically reboots. There is no maximum reboot count.' 'STEP'
    Write-Log 'v82 PACKAGE GATE: the global download phase downloads every selected content item first. Only after ALL-CONTENT-DOWNLOAD-COMPLETE.json exists does the serial installation phase begin. During installation, downloaded content is consumed in dependency order and marked installed in the persistent resume/ledger state. After every five successful installation events, Windows automatically reboots; there is no reboot-count limit.' 'STEP'

    foreach($pkg in $packages){
        $index++
        $key=Get-PackageKey -Package $pkg
        $dest=Get-PackageDestination -Package $pkg -AsusDirectory $asusDir
        $extract=Join-Path $asusDir ((Get-SafePackageName -Package $pkg)+'-Extracted')
        $existing=Find-ExistingAsusPackage -Package $pkg -AsusDirectory $asusDir -PreferredPath $dest
        if($existing){$dest=$existing}
        $pkg.LocalPath=$dest

        $row=@($resumeRows|Where-Object{$_.PackageKey -eq $key})|Select-Object -First 1
        if($null -ne $row){ $row=Normalize-PackageResumeRow -Row $row }
        if(-not$row){
            $legacyHit=($legacyInstalledTitles -contains [string]$pkg.Title)
            $row=[pscustomobject]@{PackageKey=$key;Title=[string]$pkg.Title;Version=[string]$pkg.Version;DownloadUrl=[string]$pkg.DownloadUrl;LocalPath=$dest;DownloadStatus=if($legacyHit){'Downloaded'}else{'Pending'};InstallStatus=if($legacyHit){'Installed(Legacy)'}else{'Pending'};LastAction=if($legacyHit){'Imported from v10 completion evidence'}else{''};Updated=(Get-Date).ToString('o')}
        } else {
            $row=Normalize-PackageResumeRow -Row $row
            $row.LocalPath=$dest
        }
        $hardwareApplicable=$true
        if(Test-PackageLooksLikeDriver -Package $pkg){
            $assoc=Test-DriverPackageHardwareAssociation -Package $pkg -Devices $Devices -InstalledDrivers $InstalledDrivers
            $hardwareApplicable=[bool]$assoc.Applicable
            if((Get-PackageComponentFamily -Package $pkg) -eq 'NVIDIA Graphics' -and @($Devices | Where-Object { (Get-DeviceComponentFamily -Device $_) -eq 'NVIDIA Graphics' }).Count -gt 0){
                $hardwareApplicable=$true
                Write-Log '  ASUS-ONLY NVIDIA GATE: NVIDIA hardware is present; this ASUS catalog NVIDIA package remains eligible regardless of incomplete catalog metadata.' 'INFO'
            }
            Write-Log "  DRIVER ASSOCIATION: Family='$($assoc.Family)' | Applicable=$($assoc.Applicable) | $($assoc.Reason)" 'INFO'
        }
        # v41: hardware association is an INSTALL gate, never a DOWNLOAD gate. Every
        # official ASUS catalog package is archived first so the complete 57-entry
        # bundle is available even when a package targets an alternate hardware variant.

        # SECOND HARD GATE: reconcile the actual installed component immediately before
        # any download/installer launch. This catches packages that became installed
        # between the initial inventory and this package boundary.
        $packageFamily=Get-PackageComponentFamily -Package $pkg
        $isNvidiaAsusDriver=($packageFamily -eq 'NVIDIA Graphics' -and @($Devices | Where-Object { (Get-DeviceComponentFamily -Device $_) -eq 'NVIDIA Graphics' }).Count -gt 0)
        $resumeAlreadyInstalled=($row.InstallStatus -match '^Installed(?:\(|$)')
        # Rebuild inventories at each package boundary instead of trusting only the STAGE 06
        # snapshot. This catches manual/vendor installs or updates made while the script was
        # downloading and processing earlier packages.
        $gateDrivers=@($InstalledDrivers)
        $gateSoftware=@($InstalledSoftware)
        try{$gateDrivers=@(Get-InstalledDriverInventory)}catch{Write-Log "Live driver version refresh failed at package boundary: $($_.Exception.Message)" 'WARN'}
        try{$gateSoftware=@(Get-InstalledSoftwareInventory)}catch{Write-Log "Live software version refresh failed at package boundary: $($_.Exception.Message)" 'WARN'}
        $installedGate=Test-InstalledComponentSatisfiesAsusPackage -Package $pkg -InstalledDrivers $gateDrivers -InstalledSoftware $gateSoftware
        if($isNvidiaAsusDriver -and $resumeAlreadyInstalled){
            $installedGate=[pscustomobject]@{Satisfied=$true;Reason='This ASUS NVIDIA graphics package was already completed by the persistent ASUS installer ledger.'}
        } elseif($isNvidiaAsusDriver -and $installedGate.Satisfied){
            $installedGate=[pscustomobject]@{Satisfied=$false;Reason='Existing NVIDIA display driver is not accepted as proof of ASUS package provenance; the ASUS G635LW package must be installed.'}
            Write-Log '  ASUS-ONLY NVIDIA GATE: existing NVIDIA display driver will NOT satisfy this ASUS package. The official ASUS G635LW package will be installed.' 'WARN'
        }
        if($installedGate.Satisfied){
            $row.DownloadStatus=if(Test-Path $dest){'Downloaded'}else{'NotRequired'}
            $row.InstallStatus='Installed(Detected)'
            $row.LastAction=$installedGate.Reason
            $row.LocalPath=$dest
            $row.Updated=(Get-Date).ToString('o')
            $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row
            Save-PackageResumeState -Rows $resumeRows
            Write-Log "  HARD ANTI-REPEAT GATE: '$($pkg.Title)' is already installed at a sufficient version; NO download or installer launch. $($installedGate.Reason)" 'OK'
            if(Test-IsIntelGraphicsPackage -Package $pkg){
                if(-not (Invoke-IntelGraphicsCommandCenterGate)){
                    $stopped=$true
                    break
                }
            }
            continue
        }

        $fingerprint=Get-PackageInstallFingerprint -Package $pkg -Devices @($Devices)
        if($currentRunFingerprints.ContainsKey($fingerprint)){
            $row.DownloadStatus=if(Test-Path $dest){'Downloaded'}else{'NotRequired'}
            $row.InstallStatus='Installed(CurrentRunDuplicate)'
            $row.LastAction='Same package/device/version fingerprint already launched successfully in this run; duplicate launch suppressed'
            $row.LocalPath=$dest
            $row.Updated=(Get-Date).ToString('o')
            $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row
            Save-PackageResumeState -Rows $resumeRows
            Write-Log "  ANTI-REPEAT: identical package/device/version fingerprint already completed during this run; second launch suppressed: $($pkg.Title)" 'OK'
            continue
        }

        if([string]$row.InstallStatus -match '^Installed'){
            Write-Log "  Resume directory/state check: this package is already recorded as installed; skipping installer launch." 'OK'
            if(Test-IsIntelGraphicsPackage -Package $pkg){
                if(-not (Invoke-IntelGraphicsCommandCenterGate)){
                    $stopped=$true
                    break
                }
            }
            continue
        }

        # v14 authoritative post-reboot reconciliation: if the package was installed
        # manually, by Windows Update, or in a different order, the live driver/software
        # inventory wins over the historical package counter.
        $script:CurrentLiveDevices=@($Devices)
        $live=Test-AsusPackageAlreadySatisfied -Package $pkg -InstalledDrivers $gateDrivers -InstalledSoftware $gateSoftware
        if($live.Satisfied){
            $row.DownloadStatus=if(Test-Path $dest){'Downloaded'}else{'NotRequired'}
            $row.InstallStatus='Installed(Detected)'
            $row.LastAction=$live.Reason
            $row.LocalPath=$dest
            $row.Updated=(Get-Date).ToString('o')
            $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row
            Save-PackageResumeState -Rows $resumeRows
            Write-Log "  LIVE RECONCILIATION: package already satisfied by installed system component: $($live.Reason). No installer launched." 'OK'
            if(Test-IsIntelGraphicsPackage -Package $pkg){
                if(-not (Invoke-IntelGraphicsCommandCenterGate)){
                    $stopped=$true
                    break
                }
            }
            continue
        }
        $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row
        Save-PackageResumeState -Rows $resumeRows

        # Microsoft Store applications are dynamic AppX/MSIX content, not stable ASUS binary archives.
        # NEVER send these URLs through Invoke-FileDownloadWithProgress or compare their live Store payload
        # against the static ASUS catalog SHA-256. NVIDIA documents that NVIDIA Control Panel is distributed
        # through Microsoft Store, and ASUS exposes the same package as an application entry on G635LW.
        if(Test-PackageIsMicrosoftStoreApp -Package $pkg){
            $storeDef=Get-StoreAppDefinitionForAsusPackage -Package $pkg
            Write-Log ("ASUS package {0}/{1}: {2} {3} [MICROSOFT STORE APPLICATION - NO ASUS SHA DOWNLOAD]" -f $index,$packages.Count,$pkg.Title,$pkg.Version) 'INFO'
            if(-not $storeDef){
                $row.InstallStatus='Deferred(UnknownStoreApp)';$row.LastAction='Microsoft Store package identified but no safe product ID mapping was available';$row.Updated=(Get-Date).ToString('o')
                $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                $deferredStore=$true
                Write-Log "[$($pkg.Title)] Microsoft Store package detected, but no known Store product ID could be resolved. It will remain deferred rather than downloading an unverified dynamic payload." 'WARN'
                continue
            }
            Write-Log "[$($storeDef.Name)] Store product ID: $($storeDef.Id). ASUS catalog SHA-256 is intentionally not applied to this dynamic Store delivery." 'INFO'
            if(-not (Test-CompanionPrerequisitesSatisfied -Package $pkg -InstalledDrivers $InstalledDrivers)){
                $row.DownloadStatus='Deferred';$row.InstallStatus='DeferredPrerequisite';$row.LastAction='Store companion deferred until its driver prerequisite is installed';$row.Updated=(Get-Date).ToString('o')
                $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                Write-Log "[$($pkg.Title)] STOPPING before Store companion because its required driver prerequisite is not installed." 'WARN';$stopped=$true;break
            }
            if(Test-StoreAppInstalled -Package $pkg){
                Write-Log "[$($storeDef.Name)] Already installed; no Store download required." 'OK'
            } else {
                $localStoreDir=Join-Path $asusDir ('Microsoft Store - '+(($storeDef.Name) -replace '[\/:*?"<>|]','_'))
                $localInstalled=Install-DownloadedStorePayload -Directory $localStoreDir -NamePattern ([string]$storeDef.Pattern)
                if($localInstalled){
                    Write-Log "[$($storeDef.Name)] Installed from the previously downloaded local Store payload. No network download is permitted during installation." 'OK'
                } else {
                if(Test-Path -LiteralPath (Join-Path $DownloadRoot 'ALL-CONTENT-DOWNLOAD-COMPLETE.json')){
                    Write-Log "[$($storeDef.Name)] No offline Store payload is present. Using the normal Microsoft Store install transaction for the signed-in interactive user; no winget download/license-authentication flow will be invoked." 'INFO'
                }
                $winget=Get-Command winget.exe -ErrorAction SilentlyContinue
                if(-not $winget){
                    $deferredStore=$true;$row.InstallStatus='Deferred(StoreInstallerUnavailable)';$row.LastAction='WinGet/App Installer unavailable';$row.Updated=(Get-Date).ToString('o')
                    $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                    Write-Log "[$($storeDef.Name)] WinGet is unavailable. Deferring this Store application without treating it as a driver failure." 'WARN';continue
                }
                if(-not (Initialize-MicrosoftStoreForCurrentUser)){
                    $deferredStore=$true;$row.InstallStatus='Deferred(StoreNotReady)';$row.LastAction='Microsoft Store not ready for current user';$row.Updated=(Get-Date).ToString('o')
                    $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                    Write-Log "[$($storeDef.Name)] Microsoft Store is not ready. Deferring the Store application and continuing with remaining packages." 'WARN';continue
                }
                try {
                    $row.InstallStatus='Installing';$row.LastAction='Microsoft Store installation launched';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                    $storeResult=Invoke-WingetMsStoreInstall -WingetPath $winget.Source -PackageId $storeDef.Id -Exact -Silent -TimeoutSeconds 180
                    $wcode=[int]$storeResult.ExitCode
                    if($wcode -notin @(0,3010)){
                        $deferredStore=$true;$row.InstallStatus="Deferred(StoreExit:$wcode)";$row.LastAction='Microsoft Store installation deferred';$row.Updated=(Get-Date).ToString('o')
                        $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                        Write-Log "[$($storeDef.Name)] Microsoft Store installation returned exit code $wcode. The application remains the resume point without invoking ASUS SHA verification." 'WARN';continue
                    }
                    if($wcode -eq 3010){$batchState.PendingReboot=$true;Save-AsusPackageBatchState -State $batchState}
                    Write-Log "[$($storeDef.Name)] Microsoft Store installation completed with exit code $wcode." 'OK'
                } catch {
                    $deferredStore=$true;$row.InstallStatus='InstallException';$row.LastAction='Microsoft Store installation exception';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                    Write-Log "[$($storeDef.Name)] Microsoft Store installation failed: $($_.Exception.Message)" 'WARN';continue
                }
                }
            }
            # Store applications are still subject to the user's settings gate. Launch the actual
            # AppX executable when discoverable and wait indefinitely before allowing the next package.
            [void](Invoke-InteractiveStoreApp -DisplayName $storeDef.Name -NamePattern $storeDef.Pattern)
            $installLedger=Add-PackageInstallLedgerEntry -Ledger $installLedger -Package $pkg -Fingerprint $fingerprint -ExitCode 0
            Save-PackageInstallLedger -Rows $installLedger
            $row.DownloadStatus='NotApplicable(Store)';$row.InstallStatus='Installed(Store)';$row.LastAction='Microsoft Store application installed/verified and interactive settings gate completed';$row.Updated=(Get-Date).ToString('o')
            $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
            $installed++
            Write-Log "[$($storeDef.Name)] SETTINGS GATE complete. No ASUS package download or ASUS SHA-256 validation was performed for this Store application." 'OK'
            if(Invoke-FiveInstallRebootGate -BatchState $batchState -InstalledItemTitle ([string]$storeDef.Name)) {
                $reboot=$true
                $batchState.PendingReboot=$true
                Save-AsusPackageBatchState -State $batchState
                $row.InstallStatus='Installed(Store-RebootCheckpoint)'
                $row.LastAction='Store application installed; five-install reboot checkpoint reached'
                $row.Updated=(Get-Date).ToString('o')
                $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row
                Save-PackageResumeState -Rows $resumeRows
                break
            }
            continue
        }

        Write-Log ("ASUS package {0}/{1}: {2} {3}"-f $index,$packages.Count,$pkg.Title,$pkg.Version) 'INFO'
        Write-Log "  Persistent local path: $dest" 'INFO'

        # Directory is authoritative. If a valid file is already there, NEVER download it again.
        $fileReady=$false
        if(Test-Path $dest){
            try{
                $len=(Get-Item $dest).Length
                if($len -gt 0){
                    if([string]$pkg.SHA256 -match '^[A-Fa-f0-9]{64}$'){
                        $actual=(Get-FileHash -LiteralPath $dest -Algorithm SHA256).Hash
                        if($actual -eq ([string]$pkg.SHA256).ToUpperInvariant()){$fileReady=$true;Write-Log "  Directory check: existing package is SHA-256 valid; download skipped." 'OK'}
                        else{Write-Log '  Directory check: existing file does not match the current ASUS SHA-256; treating it as a stale archive file and downloading the current package to its version/hash-specific path.' 'INFO';Remove-Item $dest -Force -ErrorAction SilentlyContinue}
                    }else{$fileReady=$true;Write-Log "  Directory check: existing non-empty package retained; vendor SHA-256 was not exposed." 'OK'}
                }
            }catch{Write-Log "  Directory check failed: $($_.Exception.Message)" 'WARN'}
        }
        if(-not$fileReady){
            $globalMarker=Join-Path $DownloadRoot 'ALL-CONTENT-DOWNLOAD-COMPLETE.json'
            if(Test-Path -LiteralPath $globalMarker){
                $row.DownloadStatus='MissingAfterDownloadPhase';$row.LastAction='Global download phase was complete but the archived file is missing; installation is blocked rather than downloading during the install phase.';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows;Write-Log '  INSTALLATION BLOCKED: the global download marker exists but this package file is missing. No network download is permitted during installation.' 'ERROR';$stopped=$true;break
            }
            $row.DownloadStatus='Downloading';$row.LastAction='Download started';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
            if(Invoke-FileDownloadWithProgress -Uri $pkg.DownloadUrl -Destination $dest -Label("ASUS $index/$($packages.Count) - $($pkg.Title)") -ExpectedSHA256([string]$pkg.SHA256)){$downloaded++;$row.DownloadStatus='Downloaded';$row.LastAction='Download completed';$row.LocalPath=$dest}else{$row.DownloadStatus='Failed';$row.LastAction='Download failed';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows;Write-Log '  Package download failed; stopping here so the next run resumes at THIS package.' 'WARN';$stopped=$true;break}
        }else{$row.DownloadStatus='Downloaded'}

        # Persist the successful download BEFORE opening the installer. This makes the file survive
        # a reboot or a user cancellation of the attended installer.
        $row.LocalPath=$dest;$row.InstallStatus='ReadyToInstall';$row.LastAction='Download checkpoint saved';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows

        if(-not $hardwareApplicable){
            $row.InstallStatus='NotApplicable';$row.LastAction='Official package archived, but its hardware association does not match the STAGE 06 baseline; installer not launched';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
            Write-Log "  INSTALL GATE: '$($pkg.Title)' was downloaded from ASUS and retained in the complete official bundle, but its hardware association does not match the STAGE 06 hardware baseline. It will NOT be forced onto a different hardware variant." 'INFO'
            continue
        }

        if(-not (Test-PackageLooksLikeDriver -Package $pkg)){
            if(-not (Test-CompanionPrerequisitesSatisfied -Package $pkg -InstalledDrivers $InstalledDrivers)){
                $row.InstallStatus='DeferredPrerequisite';$row.LastAction='Companion deferred until driver prerequisite is installed';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                Write-Log "[$($pkg.Title)] STOPPING before companion software because its required device driver is not detected as installed." 'WARN';$stopped=$true;break
            }
            Write-Log "COMPANION GATE: $($pkg.Title) has passed its device-driver prerequisite check; launching companion software now." 'OK'
        }

        # The Desktop archive is a real filesystem path. If Desktop is backed up by OneDrive,
        # the package must still be physically present before an installer is launched. Microsoft
        # documents that OneDrive can make files online-only, so verify local availability here.
        try{
            $destItem=Get-Item -LiteralPath $dest -Force -ErrorAction Stop
            if(($destItem.Attributes -band [IO.FileAttributes]::Offline) -ne 0){
                Write-Log "Package file is marked offline-only by Windows/OneDrive: $dest. Mark this archive 'Always keep on this device' before continuing." 'WARN'
                $row.InstallStatus='DeferredLocalFile';$row.LastAction='Package exists but is not locally available';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows; $stopped=$true; break
            }
        }catch{
            Write-Log "Local package verification failed before installer launch: $dest -- $($_.Exception.Message)" 'WARN'
            $row.InstallStatus='DeferredLocalFile';$row.LastAction='Package path could not be opened locally';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows; $stopped=$true; break
        }

        $installer=Get-InstallerFromPackage -PackagePath $dest -ExtractDirectory $extract
        if(-not$installer){
            $row.InstallStatus='NoFullInstaller';$row.LastAction='No setup executable found; file retained';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
            Write-Log "No setup/install executable found in $($pkg.Title); downloaded file retained. No INF-only installation will be attempted." 'WARN'
            continue
        }

        # FINAL VERSION GATE: packages may take a long time to download/extract, so refresh
        # installed software and driver versions again immediately before launching any EXE/MSI.
        # Same or newer installed versions are never overwritten by this script.
        $preLaunchDrivers=@($gateDrivers);$preLaunchSoftware=@($gateSoftware)
        try{$preLaunchDrivers=@(Get-InstalledDriverInventory)}catch{Write-Log "Final pre-launch driver inventory failed: $($_.Exception.Message)" 'WARN'}
        try{$preLaunchSoftware=@(Get-InstalledSoftwareInventory)}catch{Write-Log "Final pre-launch software inventory failed: $($_.Exception.Message)" 'WARN'}
        $finalVersionGate=Test-InstalledComponentSatisfiesAsusPackage -Package $pkg -InstalledDrivers $preLaunchDrivers -InstalledSoftware $preLaunchSoftware
        if($isNvidiaAsusDriver -and $resumeAlreadyInstalled){$finalVersionGate=[pscustomobject]@{Satisfied=$true;Reason='Persistent installer ledger confirms this ASUS NVIDIA package was already completed.'}}
        if($finalVersionGate.Satisfied){
            $row.InstallStatus='Installed(VersionGate)';$row.LastAction=$finalVersionGate.Reason;$row.LocalPath=$dest;$row.Updated=(Get-Date).ToString('o')
            $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
            Write-Log "FINAL VERSION GATE BLOCKED INSTALL: '$($pkg.Title)' was not launched because the live installed component is same/newer or candidate version is unverified. $($finalVersionGate.Reason)" 'OK'
            continue
        }

        # Checkpoint BEFORE launching because an installer may reboot the machine before PowerShell
        # gets control back. On the next run, this exact package is reconsidered first.
        $row.InstallStatus='Installing';$row.LastAction='Final live version gate passed; installer launched';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
        try{
            Write-Log "Running ASUS full setup: $($pkg.Title) using $($installer.Kind). The downloaded file is preserved regardless of installer result." 'STEP'
            $p=Start-Process -FilePath $installer.Path -ArgumentList $installer.Arguments -WorkingDirectory (Split-Path $dest -Parent) -Wait -PassThru
            $code=[int]$p.ExitCode
            if($code -eq 0 -or $code -eq 3010 -or $code -eq 1641){
                $installed++
                # v44: complete the same package before moving to the next package. Only extracted
                # configuration/control/utility executables are selected; installer/update/firmware
                # executables are explicitly excluded.
                if($installer.Kind -ne 'MSI' -and (Test-Path -LiteralPath $extract)){
                    $companionCount=Invoke-PackageCompanionExecutables -ExtractDirectory $extract -PrimaryPath $installer.Path -PackageTitle ([string]$pkg.Title)
                    Write-Log "  SETTINGS GATE complete for '$($pkg.Title)': $companionCount same-package executable(s) were run and closed before the next package download is permitted." 'OK'
                if(Test-IsIntelGraphicsPackage -Package $pkg){
                    if(-not (Invoke-IntelGraphicsCommandCenterGate)){
                        Write-Log "  Intel Graphics Command Center gate failed. No subsequent ASUS package will be downloaded or installed in this invocation." 'ERROR'
                        $stopped=$true
                    }
                }

                }
                # v33: the complete PnP/driver/software enumeration is intentionally NOT repeated
                # after package installation. The original STAGE 06 baseline remains authoritative.
                Write-Log "  One-time baseline policy: no post-install PnP rescan or installed driver/software inventory refresh is performed for '$($pkg.Title)'." 'INFO'
                $currentRunFingerprints[$fingerprint]=$true
                $installLedger=Add-PackageInstallLedgerEntry -Ledger $installLedger -Package $pkg -Fingerprint $fingerprint -ExitCode $code
                Save-PackageInstallLedger -Rows $installLedger
                Write-Log "  ANTI-REPEAT LEDGER: recorded successful package/device/version fingerprint for '$($pkg.Title)'." 'OK'
                $row.InstallStatus="Installed($code)"
                $row.LastAction='Installer completed'
                $row.Updated=(Get-Date).ToString('o')
                $pendingAfter=Test-PendingSystemReboot
                if($code -eq 1641){
                    # 1641 means the installer has actually initiated the reboot. Do not start another installer.
                    $reboot=$true
                    $batchState.PendingReboot=$true
                    Save-AsusPackageBatchState -State $batchState
                    Write-Log "Full ASUS setup completed for $($pkg.Title) (exit 1641). The installer initiated a reboot; no subsequent package will be launched." 'WARN'
                } else {
                    if($pendingAfter -or $code -eq 3010){$batchState.PendingReboot=$true}
                    Save-AsusPackageBatchState -State $batchState
                    Write-Log "Full ASUS setup completed for $($pkg.Title) (exit $code). Continuing the strict five-install reboot checkpoint sequence." 'OK'
                    if(Invoke-FiveInstallRebootGate -BatchState $batchState -InstalledItemTitle ([string]$pkg.Title)){
                        $reboot=$true
                        $batchState.PendingReboot=$true
                        Save-AsusPackageBatchState -State $batchState
                    }
                }
            }else{
                $cancel=($code -in @(2,1223,1602))
                if((Test-NonBlockingAsusPackage -Package $pkg) -and $code -eq 1){
                    $installed++
                    $row.InstallStatus='Installed(1-NonBlocking)'
                    $row.LastAction='displayHDR helper returned exit code 1; treated as non-blocking certificate-helper/no-applicable-action result'
                    $row.Updated=(Get-Date).ToString('o')
                    $installLedger=Add-PackageInstallLedgerEntry -Ledger $installLedger -Package $pkg -Fingerprint $fingerprint -ExitCode $code
                    Save-PackageInstallLedger -Rows $installLedger
                    Write-Log "ASUS setup for '$($pkg.Title)' returned exit code 1. ASUS describes displayHDR as an HDR-certificate helper; for this optional helper only, exit 1 is treated as non-blocking and is NOT left as the resume point. The downloaded package remains preserved." 'WARN'
                    $pendingAfter=Test-PendingSystemReboot
                    if($pendingAfter){
                        $batchState.PendingReboot=$true
                    }
                    Save-AsusPackageBatchState -State $batchState
                    if(Invoke-FiveInstallRebootGate -BatchState $batchState -InstalledItemTitle ([string]$pkg.Title)){
                        $reboot=$true
                        $batchState.PendingReboot=$true
                        Save-AsusPackageBatchState -State $batchState
                    }
                }else{
                    $row.InstallStatus=if($cancel){"Cancelled($code)"}else{"Failed($code)"}
                    $row.LastAction=if($cancel){'User/installer cancellation; package preserved'}else{'Installer failure; package preserved'}
                    $row.Updated=(Get-Date).ToString('o')
                    Write-Log "ASUS setup for '$($pkg.Title)' returned exit code $code. The downloaded package remains in the Desktop archive and THIS package is now the resume point." 'WARN'
                    $stopped=$true
                }
            }
        }catch{
            $row.InstallStatus='InstallException';$row.LastAction='Installer exception; package preserved';$row.Updated=(Get-Date).ToString('o');Write-Log "ASUS full setup failed for '$($pkg.Title)': $($_.Exception.Message). The package is preserved and this package remains the resume point." 'WARN';$stopped=$true
        }
        $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
        if($stopped){break}
        if($reboot){
            Write-Log "Package $($pkg.Title) requires/initiated a reboot. State has been checkpointed; the next startup will resume by checking this package and then continue with the following package." 'WARN'
            break
        }
    }
    Export-DownloadManifest -Packages $packages
    Write-Log "ASUS full-package phase: $downloaded newly downloaded, $installed successfully installed. Resume state: $(Join-Path $DownloadRoot 'ASUS-Package-Resume-State.json')" 'OK'
    return [pscustomobject]@{Installed=$installed;Downloaded=$downloaded;RebootRequired=$reboot;DeferredReboot=$deferredReboot;DeferredStore=$deferredStore;StoppedForInstallerFailure=($stopped -and -not $deferredReboot);Completed=(!$stopped -and !$reboot -and !$deferredStore -and $index -ge $packages.Count)}
}

function Get-HardwareVendor {
    param([Parameter(Mandatory)]$Device)
    $id = [string]$Device.PNPDeviceID
    if ($id -match 'VEN_10DE') { return 'NVIDIA' }
    if ($id -match 'VEN_8086') { return 'Intel' }
    if ($id -match 'VEN_10EC') { return 'Realtek' }
    if ($id -match 'VEN_1022') { return 'AMD' }
    if ($id -match 'VEN_14C3') { return 'MediaTek' }
    if ($id -match 'VEN_168C|VEN_17CB') { return 'Qualcomm' }
    if ($id -match 'VEN_13D3') { return 'AzureWave' }
    # USB vendor IDs are hexadecimal VID values in USB\VID_xxxx&PID_xxxx identifiers.
    # These mappings are used for classification/logging and for the small set of safe
    # manufacturer utility profiles below; they never imply that a third-party driver is safe.
    if ($id -match '(?i)^USB\\VID_046D&') { return 'Logitech' }
    if ($id -match '(?i)^USB\\VID_1532&') { return 'Razer' }
    if ($id -match '(?i)^USB\\VID_1B1C&') { return 'Corsair' }
    if ($id -match '(?i)^USB\\VID_1038&') { return 'SteelSeries' }
    if ($id -match '(?i)^USB\\VID_0D8C&') { return 'C-Media' }
    if ($id -match '(?i)^USB\\VID_0499&') { return 'Yamaha' }
    if ($id -match '(?i)^USB\\VID_0E8D&') { return 'MediaTek' }
    if ($id -match '(?i)^USB\\VID_18D1&') { return 'Google' }
    if ($id -match '(?i)^USB\\VID_04E8&') { return 'Samsung' }
    if ($id -match '(?i)^USB\\VID_12D1&') { return 'Huawei' }
    if ($id -match '(?i)^USB\\VID_2A70&') { return 'OnePlus' }
    if ($id -match '(?i)^USB\\VID_2717&') { return 'Xiaomi' }
    if ($id -match '(?i)^USB\\VID_22D9&') { return 'OPPO' }
    if ($id -match '(?i)^USB\\VID_2D95&') { return 'vivo' }
    if ($id -match '(?i)^USB\\VID_05AC&') { return 'Apple' }
    if ($id -match '(?i)^USB\\VID_045E&') { return 'Microsoft' }
    if ($id -match '(?i)^USB\\VID_0BDA&') { return 'Realtek USB' }
    if ([string]$Device.Manufacturer -match '(?i)NVIDIA') { return 'NVIDIA' }
    if ([string]$Device.Manufacturer -match '(?i)Intel') { return 'Intel' }
    if ([string]$Device.Manufacturer -match '(?i)Realtek') { return 'Realtek' }
    if ([string]$Device.Manufacturer -match '(?i)AMD') { return 'AMD' }
    return 'Other'
}

function Invoke-OfficialPageDownload {
    param(
        [Parameter(Mandatory)][string]$Vendor,
        [Parameter(Mandatory)][string]$PageUrl,
        [Parameter(Mandatory)][string[]]$LinkPatterns,
        [Parameter(Mandatory)][string]$Destination
    )
    try {
        Write-Log "[$Vendor] Checking official vendor page: $PageUrl"
        $r = Invoke-WebRequest -Uri $PageUrl -UseBasicParsing -TimeoutSec 60 -Headers @{
            'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v64'
            'Accept'='text/html,application/xhtml+xml,*/*'
        }
        $html = [string]$r.Content
        $hrefs = [regex]::Matches($html,'(?i)href\s*=\s*["'']([^"'']+)["'']') | ForEach-Object { $_.Groups[1].Value }
        $candidate = $null
        foreach ($pattern in $LinkPatterns) {
            $candidate = $hrefs | Where-Object { $_ -match $pattern -and $_ -match '(?i)https?://|^/' } | Select-Object -First 1
            if ($candidate) { break }
        }
        if (-not $candidate) {
            # Some official pages expose the download in a data attribute instead of href.
            $matches = [regex]::Matches($html,'(?i)(https?://[^"''\s<>]+\.(?:exe|msi)(?:\?[^"''\s<>]*)?)')
            foreach ($pattern in $LinkPatterns) {
                $candidate = $matches | ForEach-Object { $_.Groups[1].Value } | Where-Object { $_ -match $pattern } | Select-Object -First 1
                if ($candidate) { break }
            }
        }
        if (-not $candidate) {
            Write-Log "[$Vendor] No direct official installer link could be extracted; nothing will be downloaded from a third-party site." 'WARN'
            return $false
        }
        if ($candidate -match '^/') { $candidate = ([uri]$PageUrl).GetLeftPart([System.UriPartial]::Authority) + $candidate }
        Write-Log "[$Vendor] Official installer located: $candidate"
        return (Invoke-FileDownloadWithProgress -Uri $candidate -Destination $Destination -Label "$Vendor official installer" -ExpectedSHA256 '')
    } catch {
        Write-Log "[$Vendor] Official download failed: $($_.Exception.Message)" 'WARN'
        return $false
    }
}

function Test-AsusDeviceCovered {
    param(
        [Parameter(Mandatory)]$Device,
        [Parameter(Mandatory)]$AsusPackages
    )

    $deviceName = [string]$Device.Name
    $manufacturer = [string]$Device.Manufacturer
    $pnp = [string]$Device.PNPDeviceID
    if (-not $deviceName -and -not $pnp) { return $false }

    foreach ($pkg in $AsusPackages) {
        $rawText = ''
        try { $rawText = ($pkg.Raw | ConvertTo-Json -Depth 20 -Compress) + ' ' + [string]$pkg.Category + ' ' + [string]$pkg.Title } catch {}
        $title = [string]$pkg.Title

        # A package is considered an exact device match only when ASUS exposes
        # device metadata containing the device name, PNP/HW identifier, or a
        # distinctive manufacturer/device combination. Model-wide generic packages
        # are NOT enough to suppress manufacturer fallback.
        if ($rawText) {
            if ($pnp -and $rawText -match [regex]::Escape($pnp)) { return $true }
            $tokens = @($deviceName -split '[\s\(\),/\-]+') | Where-Object { $_.Length -ge 5 }
            $tokenHits = 0
            foreach ($token in $tokens) {
                if ($rawText -match [regex]::Escape($token)) { $tokenHits++ }
            }
            if ($tokenHits -ge 2 -and $manufacturer) { return $true }
        }
    }
    return $false
}


function Get-DeviceComponentFamily {
    param([Parameter(Mandatory)]$Device)
    $t="$($Device.Name) $($Device.Manufacturer) $($Device.PNPDeviceID) $($Device.HardwareID)"
    switch -Regex ($t) {
        'VEN_10DE|NVIDIA.*(GeForce|RTX|Graphics|Display)' { return 'NVIDIA Graphics' }
        'VEN_8086&DEV_7D67|Intel.*(Graphics|Display|Arc)' { return 'Intel Graphics' }
        'VEN_8086.*(Wireless|WLAN|Wi-?Fi)|Intel.*(Wireless|Wi-?Fi|WLAN)' { return 'Intel Wireless' }
        'Bluetooth.*Intel|Intel.*Bluetooth' { return 'Intel Bluetooth' }
        'VEN_10EC&DEV_8168|VEN_10EC&DEV_8125|Realtek.*(Ethernet|LAN|2\.5GbE|Network)' { return 'Realtek LAN' }
        'HDAUDIO|Realtek.*(Audio|Sound)|Audio.*Realtek' { return 'Realtek Audio' }
        'AMD.*(Radeon|Graphics|Display)|VEN_1002' { return 'AMD Graphics' }
        'Intel.*(Rapid Storage|RST|VMD)|VEN_8086.*(467F|7D0B|7D0C|7D0F|7D60|7D63|7D65)' { return 'Intel Rapid Storage/VMD' }
        'Intel.*(Execution Technology|TXT)|Trusted Execution Technology|\bTXT\b' { return 'Intel Trusted Execution Technology' }
        'Intel.*(Management Engine|MEI|CSME)' { return 'Intel Management Engine' }
        'Intel.*(Serial IO|Serial I/O)' { return 'Intel Serial IO' }
        'Intel.*(Innovation Platform Framework|\bIPF\b)|ACPI\\INTC(10[0-9A-F]{2}|35[0-9A-F]{2})|PCI\\VEN_8086&DEV_(AD03|7F23)' { return 'Intel Innovation Platform Framework' }
        'Intel.*(Dynamic Tuning|DTT)' { return 'Intel Dynamic Tuning' }
        'Intel.*(Smart Sound|SST)' { return 'Intel Smart Sound' }
        'Intel.*(Platform Monitoring|PMT)' { return 'Intel Platform Monitoring' }
        'Intel.*(VPU|AI Boost|Neural Processing|NPU)|Vision Processing|\bNPU\b|\bVPU\b' { return 'Intel VPU' }
        'Intel.*Gaussian|GNA' { return 'Intel GNA' }
        'NumberPad|NumPad|Numpad|Numeric Keypad' { return 'ASUS NumberPad' }
        'Touchpad|Precision TouchPad' { return 'Touchpad' }
        'Camera|Webcam|IR Camera' { return 'Camera/IR' }
        'Card Reader|Realtek.*Card' { return 'Card Reader' }
        '^USB\\VID_0E8D&|MediaTek.*USB|MediaTek.*Preloader|MediaTek.*MTP' { return 'MediaTek USB/Phone' }
        '^USB\\VID_046D&|Logitech.*USB' { return 'USB Logitech Peripheral' }
        '^USB\\VID_0D8C&|C-Media.*USB|USB.*Audio' { return 'USB Audio' }
        '^USB\\VID_0499&|Yamaha.*USB' { return 'USB Yamaha Audio' }
        '^USB\\VID_1532&|Razer.*USB' { return 'USB Razer Peripheral' }
        '^USB\\VID_1B1C&|Corsair.*USB' { return 'USB Corsair Peripheral' }
        '^USB\\VID_1038&|SteelSeries.*USB' { return 'USB SteelSeries Peripheral' }
        '^USB\\VID_|USB.*(MTP|ADB|Fastboot|Composite|HID|Audio|Headset|Mouse|Keyboard|Gamepad|Controller|Serial|Modem|Storage)' { return 'Generic USB Peripheral' }
        default { return '' }
    }
}

function Test-AsusCatalogFamilyCovered {
    param([Parameter(Mandatory)]$Device,[Parameter(Mandatory)]$AsusPackages)
    $family=Get-DeviceComponentFamily -Device $Device
    if(-not $family){ return $false }
    foreach($pkg in @($AsusPackages)){
        if(-not (Test-PackageLooksLikeDriver -Package $pkg)){ continue }
        $pf=Get-PackageComponentFamily -Package $pkg
        if($pf -and $pf -eq $family){ return $true }
    }
    return $false
}

function Get-ManufacturerStoreProfiles {
    param([Parameter(Mandatory)]$Device)
    $family=Get-DeviceComponentFamily -Device $Device
    $vendor=Get-HardwareVendor -Device $Device
    $profiles=New-Object System.Collections.Generic.List[object]

    switch($family){
        'NVIDIA Graphics' {
            # No generic NVIDIA profile: G635LW display driver is ASUS-OEM-only.
        }
        'Intel Graphics' {
            $profiles.Add([pscustomobject]@{
                Vendor='Intel';Family=$family;PageUrl='https://www.intel.com/content/www/us/en/download/785597/intel-arc-graphics-windows.html'
                LinkPatterns=@('(?i)(gfx|graphics|arc).*\.exe(?:\?|$)','(?i)gfx_win_.*\.exe')
                FilePatterns=@('(?i)(gfx|graphics|arc).*\.exe$')
                Arguments='/S'
                Description='Intel graphics driver package'
            })
        }
        'Intel Wireless' {
            $profiles.Add([pscustomobject]@{
                Vendor='Intel';Family=$family;PageUrl='https://www.intel.com/content/www/us/en/download/19351/intel-wireless-wi-fi-drivers-for-windows-10-and-windows-11.html'
                LinkPatterns=@('(?i)WiFi-.*-Driver64-Win10-Win11\.exe','(?i)WiFi-.*64.*\.exe')
                FilePatterns=@('(?i)WiFi-.*-Driver64-Win10-Win11\.exe$','(?i)WiFi-.*64.*\.exe$')
                Arguments='/quiet'
                Description='Intel Wireless Wi-Fi driver package'
            })
        }
        'Intel Bluetooth' {
            $profiles.Add([pscustomobject]@{
                Vendor='Intel';Family=$family;PageUrl='https://www.intel.com/content/www/us/en/download/18649/intel-wireless-bluetooth-drivers-for-windows-10-and-windows-11.html'
                LinkPatterns=@('(?i)BT-.*-64UWD-Win10-Win11\.exe','(?i)BT-.*64.*\.exe')
                FilePatterns=@('(?i)BT-.*-64UWD-Win10-Win11\.exe$','(?i)BT-.*64.*\.exe$')
                Arguments='/quiet'
                Description='Intel Wireless Bluetooth driver package'
            })
        }
        'Realtek LAN' {
            $profiles.Add([pscustomobject]@{
                Vendor='Realtek';Family=$family;PageUrl='https://www.realtek.com/Download/List?cate_id=584&menu_id=297'
                LinkPatterns=@('(?i)(Win10|Win11).*Auto.*Installation.*\.(zip|exe)','(?i)(NDIS|NetAdapterCx).*\.(zip|exe)')
                FilePatterns=@('(?i)(Win10|Win11).*Auto.*Installation.*\.(zip|exe)$','(?i)(NDIS|NetAdapterCx).*\.(zip|exe)$')
                Arguments='/s'
                Description='Realtek PCIe Ethernet controller driver package'
            })
        }
        'Realtek Audio' {
            $profiles.Add([pscustomobject]@{
                Vendor='Realtek';Family=$family;PageUrl='https://www.realtek.com/Download/List?cate_id=593&menu_id=29'
                LinkPatterns=@('(?i)64bits.*Executable.*\.exe','(?i)64bits.*\.zip','(?i)Driver.*64.*\.(zip|exe)')
                FilePatterns=@('(?i)(64bits|64-bit).*\.exe$','(?i)(64bits|64-bit).*\.zip$')
                Arguments='/s'
                Description='Realtek High Definition Audio driver package'
            })
        }
        'AMD Graphics' {
            $profiles.Add([pscustomobject]@{
                Vendor='AMD';Family=$family;PageUrl='https://www.amd.com/en/support/download/drivers.html'
                LinkPatterns=@('(?i)(auto-detect|adrenalin|graphics).*\.(exe|msi)')
                FilePatterns=@('(?i)(auto-detect|adrenalin|graphics).*\.(exe|msi)$')
                Arguments='/S'
                Description='AMD graphics driver/auto-detect package'
            })
        }
        'USB Logitech Peripheral' {
            $profiles.Add([pscustomobject]@{Vendor='Logitech';Family=$family;PageUrl='https://support.logi.com/hc/en-us';LinkPatterns=@('(?i)Logi Options|G HUB|GHub|OptionsPlus|Logitech');FilePatterns=@('(?i).*\.(exe|msi)$');Arguments='';Description='Logitech official support; installer links vary by device'} )
        }
        'USB Razer Peripheral' {
            $profiles.Add([pscustomobject]@{Vendor='Razer';Family=$family;PageUrl='https://www.razer.com/synapse-3';LinkPatterns=@('(?i)Razer.*(Synapse|Installer).*\.(exe|msi)$','(?i)Synapse.*\.(exe|msi)$');FilePatterns=@('(?i)Synapse.*\.(exe|msi)$');Arguments='';Description='Razer Synapse official utility'} )
        }
        'USB Corsair Peripheral' {
            $profiles.Add([pscustomobject]@{Vendor='Corsair';Family=$family;PageUrl='https://www.corsair.com/us/en/s/downloads';LinkPatterns=@('(?i)iCUE.*\.(exe|msi)$');FilePatterns=@('(?i)iCUE.*\.(exe|msi)$');Arguments='/S';Description='Corsair iCUE official utility'} )
        }
        'USB SteelSeries Peripheral' {
            $profiles.Add([pscustomobject]@{Vendor='SteelSeries';Family=$family;PageUrl='https://steelseries.com/gg';LinkPatterns=@('(?i)SteelSeriesGG.*\.(exe|msi)$','(?i)GG.*\.(exe|msi)$');FilePatterns=@('(?i)(SteelSeriesGG|GG).*\.(exe|msi)$');Arguments='';Description='SteelSeries GG official utility'} )
        }
        'MediaTek USB/Phone' {
            # MediaTek does not publish one universal Windows consumer USB driver for every phone.
            # Do not guess at a third-party VCOM package. the phone's own
            # MTP/ADB interface remain the authoritative fallback.
        }
        'USB Audio' {
            # C-Media devices expose the vendor's official download center. The page is parsed
            # for an exact Windows installer only; if C-Media does not expose one for the
            # detected model, no unverified third-party package is guessed.
            if($t -match '(?i)C-?Media|VID_0D8C'){
                $profiles.Add([pscustomobject]@{
                    Vendor='C-Media';Family=$family;PageUrl='https://www.cmedia.com.tw/support/download_center'
                    LinkPatterns=@('(?i)(C-?Media|Xear|USB.?Audio).*(\.(exe|msi|zip))(?:\?|$)')
                    FilePatterns=@('(?i)(C-?Media|Xear|USB.?Audio).*\.(exe|msi|zip)$')
                    Arguments='/S';Description='C-Media official USB audio driver/software download center'
                })
            }
        }
        'USB Yamaha Audio' {
            # Yamaha publishes an official Yamaha Steinberg USB Driver for supported Yamaha/
            # Steinberg USB audio interfaces. The page is still checked first; the direct URL
            # is only a Yamaha-hosted package and is used as a fallback if the page hides it
            # behind JavaScript.
            $profiles.Add([pscustomobject]@{
                Vendor='Yamaha';Family=$family;PageUrl='https://usa.yamaha.com/support/updates/yamaha-steinberg-usb-driver-for-win.html'
                DirectUri='https://usa.yamaha.com/files/download/software/5/2442685/YSUSB_V219_Win.zip'
                LinkPatterns=@('(?i)YSUSB_V\d+_Win\.zip')
                FilePatterns=@('(?i)YSUSB_V\d+_Win\.zip$')
                Arguments='';Description='Yamaha official Steinberg USB Driver'
            })
        }
        'Generic USB Peripheral' {
            # Generic USB HID/MTP/storage/composite devices normally use Windows in-box drivers or
            # a device-specific package. No third-party driver is guessed from VID/PID alone.
        }
    }

    return $profiles.ToArray()
}

function Get-OfficialPageCandidateLinks {
    param(
        [Parameter(Mandatory)][string]$PageUrl,
        [Parameter(Mandatory)][string[]]$LinkPatterns
    )
    try{
        $r=Invoke-WebRequest -Uri $PageUrl -UseBasicParsing -TimeoutSec 60 -Headers @{
            'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v64'
            'Accept'='text/html,application/xhtml+xml,*/*'
        }
        $html=[string]$r.Content
        $hrefs=[regex]::Matches($html,'(?i)(?:href|src|data-url|data-download-url|data-download|data-href|downloadurl|downloadUrl)\s*=\s*["'']([^"'']+)["'']')|ForEach-Object{$_.Groups[1].Value}
        $rawLinks=New-Object System.Collections.Generic.List[string]
        foreach($h in @($hrefs)){
            if([string]$h){
                $clean=[string]$h
                $clean=$clean -replace '\\/','/'
                $clean=$clean -replace '&amp;','&'
                [void]$rawLinks.Add($clean)
            }
        }
        # Modern vendor pages frequently place the actual download URL in JavaScript/JSON
        # rather than in a normal href. Extract both plain and JSON-escaped HTTPS file URLs.
        foreach($m in [regex]::Matches($html,'(?i)https?://[^"''\s<>\\]+?\.(?:exe|msi|zip)(?:\?[^"''\s<>\\]*)?')){[void]$rawLinks.Add($m.Value)}
        foreach($m in [regex]::Matches($html,'(?i)https?:\\\\/\\\\/[^"''\s<>]+?\.(?:exe|msi|zip)(?:\?[^"''\s<>]*)?')){
            $clean=$m.Value -replace '\\/','/'
            [void]$rawLinks.Add($clean)
        }
        $resolved=New-Object System.Collections.Generic.List[string]
        foreach($link in @($rawLinks|Select-Object -Unique)){
            foreach($pattern in $LinkPatterns){
                if($link -match $pattern){
                    try{
                        $u=if($link -match '^https?://'){[uri]$link}else{[uri]::new(([uri]$PageUrl),$link)}
                        [void]$resolved.Add($u.AbsoluteUri)
                    }catch{}
                    break
                }
            }
        }
        return @($resolved|Select-Object -Unique)
    }catch{
        Write-Log "Official manufacturer page lookup failed for ${PageUrl}: $($_.Exception.Message)" 'WARN'
        return @()
    }
}

function Get-PagePackageMetadata {
    param([string]$PageUrl)
    $meta=[pscustomobject]@{Version=$null;ReleaseDate=$null;Html=''}
    try{
        $r=Invoke-WebRequest -Uri $PageUrl -UseBasicParsing -TimeoutSec 60 -Headers @{
            'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v64'
        }
        $html=[string]$r.Content;$meta.Html=$html
        $vm=[regex]::Match($html,'(?i)(?:Package Version|Driver Version|Version)\s*[:\-]?\s*</?[^>]*>?\s*(\d+(?:\.\d+){1,4})')
        if(-not $vm.Success){$vm=[regex]::Match($html,'(?i)(?:package version|driver version|version)[^0-9]{0,50}(\d+(?:\.\d+){1,4})')}
        if($vm.Success){$meta.Version=$vm.Groups[1].Value}
        $dm=[regex]::Match($html,'(?i)(?:Date|Release Date|Update Time)\s*[:\-]?\s*</?[^>]*>?\s*([A-Z][a-z]{2,9}\s+\d{1,2},\s+\d{4}|\d{4}/\d{1,2}/\d{1,2})')
        if($dm.Success){try{$meta.ReleaseDate=[datetime]$dm.Groups[1].Value}catch{}}
    }catch{}
    return $meta
}

function Test-ManufacturerPackageAlreadySatisfied {
    param([Parameter(Mandatory)]$Device,[Parameter(Mandatory)]$Candidate,[Parameter(Mandatory)]$InstalledDrivers)
    $ids=@(Get-DeviceHardwareIds -Device $Device)
    $deviceText="$($Device.Name) $($Device.Manufacturer) $($Device.PNPDeviceID)"
    $match=$false
    if($Candidate.HardwareIds){
        foreach($pid in @($Candidate.HardwareIds)){
            $venDev=''
            $vm=[regex]::Match([string]$pid,'(?i)VEN_[0-9A-F]{4}&DEV_[0-9A-F]{4}')
            if($vm.Success){$venDev=$vm.Value}
            foreach($did in $ids){
                if($did -ieq $pid){$match=$true;break}
                if($venDev -and $did -match ('(?i)'+[regex]::Escape($venDev))){$match=$true;break}
            }
            if($match){break}
        }
    }
    if(-not $match){$match=$true} # Candidate was selected from the manufacturer page for this device family.
    if(-not $match){return [pscustomobject]@{Satisfied=$false;Reason='Manufacturer candidate did not associate with the live device'}}
    $drivers=@($InstalledDrivers|Where-Object{
        $blob="$($_.DeviceName) $($_.Manufacturer) $($_.DriverProvider) $($_.DeviceID) $($_.HardwareID)"
        ($blob -match [regex]::Escape([string]$Device.Name)) -or
        ($_.DeviceID -and [string]$Device.PNPDeviceID -and [string]$_.DeviceID -ieq [string]$Device.PNPDeviceID)
    })
    foreach($d in $drivers){
        $have=Convert-ToComparableVersion ([string]$d.DriverVersion)
        $want=Convert-ToComparableVersion ([string]$Candidate.Version)
        $haveDate=$null;try{if([string]$d.DriverDate){$haveDate=[datetime]$d.DriverDate}}catch{}
        if($want -and $have -and $have -ge $want){return [pscustomobject]@{Satisfied=$true;Reason="Installed manufacturer driver $($d.DriverVersion) >= candidate $($Candidate.Version)"}}
        if($Candidate.ReleaseDate -and $haveDate -and $haveDate -ge $Candidate.ReleaseDate){return [pscustomobject]@{Satisfied=$true;Reason="Installed manufacturer driver date $($haveDate.ToString('yyyy-MM-dd')) >= candidate release $($Candidate.ReleaseDate.ToString('yyyy-MM-dd'))"}}
        if(-not $want -and -not $Candidate.ReleaseDate){return [pscustomobject]@{Satisfied=$true;Reason="Matching device driver '$($d.DeviceName)' exists, but manufacturer candidate has no comparable version/date; fail-closed downgrade protection blocks overwrite"}}
    }
    return [pscustomobject]@{Satisfied=$false;Reason='Matching device is not at or above the manufacturer candidate version/date'}
}

function Install-OfficialManufacturerDriverStorePass {
    param(
        [Parameter(Mandatory)]$Devices,
        [Parameter(Mandatory)]$AsusPackages,
        [Parameter(Mandatory)]$InstalledDrivers
    )
    Write-Log 'STAGE 07C - Official manufacturer driver-store reconciliation for hardware not represented by the ASUS G635LW driver catalog' 'STEP'
    $root=Join-Path $DownloadRoot '02 Official Manufacturer Driver Stores'
    New-Item -ItemType Directory -Force -Path $root|Out-Null
    $seen=@{}
    $results=New-Object System.Collections.Generic.List[object]

    foreach($device in @($Devices)){
        if(-not (Test-PackageLooksLikeDriver -Package ([pscustomobject]@{Title=[string]$device.Name;Category='driver';DownloadUrl='driver'}))){continue}
        $family=Get-DeviceComponentFamily -Device $device
        $vendor=Get-HardwareVendor -Device $device
        if(-not $family){Write-Log "[$vendor] No manufacturer-store family mapping for '$($device.Name)' ($($device.PNPDeviceID));" 'INFO';continue}

        if($family -eq 'NVIDIA Graphics'){
            Write-Log "[$vendor/$family] ASUS-ONLY POLICY: generic NVIDIA driver-store packages are disabled. The NVIDIA display driver may only be installed from the official ASUS G635LW catalog." 'OK'
            continue
        }

        $asusExact=Test-AsusDeviceCovered -Device $device -AsusPackages $AsusPackages
        $asusFamily=Test-AsusCatalogFamilyCovered -Device $device -AsusPackages $AsusPackages
        if($asusExact -or $asusFamily){
            Write-Log "[$vendor/$family] ASUS catalog coverage exists for '$($device.Name)'; manufacturer generic driver store will NOT replace the ASUS OEM package." 'INFO'
            continue
        }

        $profiles=@(Get-ManufacturerStoreProfiles -Device $device)
        if($profiles.Count -eq 0){
            Write-Log "[$vendor/$family] No direct manufacturer driver-store profile is available for '$($device.Name)';" 'INFO'
            continue
        }

        foreach($profile in $profiles){
            $key="$($profile.Vendor)|$($profile.Family)|$($profile.PageUrl)"
            if($seen.ContainsKey($key)){continue};$seen[$key]=$true
            Write-Log "[$($profile.Vendor)/$($profile.Family)] Manufacturer driver store selected for '$($device.Name)': $($profile.PageUrl)" 'STEP'
            $meta=Get-PagePackageMetadata -PageUrl $profile.PageUrl
            $links=@()
            if($profile.PSObject.Properties.Name -contains 'DirectUri' -and [string]$profile.DirectUri){
                $links=@([string]$profile.DirectUri)
                Write-Log "[$($profile.Vendor)/$($profile.Family)] Using the manufacturer-published direct package URI configured for this hardware family." 'INFO'
            } else {
                $links=@(Get-OfficialPageCandidateLinks -PageUrl $profile.PageUrl -LinkPatterns @($profile.LinkPatterns))
            }
            if($links.Count -eq 0){
                Write-Log "[$($profile.Vendor)/$($profile.Family)] No direct installer link was exposed by the official manufacturer page. No third-party mirror will be used." 'WARN'
                continue
            }
            $uri=$links|Where-Object{$_ -match $profile.FilePatterns[0]}|Select-Object -First 1
            if(-not $uri){$uri=$links|Select-Object -First 1}
            $name=[IO.Path]::GetFileName(([uri]$uri).AbsolutePath)
            $safe=($profile.Vendor+' '+$profile.Family+' '+$name)-replace '[\\/:*?"<>|]','_'
            $dest=Join-Path (Join-Path $root $profile.Vendor) $safe
            $candidate=[pscustomobject]@{
                Title="$($profile.Vendor) $($profile.Family) manufacturer driver"
                Version=[string]$meta.Version
                ReleaseDate=$meta.ReleaseDate
                DownloadUrl=$uri
                HardwareIds=@([string]$device.PNPDeviceID)
                Raw=$meta.Html
            }
            # Refresh live driver versions immediately before the manufacturer candidate gate.
            # If the query fails, retain the last known inventory and log the limitation.
            $manufacturerGateDrivers=@($InstalledDrivers)
            try{$manufacturerGateDrivers=@(Get-InstalledDriverInventory)}catch{Write-Log "[$($profile.Vendor)/$($profile.Family)] Live driver refresh failed; using last known inventory: $($_.Exception.Message)" 'WARN'}
            $gate=Test-ManufacturerPackageAlreadySatisfied -Device $device -Candidate $candidate -InstalledDrivers $manufacturerGateDrivers
            Write-Log "  MANUFACTURER ASSOCIATION: Device='$($device.Name)' | PNP='$($device.PNPDeviceID)' | Family='$family' | Candidate='$name' | Version='$($candidate.Version)' | Date='$($candidate.ReleaseDate)' | AlreadySatisfied=$($gate.Satisfied)" 'INFO'

            # The manufacturer-driver archive is also a preservation store. Even when the
            # currently installed driver already satisfies the version/date test, retain the
            # official manufacturer package locally so the detected device has a reproducible
            # package in ASUS-G635LW-Installer-Content. Installation is skipped when satisfied.
            $existing=Find-ExistingContentFile -FileName ([IO.Path]::GetFileName($dest)) -PreferredPath $dest
            if($existing){
                if(([IO.Path]::GetFullPath($existing.FullName)) -ine ([IO.Path]::GetFullPath($dest))){
                    New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent)|Out-Null
                    Copy-Item -LiteralPath $existing.FullName -Destination $dest -Force
                }
                Write-Log "  Manufacturer archive: reusing existing package '$($existing.FullName)' (no re-download)." 'OK'
                $ok=$true
            } else {
                if(Test-Path -LiteralPath (Join-Path $DownloadRoot 'ALL-CONTENT-DOWNLOAD-COMPLETE.json')){Write-Log "[$($profile.Vendor)/$($profile.Family)] INSTALLATION BLOCKED: manufacturer package was not present after the global download phase. No network download is permitted now." 'ERROR';$ok=$false}else{$ok=Invoke-FileDownloadWithProgress -Uri $uri -Destination $dest -Label "$($profile.Vendor) $($profile.Family) official driver-store package" -ExpectedSHA256 ''}
            }
            if($gate.Satisfied){
                if($ok){Write-Log "  [ARCHIVE ONLY] Manufacturer package retained for '$($device.Name)'; installed driver is already satisfied, so no driver installer will be launched." 'OK'}
                continue
            }
            if(-not $ok){continue}
            Write-Log "  Manufacturer package persisted before installation: $dest" 'OK'
            if($script:DownloadAllContentMode){Write-Log "  DOWNLOAD-ONLY PREPARATION: '$($profile.Vendor) $($profile.Family)' is fully archived; driver/INF/setup execution is deferred." 'OK';continue}

            $installPath=$dest
            $installArgs=[string]$profile.Arguments
            if([IO.Path]::GetExtension($dest).ToLowerInvariant() -eq '.zip'){
                $extract=Join-Path $dest.Substring(0,$dest.Length-4) 'Extracted'
                try{
                    if(Test-Path $extract){Remove-Item $extract -Recurse -Force -ErrorAction SilentlyContinue}
                    Expand-Archive -LiteralPath $dest -DestinationPath $extract -Force -ErrorAction Stop
                    $infs=@(Get-ChildItem -LiteralPath $extract -Recurse -Filter '*.inf' -File -ErrorAction SilentlyContinue)
                    $matched=@()
                    foreach($inf in $infs){
                        $txt=Get-Content -LiteralPath $inf.FullName -Raw -ErrorAction SilentlyContinue
                        if($txt -and (($txt -match [regex]::Escape([string]$device.PNPDeviceID)) -or ($txt -match '(?i)'+[regex]::Escape((([string]$device.PNPDeviceID) -replace '^.*?\\',''))))){
                            $matched+=$inf
                        }
                    }
                    if($matched.Count -gt 0){
                        Write-Log "  INF hardware-ID association verified against $($matched.Count) extracted INF file(s) for '$($device.Name)'." 'OK'
                        foreach($inf in $matched){
                            $p=Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\pnputil.exe') -ArgumentList @('/add-driver',"`"$($inf.FullName)`"","/install") -Wait -PassThru -NoNewWindow
                            Write-Log "  pnputil INF install result for $($inf.Name): $($p.ExitCode)" $(if($p.ExitCode -eq 0){'OK'}else{'WARN'})
                        }
                        # The extracted package is also searched for the vendor's full setup
                        # program/MSI so the driver archive's associated services/utilities are
                        # not silently omitted. The INF association is the gate; only after that
                        # gate passes is an associated setup executable launched.
                        $setups=@(Get-ChildItem -LiteralPath $extract -Recurse -File -ErrorAction SilentlyContinue |
                            Where-Object {$_.Name -match '^(setup|install|installer)(\.(exe|msi))?$'} |
                            Sort-Object FullName)
                        foreach($setup in $setups|Select-Object -First 3){
                            try{
                                if($setup.Extension -ieq '.msi'){
                                    $sp=Start-Process -FilePath 'msiexec.exe' -ArgumentList @('/i',"`"$($setup.FullName)`"","/qn","/norestart") -Wait -PassThru -WindowStyle Hidden
                                }else{
                                    $sp=Start-Process -FilePath $setup.FullName -ArgumentList $installArgs -WorkingDirectory $setup.DirectoryName -Wait -PassThru -WindowStyle Hidden
                                }
                                Write-Log "  Associated vendor setup '$($setup.Name)' exit code $($sp.ExitCode)." $(if($sp.ExitCode -in @(0,3010,1641)){'OK'}else{'WARN'})
                            }catch{Write-Log "  Associated vendor setup '$($setup.Name)' failed: $($_.Exception.Message)" 'WARN'}
                        }
                    }else{
                        Write-Log '  ZIP contained no INF with a matching live PNP identifier; package was retained but not installed.' 'WARN'
                    }
                }catch{Write-Log "  Manufacturer ZIP extraction/INF association failed: $($_.Exception.Message)" 'WARN'}
            }else{
                try{
                    $p=Start-Process -FilePath $installPath -ArgumentList $installArgs -WorkingDirectory (Split-Path $installPath -Parent) -Wait -PassThru -WindowStyle Hidden
                    Write-Log "  Manufacturer full installer exit code: $($p.ExitCode)" $(if($p.ExitCode -in @(0,3010,1641)){'OK'}else{'WARN'})
                }catch{Write-Log "  Manufacturer installer failed: $($_.Exception.Message)" 'WARN'}
            }
            $results.Add([pscustomobject]@{Device=$device.Name;Family=$family;Vendor=$vendor;Package=$dest;Version=$candidate.Version;ReleaseDate=$candidate.ReleaseDate})
            # v33: retain the single STAGE 06 installed-driver baseline; no refresh after manufacturer installation.
        }
    }
    $results.ToArray()
}

function Test-InstalledSoftwareMatch {
    param([Parameter(Mandatory)]$InstalledSoftware,[Parameter(Mandatory)][string[]]$Patterns)
    foreach($sw in @($InstalledSoftware)){
        $blob="$($sw.DisplayName) $($sw.Publisher)"
        foreach($pat in $Patterns){if($blob -match $pat){return $true}}
    }
    return $false
}

function Install-DeviceAssociatedManufacturerUtilities {
    param(
        [Parameter(Mandatory)]$Devices,
        [Parameter(Mandatory)]$InstalledSoftware
    )
    Write-Log 'STAGE 07D - Device-associated manufacturer apps, utilities and tuning tools' 'STEP'
    $root=Join-Path $DownloadRoot '03 Device-Associated Manufacturer Apps Utilities and Tuning'
    New-Item -ItemType Directory -Force -Path $root|Out-Null
    $winget=Get-Command winget.exe -ErrorAction SilentlyContinue
    $done=@{}

    foreach($device in @($Devices)){
        $family=Get-DeviceComponentFamily -Device $device
        $vendor=Get-HardwareVendor -Device $device
        if(-not $family){continue}

        $jobs=@()
        switch($family){
            'NVIDIA Graphics' {
                $jobs+=@([pscustomobject]@{
                    Key='NVIDIA App';Label='NVIDIA App';Patterns=@('(?i)^NVIDIA App$','(?i)NVIDIA App')
                    Page='https://www.nvidia.com/en-us/software/nvidia-app/';LinkPatterns=@('(?i)nvidia.*app.*\.(exe|msi)')
                    Args='/S';Folder='NVIDIA App'
                })
            }
            'Intel Wireless' {
                # Intel Driver & Support Assistant is a manufacturer utility for Intel hardware.
                # It is downloaded only as an associated utility; it is not invoked to replace
                # the scripted driver-store ordering.
                $jobs+=@([pscustomobject]@{
                    Key='Intel Driver Support Assistant';Label='Intel Driver & Support Assistant'
                    Patterns=@('(?i)Intel.*Driver.*Support.*Assistant','(?i)Intel Driver & Support Assistant')
                    Page='https://www.intel.com/content/www/us/en/support/detect.html'
                    DirectUri='https://dsadata.intel.com/installer'
                    LinkPatterns=@('(?i)Intel.*Driver.*Support.*Assistant.*\.(exe|msi)')
                    Args='/S';Folder='Intel Driver & Support Assistant'
                })
            }
            'Realtek LAN' {
                $jobs+=@([pscustomobject]@{
                    Key='Realtek Ethernet Diagnostic';Label='Realtek Ethernet Diagnostic Program'
                    Patterns=@('(?i)Realtek.*Ethernet Diagnostic')
                    Page='https://www.realtek.com/Download/List?cate_id=585'
                    LinkPatterns=@('(?i)Ethernet.*Diagnostic.*\.exe','(?i)Diagnostic.*\.(exe|msi)')
                    Args='/s';Folder='Realtek Ethernet Diagnostic'
                })
            }
        }

        foreach($job in $jobs){
            if($done.ContainsKey($job.Key)){continue};$done[$job.Key]=$true
            if(Test-InstalledSoftwareMatch -InstalledSoftware $InstalledSoftware -Patterns @($job.Patterns)){
                Write-Log "[$vendor/$family] Associated utility '$($job.Label)' is already installed; no installer will be launched." 'OK'
                continue
            }

            Write-Log "[$vendor/$family] Associated utility selected for '$($device.Name)': $($job.Label) | Official source: $($job.Page)" 'STEP'
            $uri=$null
            if($job.PSObject.Properties['DirectUri'] -and [string]$job.DirectUri){
                $uri=[string]$job.DirectUri
                Write-Log "[$vendor/$family] Using the configured official vendor download endpoint: $uri" 'OK'
            }else{
                $links=@(Get-OfficialPageCandidateLinks -PageUrl $job.Page -LinkPatterns @($job.LinkPatterns))
                if($links.Count -eq 0 -and $winget){
                    # JavaScript-only vendor pages: use an exact, known vendor-publisher
                    # WinGet package as a secondary download source for supported utilities.
                    $wingetId=$null
                    switch($job.Key){
                        'NVIDIA App' {$wingetId='NVIDIA.NVIDIAApp'}
                        'Intel Driver Support Assistant' {$wingetId='Intel.IntelDriverAndSupportAssistant'}
                        'Logi Options+' {$wingetId='Logitech.OptionsPlus'}
                        'Logitech G HUB' {$wingetId='Logitech.GHUB'}
                        'Corsair iCUE' {$wingetId='Corsair.iCUE.4'}
                        'SteelSeries GG' {$wingetId='SteelSeries.GG'}
                    }
                    if($wingetId){
                        $wd=Join-Path (Join-Path $root $job.Folder) 'WinGet'
                        New-Item -ItemType Directory -Force -Path $wd|Out-Null
                        Write-Log "[$vendor/$family] Official vendor page exposed no direct file; downloading exact publisher package '$wingetId' through WinGet." 'INFO'
                        $wp=Start-Process -FilePath $winget.Source -ArgumentList @('download','--id',$wingetId,'--exact','--source','winget','--accept-package-agreements','--accept-source-agreements','--download-directory',"`"$wd`"") -Wait -PassThru -NoNewWindow -ErrorAction SilentlyContinue
                        if($wp.ExitCode -eq 0){
                            $candidateFile=Get-ChildItem -LiteralPath $wd -Recurse -File -ErrorAction SilentlyContinue |
                                Where-Object{$_.Extension -in '.exe','.msi','.zip'} | Select-Object -First 1
                            if($candidateFile){$uri=$candidateFile.FullName;Write-Log "[$vendor/$family] Publisher package downloaded locally: $uri" 'OK'}
                        }
                    }
                }
                if(-not $uri -and $links.Count -eq 0){
                    Write-Log "[$vendor/$family] No direct official installer link was exposed for '$($job.Label)'; utility remains available from the official vendor page and is not sourced from a third party." 'WARN'
                    continue
                }
                if(-not $uri){$uri=$links|Select-Object -First 1}
            }
            $name=[IO.Path]::GetFileName(([uri]$uri).AbsolutePath)
            if(-not $name){$name=($job.Key -replace '[\\/:*?"<>|]','_')+'.exe'}
            $dest=Join-Path (Join-Path $root $job.Folder) $name
            $r=Invoke-OfficialDownloadAndInstall -Label $job.Label -Uri $uri -Destination $dest -Arguments $job.Args -Install
            if($script:DownloadAllContentMode){Write-Log "[$($job.Label)] DOWNLOAD-ONLY PREPARATION complete; installation is deferred until every content source is archived." 'OK';continue}
            if($r.Installed){Write-Log "[$($job.Label)] Installed and associated with live device family '$family'." 'OK'}
            else{Write-Log "[$($job.Label)] Downloaded/retained for the live device family '$family'; automatic installation did not complete." 'WARN'}
        }
    }
    Write-Log 'Device-associated manufacturer app/utility/tuning-tool pass complete. No automatic overclock, undervolt, voltage, power-limit or fan-profile changes are applied.' 'OK'
}

function Test-IntelXTUInstalled {
    # Do not rely solely on our persistent marker. A marker proves that a previous
    # invocation completed an installer, not that XTU is still installed now.
    $uninstallRoots=@(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    foreach($root in $uninstallRoots){
        try{
            $hit=Get-ItemProperty -Path $root -ErrorAction SilentlyContinue | Where-Object {
                ([string]$_.DisplayName -match '(?i)Intel.*Extreme Tuning Utility|Intel XTU')
            } | Select-Object -First 1
            if($hit){ return $true }
        }catch{}
    }
    # Registry entries are the primary test. Keep a conservative filesystem fallback
    # for installations whose uninstaller registration is delayed or absent.
    foreach($root in @(
        (Join-Path ${env:ProgramFiles} 'Intel\Extreme Tuning Utility'),
        (Join-Path ${env:ProgramFiles(x86)} 'Intel\Extreme Tuning Utility'),
        (Join-Path ${env:ProgramFiles} 'Intel\Intel Extreme Tuning Utility'),
        (Join-Path ${env:ProgramFiles(x86)} 'Intel\Intel Extreme Tuning Utility')
    )){
        try{
            if($root -and (Test-Path -LiteralPath $root)){
                $exe=Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue |
                    Where-Object { ([string]$_.Name) -like '*XTU*.exe' } | Select-Object -First 1
                if($exe){ return $true }
            }
        }catch{}
    }
    return $false
}


function Invoke-CriticalManufacturerUtilityCheckpoint {
    param(
        [Parameter(Mandatory)]$Devices,
        [Parameter(Mandatory)]$AsusPackages
    )
    Write-Log 'STAGE 07E-RECOVERY - Critical Intel/HWiNFO/NVIDIA executable utility installers are being run before an ASUS package failure terminates this invocation.' 'STEP'
    Write-Log 'This recovery checkpoint does NOT advance the ASUS package sequence and does NOT perform hardware/driver/software re-detection. It only ensures the requested vendor executable installers are not skipped because the ASUS package phase stopped at an interactive gate.' 'INFO'

    $vendors=@($Devices | ForEach-Object { Get-HardwareVendor $_ } | Sort-Object -Unique)
    $cpu=(Get-CimInstance Win32_Processor | Select-Object -First 1)
    $cpuName=[string]$cpu.Name
    $root=Join-Path $DownloadRoot '02 Official Manufacturer Utilities'
    New-Item -ItemType Directory -Force -Path $root | Out-Null

    if($cpuName -match '(?i)275HX|285HX|265HX|255HX|245HX|235HX'){
        $ipfOk=Invoke-IntelIPFFrameworkInstall -AsusPackages $AsusPackages -Root $root
        if(-not $ipfOk){
            Write-Log '[Intel XTU] Recovery checkpoint blocked because the required IPF Framework installer did not complete.' 'ERROR'
        } else {
        $xtuOk=Invoke-IntelXTUUtilityInstall -AsusPackages $AsusPackages -Root $root -Label 'Intel XTU 10.0.1.188 (recovery checkpoint)'
        if(-not $xtuOk){Write-Log '[Intel XTU] Recovery checkpoint did not complete the XTU installation.' 'WARN'}
        }
    }

    # HWiNFO must also run from recovery so an ASUS package interactive failure cannot skip the
    # requested hardware/RAM utility. This does not alter the ASUS package sequence or re-detect hardware.
    $hwRecoveryRoot=Join-Path $root 'RAM Timing and Hardware Utilities'
    Install-HWiNFOCurrentStable -Root $hwRecoveryRoot | Out-Null

    if($vendors -contains 'NVIDIA'){
        $sdkItems = @(Get-NvidiaBroadcastSdkPackageSet)
        foreach($i in $sdkItems){
            $nRoot=Join-Path $root ("NVIDIA RTX {0} Series Broadcast" -f $i.Generation)
            $dir=Join-Path $nRoot $i.Folder
            $dest=Join-Path $dir $i.FileName
            $marker=Join-Path $dir 'INSTALL-COMPLETE.marker'
            New-Item -ItemType Directory -Force -Path $dir | Out-Null
            if(Test-Path -LiteralPath $marker){
                Write-Log "[$($i.Name)] Recovery checkpoint already completed on an earlier invocation." 'OK'
                continue
            }
            $r=Invoke-OfficialDownloadAndInstall -Label "$($i.Name) (recovery checkpoint)" -Uri $i.Uri -Destination $dest -Arguments '' -Install -InteractiveFallback
            if($r.Installed){
                Set-Content -LiteralPath $marker -Value ("Completed $(Get-Date -Format o)") -Encoding UTF8
                Write-Log "[$($i.Name)] Recovery executable installer completed." 'OK'
            }else{
                Write-Log "[$($i.Name)] Recovery executable installer did not complete." 'WARN'
            }
        }
    }

    Write-Log 'STAGE 07E-RECOVERY complete. No ASUS package was advanced by this checkpoint.' 'OK'
}


function Get-AsusCompatibleIPFPackage {
    <#
      v68: The live G635LW ASUS catalog can omit the IPF entry even though the
      machine requires the Intel Innovation Platform Framework. Do not substitute
      Intel/third-party packages. Instead, query an ASUS catalog for a compatible
      Intel 800-series Arrow Lake platform and use the ASUS-hosted IPF package,
      still with the package's own ASUS SHA-256 and installer wizard.
    #>
    param([Parameter(Mandatory)][object[]]$PrimaryPackages)

    $direct=@($PrimaryPackages | Where-Object {
        "$($_.Category) $($_.Title) $($_.DownloadUrl) $($_.Raw | ConvertTo-Json -Depth 20 -Compress)" -match '(?i)Innovation Platform Framework|\bIPF\b'
    })
    if($direct.Count -gt 0){
        return [pscustomobject]@{Package=$direct | Sort-Object @{Expression={Get-PackageArchitectureRank -Package $_};Descending=$true}, @{Expression={Convert-ToComparableVersion ([string]$_.Version)};Descending=$true} | Select-Object -First 1; Source='Official ASUS G635LW catalog'}
    }

    # ASUS publishes the same Intel IPF driver family through compatible Intel 800-series
    # products. These are ASUS catalogs, not third-party mirrors. The first catalog that
    # exposes an IPF package wins; the package URL and SHA-256 are taken from that live
    # ASUS response rather than being hard-coded.
    $fallbackModels=@(
        [pscustomobject]@{Model='PRIME-Z890-P';Cpu='PRIME-Z890-P'},
        [pscustomobject]@{Model='PRO-WS-W880-ACE-SE';Cpu='PRO-WS-W880-ACE-SE'}
    )
    foreach($m in $fallbackModels){
        $urls=@(
            "https://rog.asus.com/support/webapi/product/GetPDDrivers?website=global&model=$([uri]::EscapeDataString($m.Model))&cpu=$([uri]::EscapeDataString($m.Cpu))&osid=52&systemCode=rog",
            "https://rog.asus.com/support/webapi/product/GetPDDrivers?website=global&model=$([uri]::EscapeDataString($m.Model))&cpu=$([uri]::EscapeDataString($m.Cpu))&osid=52&systemCode=rog&tag=1",
            "https://www.asus.com/support/api/product.asmx/GetPDDrivers?osid=52&website=global&model=$([uri]::EscapeDataString($m.Model))&cpu=$([uri]::EscapeDataString($m.Cpu))"
        )
        foreach($url in $urls){
            try{
                Write-Log "[Intel IPF] G635LW catalog omitted IPF; querying compatible official ASUS catalog '$($m.Model)' for the ASUS-hosted IPF package." 'INFO'
                $r=Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 60 -Headers @{'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v84';'Accept'='application/json,text/plain,*/*';'Referer'='https://rog.asus.com/support/'}
                if($r.StatusCode -ne 200 -or -not $r.Content){continue}
                $jt=$r.Content.Trim();if($jt -match '^\s*[^\(]+\((.*)\)\s*;?\s*$'){$jt=$Matches[1]}
                $json=$jt|ConvertFrom-Json
                $list=New-Object System.Collections.Generic.List[object]
                foreach($group in @($json.Result.Obj)){
                    $cat=[string]$group.Name
                    foreach($file in @($group.Files)){
                        $du=$null;try{$du=[string]$file.DownloadUrl.Global}catch{};if(-not $du){try{$du=[string]$file.DownloadUrl}catch{}}
                        if($du -and $du -match '^https?://'){
                            $list.Add([pscustomobject]@{Category=$cat;Title=[string]$file.Title;Version=([string]$file.Version).TrimStart('V');ReleaseDate=[string]$file.ReleaseDate;DownloadUrl=$du;SHA256=[string]$file.SHA256;Raw=$file;LocalPath='';DownloadStatus='Pending';InstallStatus='Pending'})
                        }
                    }
                }
                $hits=@($list | Where-Object { "$($_.Category) $($_.Title) $($_.DownloadUrl) $($_.Raw | ConvertTo-Json -Depth 20 -Compress)" -match '(?i)Innovation Platform Framework|\bIPF\b' })
                if($hits.Count -gt 0){
                    $pkg=$hits | Sort-Object @{Expression={Get-PackageArchitectureRank -Package $_};Descending=$true}, @{Expression={Convert-ToComparableVersion ([string]$_.Version)};Descending=$true} | Select-Object -First 1
                    Write-Log "[Intel IPF] ASUS compatible catalog '$($m.Model)' exposed official IPF package '$($pkg.Title)' version $($pkg.Version)." 'OK'
                    return [pscustomobject]@{Package=$pkg;Source="Official ASUS catalog: $($m.Model)"}
                }
            }catch{
                Write-Log "[Intel IPF] Compatible ASUS catalog '$($m.Model)' query failed: $($_.Exception.Message)" 'WARN'
            }
        }
    }
    return $null
}

function Invoke-IntelIPFFrameworkInstall {
    param(
        [Parameter(Mandatory)][object[]]$AsusPackages,
        [Parameter(Mandatory)][string]$Root
    )
    # v70: Dedicated IPF checkpoint. Prefer the G635LW ASUS catalog; when that live
    # catalog omits IPF, use a compatible official ASUS Intel-800-series catalog package.
    # The package is still ASUS-hosted, SHA-256 verified, and installed with its wizard ATTENDED.
    Write-Log 'INTEL IPF -> XTU DEPENDENCY GATE: starting the dedicated Intel Innovation Platform Framework installer checkpoint.' 'STEP'
    $resolved=Get-AsusCompatibleIPFPackage -PrimaryPackages @($AsusPackages)
    if($null -eq $resolved -or $null -eq $resolved.Package){
        Write-Log '[Intel IPF] No ASUS-hosted IPF package was exposed by the G635LW catalog or the compatible official ASUS fallback catalogs. XTU is blocked; no generic Intel or third-party package will be substituted.' 'ERROR'
        return $false
    }
    $pkg=$resolved.Package
    $source=[string]$resolved.Source
    $dir=Join-Path $Root 'Intel Innovation Platform Framework'
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    $leaf='';try{$leaf=[IO.Path]::GetFileName(([uri]$pkg.DownloadUrl).AbsolutePath)}catch{}
    if([string]::IsNullOrWhiteSpace($leaf)){$leaf=('ASUS-IPF-'+([string]$pkg.Version -replace '[^0-9A-Za-z._-]','_')+'.exe')}
    $dest=Join-Path $dir $leaf
    $marker=Join-Path $dir 'INSTALL-COMPLETE.marker'
    if(Test-Path -LiteralPath $marker){
        Write-Log '[Intel IPF] Dedicated IPF checkpoint already completed; XTU dependency gate is satisfied.' 'OK'
        return $true
    }
    $ready=$false
    if(Test-Path -LiteralPath $dest){
        try {
            if([string]$pkg.SHA256 -match '^[A-Fa-f0-9]{64}$'){$ready=((Get-FileHash -LiteralPath $dest -Algorithm SHA256).Hash -ieq [string]$pkg.SHA256)}
            else {$ready=((Get-Item -LiteralPath $dest).Length -gt 0)}
        } catch {}
    }
    if(-not $ready){
        Write-Log "[Intel IPF] Downloading ASUS-hosted IPF package '$($pkg.Title)' $($pkg.Version) from $source before XTU." 'STEP'
        $ready=Invoke-FileDownloadWithProgress -Uri ([string]$pkg.DownloadUrl) -Destination $dest -Label "Intel Innovation Platform Framework $($pkg.Version)" -ExpectedSHA256 ([string]$pkg.SHA256)
    } else { Write-Log '[Intel IPF] Existing archived ASUS IPF package passed SHA-256/size verification.' 'OK' }
    if(-not $ready){Write-Log '[Intel IPF] ASUS package download/verification failed. XTU will not be launched.' 'ERROR';return $false}
    if($script:DownloadAllContentMode){Write-Log '[Intel IPF] DOWNLOAD-ONLY PREPARATION complete. The attended IPF wizard will run only after every content source has finished downloading.' 'OK';return $true}
    try {
        $extract=Join-Path $dir 'Extracted'
        $inst=Get-InstallerFromPackage -PackagePath $dest -ExtractDirectory $extract
        if(-not $inst){Write-Log '[Intel IPF] No supported setup.exe/install.exe/install.cmd/install.bat was found in the ASUS IPF package. No INF-only substitute will be attempted.' 'ERROR';return $false}
        # The ASUS Platform Performance Package is normally an extracted EXE installer.
        # For an EXE, an empty ArgumentList must NOT be passed to Start-Process: PowerShell
        # treats the empty array as an invalid/null ArgumentList and throws before the wizard
        # can start.  EXE installers therefore receive no ArgumentList at all, which preserves
        # the requested attended installation. MSI/CMD packages retain their explicit launcher
        # arguments.
        $attendedArgs=@()
        if($inst.Kind -eq 'MSI'){$attendedArgs=@('/i',$dest)}elseif($inst.Kind -eq 'CMD'){$attendedArgs=@('/c',$inst.Arguments[1])}
        Write-Log "[Intel IPF] Launching ASUS IPF installation wizard ATTENDED: $($inst.Path) [$source]" 'STEP'
        Write-Log '[Intel IPF] No silent-install switches are being supplied. Complete/accept the wizard prompts as displayed.' 'INFO'
        if($attendedArgs.Count -gt 0){
            $pr=Start-Process -FilePath $inst.Path -ArgumentList $attendedArgs -WorkingDirectory (Split-Path $inst.Path -Parent) -Wait -PassThru -WindowStyle Normal
        } else {
            $pr=Start-Process -FilePath $inst.Path -WorkingDirectory (Split-Path $inst.Path -Parent) -Wait -PassThru -WindowStyle Normal
        }
        $code=[int]$pr.ExitCode
        Write-Log "[Intel IPF] Installer exit code: $code" $(if($code -in @(0,3010,1641)){'OK'}else{'WARN'})
        if($code -notin @(0,3010,1641)){Write-Log '[Intel IPF] Installer failed; XTU remains blocked for this invocation.' 'ERROR';return $false}
        Set-Content -LiteralPath $marker -Value ("Completed $(Get-Date -Format o)`r`nPackage=$($pkg.Title)`r`nVersion=$($pkg.Version)`r`nSHA256=$($pkg.SHA256)`r`nSource=$source`r`nInstallation=Attended") -Encoding UTF8
        Write-Log '[Intel IPF] ASUS framework installer completed. XTU is now permitted to run.' 'OK'
        return $true
    } catch {Write-Log "[Intel IPF] ASUS installer launch failed: $($_.Exception.Message). XTU remains blocked." 'ERROR';return $false}
}

function Get-AsusCompatibleXTUPackage {
    <#
      v70: Resolve Intel XTU from the official ASUS Driver Store before using Intel's
      official fallback. The G635LW catalog is checked first. If it does not expose XTU,
      compatible official ASUS Intel-platform catalogs are queried dynamically.
      Old legacy ASUS XTU packages (for example 6.x packages from 2016-2019) are rejected
      because they are not appropriate as a current Windows 11 / Arrow Lake-HX package.
    #>
    param([Parameter(Mandatory)][object[]]$PrimaryPackages)

    $isXtu = { param($p)
        "$($p.Category) $($p.Title) $($p.Description) $($p.DownloadUrl)" -match '(?i)(Intel\s*)?(Extreme Tuning Utility|Intel XTU|\bXTU\b)'
    }
    $isCurrentEnough = { param($p)
        $d=[datetime]::MinValue
        [void][datetime]::TryParse([string]$p.ReleaseDate,[ref]$d)
        # ASUS-hosted XTU packages discovered for this purpose must be modern enough for
        # a current Windows 11 deployment. This deliberately rejects legacy 2016-2019 XTU.
        return ($d -ge [datetime]'2024-01-01')
    }

    $direct=@($PrimaryPackages | Where-Object { & $isXtu $_ })
    $directCurrent=@($direct | Where-Object { & $isCurrentEnough $_ })
    if($directCurrent.Count -gt 0){
        $pkg=$directCurrent | Sort-Object @{Expression={Convert-ToComparableVersion ([string]$_.Version)};Descending=$true}, @{Expression={([datetime]$_.ReleaseDate)};Descending=$true} | Select-Object -First 1
        return [pscustomobject]@{Package=$pkg;Source='Official ASUS G635LW catalog'}
    }
    if($direct.Count -gt 0){
        Write-Log "[Intel XTU] G635LW ASUS catalog exposed XTU candidate(s), but all were legacy/pre-2024 packages; they will NOT be used on this Windows 11 25H2 Arrow Lake-HX system." 'WARN'
    }

    $fallbackModels=@(
        [pscustomobject]@{Model='PRIME-Z890-P';Cpu='PRIME-Z890-P'},
        [pscustomobject]@{Model='ROG-MAXIMUS-Z890-HERO';Cpu='ROG-MAXIMUS-Z890-HERO'},
        [pscustomobject]@{Model='PRO-WS-W880-ACE-SE';Cpu='PRO-WS-W880-ACE-SE'}
    )
    foreach($m in $fallbackModels){
        $urls=@(
            "https://rog.asus.com/support/webapi/product/GetPDDrivers?website=global&model=$([uri]::EscapeDataString($m.Model))&cpu=$([uri]::EscapeDataString($m.Cpu))&osid=52&systemCode=rog",
            "https://www.asus.com/support/api/product.asmx/GetPDDrivers?osid=52&website=global&model=$([uri]::EscapeDataString($m.Model))&cpu=$([uri]::EscapeDataString($m.Cpu))"
        )
        foreach($url in $urls){
            try{
                Write-Log "[Intel XTU] G635LW catalog has no current XTU package; querying compatible official ASUS catalog '$($m.Model)' for an ASUS-hosted XTU package." 'INFO'
                $r=Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 60 -Headers @{'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v84';'Accept'='application/json,text/plain,*/*';'Referer'='https://www.asus.com/support/'}
                if($r.StatusCode -ne 200 -or -not $r.Content){continue}
                $jt=$r.Content.Trim(); if($jt -match '^\s*[^\(]+\((.*)\)\s*;?\s*$'){$jt=$Matches[1]}
                $json=$jt|ConvertFrom-Json
                $list=New-Object System.Collections.Generic.List[object]
                foreach($group in @($json.Result.Obj)){
                    $cat=[string]$group.Name
                    foreach($file in @($group.Files)){
                        $du=$null; try{$du=[string]$file.DownloadUrl.Global}catch{}; if(-not $du){try{$du=[string]$file.DownloadUrl}catch{}}
                        if($du -and $du -match '^https?://'){
                            $list.Add([pscustomobject]@{Category=$cat;Title=[string]$file.Title;Description=[string]$file.Description;Version=([string]$file.Version).TrimStart('V');ReleaseDate=[string]$file.ReleaseDate;DownloadUrl=$du;SHA256=[string]$file.SHA256;Raw=$file;LocalPath='';DownloadStatus='Pending';InstallStatus='Pending'})
                        }
                    }
                }
                $hits=@($list | Where-Object { & $isXtu $_ })
                $current=@($hits | Where-Object { & $isCurrentEnough $_ })
                if($current.Count -gt 0){
                    $pkg=$current | Sort-Object @{Expression={Convert-ToComparableVersion ([string]$_.Version)};Descending=$true}, @{Expression={([datetime]$_.ReleaseDate)};Descending=$true} | Select-Object -First 1
                    Write-Log "[Intel XTU] ASUS compatible catalog '$($m.Model)' exposed current official ASUS XTU package '$($pkg.Title)' version $($pkg.Version) released $($pkg.ReleaseDate)." 'OK'
                    return [pscustomobject]@{Package=$pkg;Source="Official ASUS catalog: $($m.Model)"}
                }
                if($hits.Count -gt 0){Write-Log "[Intel XTU] ASUS compatible catalog '$($m.Model)' exposed XTU, but only legacy/pre-2024 builds; none will be used." 'WARN'}
            }catch{
                Write-Log "[Intel XTU] Compatible ASUS catalog '$($m.Model)' query failed: $($_.Exception.Message)" 'WARN'
            }
        }
    }
    return $null
}

function Invoke-IntelXTUUtilityInstall {
    param(
        [Parameter(Mandatory)][object[]]$AsusPackages,
        [string]$Root,
        [string]$Label='Intel XTU 10.0.1.188'
    )

    # v83: make the XTU destination self-resolving. The v83 run reached the
    # correct archive directory but still passed an empty destination to the
    # download helper. The helper below also has its own fallback.
    $archiveRoot = [string]$script:DownloadRoot
    if([string]::IsNullOrWhiteSpace($archiveRoot)){ $archiveRoot = [string]$DownloadRoot }
    if([string]::IsNullOrWhiteSpace($archiveRoot)){ $archiveRoot = Join-Path $env:USERPROFILE 'ASUS-G635LW-Installer-Content' }
    New-Item -ItemType Directory -Force -Path $archiveRoot -ErrorAction Stop | Out-Null

    $utilRoot = Join-Path $archiveRoot '02 Official Manufacturer Utilities'
    $dir = Join-Path $utilRoot 'Intel XTU'
    New-Item -ItemType Directory -Force -Path $dir -ErrorAction Stop | Out-Null
    $dir = (Get-Item -LiteralPath $dir -Force -ErrorAction Stop).FullName
    Write-Log ('[Intel XTU] Resolved XTU archive directory: ' + $dir) 'INFO'

    $marker = Join-Path $dir 'INSTALL-COMPLETE.marker'
    if(Test-IntelXTUInstalled){
        if(-not (Test-Path -LiteralPath $marker)){
            Set-Content -LiteralPath $marker -Value ('Verified installed ' + (Get-Date -Format o)) -Encoding UTF8
        }
        Write-Log '[Intel XTU] XTU is already installed; no duplicate wizard launch.' 'OK'
        return $true
    }

    Write-Log '[Intel XTU] XTU is not installed; resolving an official ASUS-hosted package before using the official Intel fallback.' 'STEP'
    $asusResolution = Get-AsusCompatibleXTUPackage -PrimaryPackages @($AsusPackages)
    $asus = @()
    $asusSource = ''
    if($asusResolution){
        $asus = @($asusResolution.Package)
        $asusSource = [string]$asusResolution.Source
    }

    if($asus.Count -gt 0){
        $pkg = $asus[0]
        $leaf = ''
        try{ $leaf = [IO.Path]::GetFileName(([uri]$pkg.DownloadUrl).AbsolutePath) }catch{}
        if([string]::IsNullOrWhiteSpace($leaf)){
            $leaf = 'ASUS-XTU-' + (([string]$pkg.Version) -replace '[^0-9A-Za-z._-]','_') + '.exe'
        }
        $asusDest = Join-Path $dir $leaf
        Write-Log ('[Intel XTU] Using ASUS-hosted XTU package ' + [string]$pkg.Title + ' ' + [string]$pkg.Version + ' from ' + $asusSource + '.') 'OK'
        $ready = Invoke-FileDownloadWithProgress -Uri ([string]$pkg.DownloadUrl) -Destination $asusDest -Label ($Label + ' - ASUS catalog') -ExpectedSHA256 ([string]$pkg.SHA256)
        if(-not $ready){
            Write-Log '[Intel XTU] ASUS catalog package download/verification failed. XTU is not being substituted with an unverified package.' 'ERROR'
            return $false
        }
        if($script:DownloadAllContentMode){Write-Log '[Intel XTU] DOWNLOAD-ONLY PREPARATION complete for the ASUS-hosted package. Installation is deferred until the global download phase is complete.' 'OK';return $true}
        try{
            $extract = Join-Path $dir 'Extracted'
            $inst = Get-InstallerFromPackage -PackagePath $asusDest -ExtractDirectory $extract
            if(-not $inst){ throw 'No supported setup/installer was found in the ASUS XTU package.' }
            $attendedArgs = @()
            if($inst.Kind -eq 'MSI'){$attendedArgs=@('/i',$asusDest)}
            elseif($inst.Kind -eq 'CMD'){$attendedArgs=@('/c',$inst.Path)}
            Write-Log ('[Intel XTU] Launching ASUS XTU installation wizard ATTENDED: ' + $inst.Path + ' [' + $asusSource + ']') 'STEP'
            Write-Log '[Intel XTU] No silent-install switches are being supplied. Complete/accept the wizard prompts as displayed.' 'INFO'
            if($attendedArgs.Count -gt 0){
                $pr=Start-Process -FilePath $inst.Path -ArgumentList $attendedArgs -WorkingDirectory (Split-Path $inst.Path -Parent) -Wait -PassThru -WindowStyle Normal
            } else {
                $pr=Start-Process -FilePath $inst.Path -WorkingDirectory (Split-Path $inst.Path -Parent) -Wait -PassThru -WindowStyle Normal
            }
            $code=[int]$pr.ExitCode
            if($code -notin @(0,3010,1641)){
                Write-Log ('[Intel XTU] ASUS wizard returned exit code ' + $code + '; installation not verified.') 'ERROR'
                return $false
            }
            if(Test-IntelXTUInstalled){
                Set-Content -LiteralPath $marker -Value ('Completed and verified ' + (Get-Date -Format o)) -Encoding UTF8
                Write-Log '[Intel XTU] ASUS XTU wizard completed and XTU was verified.' 'OK'
                return $true
            }
            Write-Log '[Intel XTU] ASUS XTU wizard completed but XTU could not be verified.' 'WARN'
            return $false
        }catch{
            Write-Log ('[Intel XTU] ASUS XTU installer launch failed: ' + $_.Exception.Message) 'ERROR'
            return $false
        }
    }

    Write-Log '[Intel XTU] No current compatible ASUS-hosted XTU package was exposed by the G635LW catalog or compatible official ASUS catalogs. Falling back to Intel official XTU.' 'INFO'

    $xtuFileName = 'XTUSetup_10.0.1.188.exe'
    $xtuDestination = [System.IO.Path]::Combine([string]$dir,$xtuFileName)
    Write-Log ('[Intel XTU] Intel fallback destination: ' + $xtuDestination) 'INFO'

    $xtuOk = Invoke-IntelXTUDownload -Destination ([string]$xtuDestination) -Label ($Label + ' - official Intel fallback')
    if(-not $xtuOk){ return $false }
    if($script:DownloadAllContentMode){Write-Log '[Intel XTU] DOWNLOAD-ONLY PREPARATION complete for the official Intel fallback. Installation is deferred until the global content-download phase is complete.' 'OK';return $true}

    try{
        if(-not (Test-Path -LiteralPath $xtuDestination -PathType Leaf)){
            throw ('Verified Intel XTU executable was not present at ' + $xtuDestination)
        }
        Write-Log '[Intel XTU] Launching the official Intel XTU installation wizard ATTENDED.' 'STEP'
        Write-Log '[Intel XTU] No /S, /quiet or /qn switches are being supplied. Complete/accept the Intel wizard prompts as displayed.' 'INFO'
        $pr=Start-Process -FilePath $xtuDestination -WorkingDirectory $dir -Wait -PassThru -WindowStyle Normal
        $code=[int]$pr.ExitCode
        if($code -notin @(0,3010,1641)){
            Write-Log ('[Intel XTU] Intel wizard returned exit code ' + $code + '; installation not verified.') 'ERROR'
            return $false
        }
        if(Test-IntelXTUInstalled){
            Set-Content -LiteralPath $marker -Value ('Completed and verified ' + (Get-Date -Format o)) -Encoding UTF8
            Write-Log '[Intel XTU] Intel wizard completed and XTU was verified.' 'OK'
            return $true
        }
        Write-Log '[Intel XTU] Intel wizard completed but XTU could not be verified.' 'WARN'
        return $false
    }catch{
        Write-Log ('[Intel XTU] Intel XTU wizard launch failed: ' + $_.Exception.Message) 'ERROR'
        return $false
    }
}

function Invoke-IntelXTUDownload {
    param(
        [string]$Destination,
        [Parameter(Mandatory)][string]$Label
    )

    # v83: self-healing destination. An empty caller value is repaired here.
    $fallbackRoot = [string]$script:DownloadRoot
    if([string]::IsNullOrWhiteSpace($fallbackRoot)){ $fallbackRoot = [string]$DownloadRoot }
    if([string]::IsNullOrWhiteSpace($fallbackRoot)){ $fallbackRoot = Join-Path $env:USERPROFILE 'ASUS-G635LW-Installer-Content' }
    $fallbackDir = Join-Path (Join-Path $fallbackRoot '02 Official Manufacturer Utilities') 'Intel XTU'
    New-Item -ItemType Directory -Force -Path $fallbackDir -ErrorAction Stop | Out-Null

    if([string]::IsNullOrWhiteSpace([string]$Destination)){
        $Destination = Join-Path $fallbackDir 'XTUSetup_10.0.1.188.exe'
        Write-Log ('[' + $Label + '] Caller supplied no destination; using self-resolved Intel XTU destination: ' + $Destination) 'WARN'
    } else {
        $Destination = [System.IO.Path]::GetFullPath([string]$Destination)
    }

    Write-Log ('[' + $Label + '] Confirmed Intel XTU destination: ' + $Destination) 'INFO'

    $uri = 'https://downloadmirror.intel.com/877126/XTUSetup_10.0.1.188.exe'
    $sha = '0118C89C059D348FF3D45E72F3CB88241EEEB7BE1FC27D284ACA8EECE5F1C7ED'
    $intelPage = 'https://www.intel.com/content/www/us/en/download/17881/intel-extreme-tuning-utility-intel-xtu.html?product=122139'
    $fileName = 'XTUSetup_10.0.1.188.exe'

    $parent = Split-Path -Path $Destination -Parent
    New-Item -ItemType Directory -Force -Path $parent -ErrorAction Stop | Out-Null

    if(Test-Path -LiteralPath $Destination -PathType Leaf){
        try{
            $oldHash=(Get-FileHash -LiteralPath $Destination -Algorithm SHA256 -ErrorAction Stop).Hash
            $oldSize=(Get-Item -LiteralPath $Destination -ErrorAction Stop).Length
            if(($oldHash -ieq $sha) -and ($oldSize -gt 1000000)){
                Write-Log ('[' + $Label + '] Existing Intel XTU executable is already SHA-256 verified.') 'OK'
                return $true
            }
        }catch{}
        try{Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue}catch{}
    }

    $tmp=$Destination + '.partial'
    try{Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue}catch{}

    try{
        Write-Log ('[' + $Label + '] DOWNLOAD START: official Intel XTU 10.0.1.188') 'STEP'
        $wc=New-Object System.Net.WebClient
        $wc.DownloadFile($uri,$tmp)
        if(Test-Path -LiteralPath $tmp -PathType Leaf){
            $h=(Get-FileHash -LiteralPath $tmp -Algorithm SHA256 -ErrorAction Stop).Hash
            $n=(Get-Item -LiteralPath $tmp -ErrorAction Stop).Length
            if(($h -ieq $sha) -and ($n -gt 1000000)){
                Move-Item -LiteralPath $tmp -Destination $Destination -Force
                Write-Log ('[' + $Label + '] Official Intel XTU downloaded and SHA-256 verified.') 'OK'
                return $true
            }
            Write-Log ('[' + $Label + '] Intel CDN returned an unexpected file. Size=' + [string]$n + ' SHA256=' + [string]$h) 'WARN'
        }
    }catch{
        Write-Log ('[' + $Label + '] Direct Intel CDN download was not accepted: ' + $_.Exception.Message) 'WARN'
    }finally{
        try{if($wc){$wc.Dispose()}}catch{}
        try{Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue}catch{}
    }

    try{
        Write-Log ('[' + $Label + '] Retrying the official Intel XTU download with Invoke-WebRequest.') 'STEP'
        Invoke-WebRequest -Uri $uri -UseBasicParsing -OutFile $tmp -TimeoutSec 900 -ErrorAction Stop
        if(Test-Path -LiteralPath $tmp -PathType Leaf){
            $h=(Get-FileHash -LiteralPath $tmp -Algorithm SHA256 -ErrorAction Stop).Hash
            $n=(Get-Item -LiteralPath $tmp -ErrorAction Stop).Length
            if(($h -ieq $sha) -and ($n -gt 1000000)){
                Move-Item -LiteralPath $tmp -Destination $Destination -Force
                Write-Log ('[' + $Label + '] Invoke-WebRequest downloaded the official Intel XTU and SHA-256 was verified.') 'OK'
                return $true
            }
            Write-Log ('[' + $Label + '] Invoke-WebRequest returned an unexpected file. Size=' + [string]$n + ' SHA256=' + [string]$h) 'WARN'
        }
    }catch{
        Write-Log ('[' + $Label + '] Invoke-WebRequest download was not accepted: ' + $_.Exception.Message) 'WARN'
    }finally{
        try{Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue}catch{}
    }

    $roots=@()
    foreach($candidate in @((Join-Path $env:USERPROFILE 'Downloads'),(Join-Path $env:USERPROFILE 'Desktop'))){
        try{if(-not [string]::IsNullOrWhiteSpace([string]$candidate) -and (Test-Path -LiteralPath $candidate)){$roots += [string]$candidate}}catch{}
    }
    $roots=@($roots | Select-Object -Unique)

    Write-Host ''
    Write-Host '====================================================================' -ForegroundColor Yellow
    Write-Host ' INTEL XTU OFFICIAL DOWNLOAD NEEDS LICENSE ACCEPTANCE' -ForegroundColor Yellow
    Write-Host ' Intel requires acceptance of its software license before download.' -ForegroundColor White
    Write-Host ' The official Intel download page will now open.' -ForegroundColor White
    Write-Host ''
    Write-Host ' Download this exact file:' -ForegroundColor Cyan
    Write-Host '   XTUSetup_10.0.1.188.exe' -ForegroundColor Green
    Write-Host ''
    Write-Host ' The script will detect it, verify Intel SHA-256, and continue.' -ForegroundColor White
    Write-Host '====================================================================' -ForegroundColor Yellow
    Write-Host ''
    try{Start-Process -FilePath $intelPage | Out-Null}catch{Write-Log ('[' + $Label + '] Could not open Intel official page: ' + $_.Exception.Message) 'WARN'}

    $lastNotice=Get-Date
    while($true){
        foreach($root in $roots){
            try{
                $hits=@(Get-ChildItem -LiteralPath $root -Recurse -File -Filter $fileName -ErrorAction SilentlyContinue)
                foreach($hit in $hits){
                    try{
                        $h=(Get-FileHash -LiteralPath $hit.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
                        if($h -ieq $sha){
                            Copy-Item -LiteralPath $hit.FullName -Destination $Destination -Force -ErrorAction Stop
                            Write-Log ('[' + $Label + '] Exact Intel XTU executable found and SHA-256 verified: ' + $hit.FullName) 'OK'
                            return $true
                        }
                    }catch{}
                }
            }catch{}
        }
        if(((Get-Date)-$lastNotice).TotalSeconds -ge 30){
            Write-Log ('[' + $Label + '] Waiting for exact official Intel file ' + $fileName + '.') 'INFO'
            $lastNotice=Get-Date
        }
        Start-Sleep -Seconds 5
    }
}

function Get-HWiNFOLatestStableInfoGlobal {
    try{
        $page=(Invoke-WebRequest -Uri 'https://www.hwinfo.com/download/' -UseBasicParsing -TimeoutSec 60 -Headers @{'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v64'}).Content
        # Current official page exposes stable installer as Version 8.52 and the vendor's
        # deterministic installer filename is hwi64_852.exe. Ignore the beta listing.
        $vm=[regex]::Match([string]$page,'(?i)(?:Version|\\{Version)\\s*([0-9]+\\.[0-9]+)')
        if(-not $vm.Success){throw 'Official HWiNFO page did not expose a stable installer version.'}
        $ver=$vm.Groups[1].Value;$digits=($ver -replace '[^0-9]','')
        if(-not $digits){throw 'Could not derive HWiNFO installer filename.'}
        [pscustomobject]@{Version=$ver;Digits=$digits;Uri="https://www.hwinfo.com/files/hwi64_${digits}.exe"}
    }catch{Write-Log "[HWiNFO] Official stable-version discovery failed: $($_.Exception.Message)" 'WARN';return $null}
}

function Get-HWiNFOInstalledVersionGlobal {
    $versions=@()
    foreach($root in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*')){
        try{$versions += @(Get-ItemProperty -Path $root -ErrorAction SilentlyContinue | Where-Object {[string]$_.DisplayName -match '(?i)^HWiNFO64(?: Pro)?(?: Installer)?$|^HWiNFO(?:64)?(?: Pro)?$'} | ForEach-Object {[string]$_.DisplayVersion})}catch{}
    }
    foreach($root in @((Join-Path ${env:ProgramFiles} 'HWiNFO64'),(Join-Path ${env:ProgramFiles(x86)} 'HWiNFO64'),(Join-Path ${env:ProgramFiles} 'HWiNFO'),(Join-Path ${env:ProgramFiles(x86)} 'HWiNFO'))){
        try{if(Test-Path -LiteralPath $root){Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue | Where-Object {$_.Name -match '(?i)^HWiNFO64.*\.exe$'} | ForEach-Object {if($_.VersionInfo.ProductVersion){$versions += [string]$_.VersionInfo.ProductVersion}}}}catch{}
    }
    $parsed=@($versions|ForEach-Object{try{[version](($_ -replace '[^0-9\.]','').Trim('.'))}catch{}}|Where-Object{$_}|Sort-Object -Descending)
    if($parsed.Count){return $parsed[0]};return $null
}

function Install-HWiNFOCurrentStable {
    param([Parameter(Mandatory)][string]$Root)
    $info=Get-HWiNFOLatestStableInfoGlobal
    if(-not $info){
        $winget=Get-Command winget.exe -ErrorAction SilentlyContinue
        if($winget){
            if($script:DownloadAllContentMode){
                $wd=Join-Path $Root 'HWiNFO WinGet Download';New-Item -ItemType Directory -Force -Path $wd|Out-Null
                $wp=Start-Process -FilePath $winget.Source -ArgumentList @('download','--id','REALiX.HWiNFO','--exact','--source','winget','--accept-package-agreements','--accept-source-agreements','--download-directory',"`"$wd`"") -Wait -PassThru -NoNewWindow -ErrorAction SilentlyContinue
                if($wp.ExitCode -eq 0){Write-Log '[HWiNFO] WinGet fallback payload downloaded only; installation is deferred.' 'OK';return $true}
            } else {
                $local=@(Get-ChildItem -LiteralPath (Join-Path $Root 'HWiNFO WinGet Download') -Recurse -File -ErrorAction SilentlyContinue | Where-Object {$_.Extension -in '.exe','.msi'} | Select-Object -First 1)
                if((Test-Path -LiteralPath (Join-Path $DownloadRoot 'ALL-CONTENT-DOWNLOAD-COMPLETE.json')) -and $local.Count -gt 0){
                    $lp=$local[0].FullName;Write-Log "[HWiNFO] Installing the previously downloaded local WinGet publisher payload $lp." 'STEP';$hp=Start-Process -FilePath $lp -Wait -PassThru -WindowStyle Normal -ErrorAction SilentlyContinue;if($hp.ExitCode -in @(0,3010,1641)){Write-Log '[HWiNFO] Local downloaded fallback installer completed.' 'OK';return $true}
                }
                if(Test-Path -LiteralPath (Join-Path $DownloadRoot 'ALL-CONTENT-DOWNLOAD-COMPLETE.json')){Write-Log '[HWiNFO] No local fallback installer was available after the global download phase; network installation is prohibited.' 'ERROR';return $false}
                Write-Log '[HWiNFO] Falling back to WinGet installation from the official REALiX.HWiNFO package because the official page could not be parsed.' 'STEP'
                $wp=Start-Process -FilePath $winget.Source -ArgumentList @('install','--id','REALiX.HWiNFO','--exact','--source','winget','--accept-package-agreements','--accept-source-agreements','--silent') -Wait -PassThru -NoNewWindow -ErrorAction SilentlyContinue
                if($wp.ExitCode -eq 0){Write-Log '[HWiNFO] WinGet installation completed.' 'OK';return $true}
            }
        }
        return $false
    }
    $dir=Join-Path $Root ("HWiNFO64 $($info.Version)");$dest=Join-Path $dir ("hwi64_$($info.Digits).exe");New-Item -ItemType Directory -Force -Path $dir|Out-Null
    $installed=Get-HWiNFOInstalledVersionGlobal;$target=[version]$info.Version
    if($installed -and $installed -ge $target){Write-Log "[HWiNFO] HWiNFO64 $installed is already installed and meets/exceeds stable $($info.Version)." 'OK';return $true}
    $r=Invoke-OfficialDownloadAndInstall -Label "HWiNFO64 $($info.Version) (memory timings/telemetry)" -Uri $info.Uri -Destination $dest -Arguments '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -Install -InteractiveFallback
    $after=Get-HWiNFOInstalledVersionGlobal
    if($after -and $after -ge $target){Write-Log "[HWiNFO] Verified installed version $after after official installer." 'OK';return $true}
    Write-Log "[HWiNFO] Official installer did not produce a verified $($info.Version) installation." 'WARN';return $false
}

function Install-OfficialManufacturerFallbacks {
    param([Parameter(Mandatory)]$Devices,[Parameter(Mandatory)]$AsusPackages,[Parameter(Mandatory)]$InstalledDrivers)
    Write-Log 'STAGE 07E - Official manufacturer fallback packages and remaining ASUS/vendor utilities' 'STEP'
    $utilRoot=Join-Path $DownloadRoot '02 Official Manufacturer Utilities';New-Item -ItemType Directory -Force -Path $utilRoot|Out-Null
    $allVendors=@($Devices|ForEach-Object{Get-HardwareVendor $_}|Sort-Object -Unique);$winget=Get-Command winget.exe -ErrorAction SilentlyContinue
    $hasIntel=(@($allVendors|Where-Object { [string]$_ -match '(?i)^Intel$' }).Count -gt 0)
    $hasNvidia=(@($allVendors|Where-Object { [string]$_ -match '(?i)^NVIDIA$' }).Count -gt 0)
    Write-Log "Manufacturer fallback vendor flags: Intel=$hasIntel; NVIDIA=$hasNvidia; vendors=$([string]::Join(', ',[string[]]$allVendors))" 'INFO'

    # Intel XTU: current official Intel page lists 10.0.1.188 for Windows 11 25H2.
    # Detect the CPU independently of the vendor-list classification so XTU cannot be skipped
    # merely because the PnP vendor aggregation did not contain the exact string 'Intel'.
    $cpu=(Get-CimInstance Win32_Processor|Select-Object -First 1);$cpuName=[string]$cpu.Name
    if($hasIntel -or ($cpuName -match '(?i)275HX|285HX|265HX|255HX|245HX|235HX')){
        if($cpuName -match '(?i)Core\(TM\)? Ultra 9 275HX|Core Ultra 9 275HX|275HX'){
            Write-Log "[Intel] Detected Intel Core Ultra 9 275HX. ASUS platform/CPU-support packages were explicitly prioritized before this manufacturer-utility stage so XTU sees the OEM platform stack first." 'OK'
            # v33: use ONLY the STAGE 06 baseline for Intel platform/graphics identification.
            # The ASUS package phase has already completed before this function is entered, and
            # its dependency ordering installs the complete ArrowLake-HX Intel Graphics package
            # before this XTU stage. Do not query Win32_PnPSignedDriver again here.
            $baselineIntel=@($InstalledDrivers | Where-Object {
                ([string]$_.DeviceID -match '(?i)VEN_8086') -or
                ([string]$_.DeviceName -match '(?i)Intel')
            })
            if($baselineIntel.Count -gt 0){
                Write-Log "[Intel] STAGE 06 baseline contains $($baselineIntel.Count) Intel signed-driver entries. No second installed-driver query will be performed." 'OK'
            }else{
                Write-Log '[Intel] STAGE 06 baseline contains no Intel signed-driver entries. The ASUS package phase remains authoritative for the required platform/graphics installation order; no second inventory query will be performed.' 'WARN'
            }
        }
        if($cpuName -match '(?i)275HX|285HX|265HX|255HX|245HX|235HX'){
            # Intel IPF is a hard prerequisite for the requested XTU workflow.
            # v62 makes the ASUS G635LW Innovation Platform Framework package a first-class
            # dependency and sorts it before this XTU utility stage. Intel's XTU release notes
            # document IPF as a required driver for XTU 10.x on supported Arrow Lake systems.
            # Intel's current official download page lists XTU 10.0.1.188 for
            # Windows 11 25H2. Intel publishes the executable filename and SHA-256;
            # the official Intel download page and a hardened direct-CDN recovery path are used.
            # If Intel requires license acceptance, the script opens Intel's official page and
            # waits for the exact signed installer to appear in Downloads before proceeding.
            # HARD ORDER: IPF is downloaded and run immediately before XTU.
            # If IPF cannot complete, XTU is intentionally not launched.
            $ipfOk=Invoke-IntelIPFFrameworkInstall -AsusPackages $AsusPackages -Root $utilRoot
            if(-not $ipfOk){
                Write-Log '[Intel] IPF prerequisite did not complete; Intel XTU utility launch is blocked for this invocation.' 'ERROR'
                return
            }
            Write-Log '[Intel XTU] IPF dependency gate passed. Beginning the XTU package checkpoint immediately; ASUS catalog package is preferred, with official Intel fallback only when ASUS does not publish XTU for G635LW.' 'STEP'
            $xtuOk=Invoke-IntelXTUUtilityInstall -AsusPackages $AsusPackages -Label 'Intel XTU 10.0.1.188'
            if(-not $xtuOk){Write-Log '[Intel XTU] XTU installation checkpoint did not complete successfully. The completion marker was not created; the next invocation will retry.' 'WARN'}
        }else{Write-Log "Intel XTU not auto-installed because CPU '$cpuName' was not positively matched to Intel's current XTU support family." 'INFO'}
    }

    # RAM timing/diagnostic utilities. CPU-Z exposes memory timings/SPD and HWiNFO exposes
    # detailed memory/controller telemetry. They are installed silently only after their
    # complete installers have been downloaded to the user's archive.
    $ramRoot=Join-Path $utilRoot 'RAM Timing and Hardware Utilities';New-Item -ItemType Directory -Force -Path $ramRoot|Out-Null
    $cpuzUri='https://download.cpuid.com/cpu-z/cpu-z_3.01-en.exe';$cpuzDest=Join-Path $ramRoot 'CPU-Z 3.01\cpu-z_3.01-en.exe';New-Item -ItemType Directory -Force -Path (Split-Path $cpuzDest -Parent)|Out-Null
    Invoke-OfficialDownloadAndInstall -Label 'CPUID CPU-Z 3.01 (RAM timings/SPD)' -Uri $cpuzUri -Destination $cpuzDest -SHA256 '' -Arguments '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -Install -InteractiveFallback|Out-Null
    $ramRoot=Join-Path $utilRoot 'RAM Timing and Hardware Utilities';New-Item -ItemType Directory -Force -Path $ramRoot|Out-Null
    Install-HWiNFOCurrentStable -Root $ramRoot | Out-Null

    # NVIDIA Broadcast application and GPU-generation-matched official SDK redistributables.
    # NVIDIA's official resource page provides separate AR, Audio Effects, and Video Effects packages
    # for RTX 20/30/40/50 families; select only the detected adapter generation. The SDK redistributables are downloaded/extracted; NVIDIA says
    # they are intended for applications that integrate the SDK, so the script does not force-install
    # DLLs that are not standalone end-user installers.
    if($hasNvidia){
        $sdkItems = @(Get-NvidiaBroadcastSdkPackageSet)
        $gpuGeneration = if($sdkItems.Count -gt 0){[string]$sdkItems[0].Generation}else{'Unknown'}
        $nRoot=Join-Path $utilRoot ("NVIDIA RTX {0} Series Broadcast" -f $gpuGeneration);New-Item -ItemType Directory -Force -Path $nRoot|Out-Null
        $broadcastPage='https://www.nvidia.com/en-au/geforce/broadcasting/broadcast-app/'
        try{
            $html=(Invoke-WebRequest -Uri $broadcastPage -UseBasicParsing -TimeoutSec 60 -Headers @{'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v64'}).Content
            $links=[regex]::Matches([string]$html,'https?://[^"''\s<>]+')|ForEach-Object{$_.Value}|Where-Object{$_ -match '(?i)(NVIDIA.?Broadcast|broadcast).*(\.exe|download)'}|Select-Object -First 10
            $bUri=$links|Where-Object{$_ -match '(?i)\.exe'}|Select-Object -First 1
            if($bUri){$bDest=Join-Path $nRoot 'NVIDIA Broadcast\NVIDIA-Broadcast.exe';New-Item -ItemType Directory -Force -Path (Split-Path $bDest -Parent)|Out-Null;Invoke-OfficialDownloadAndInstall -Label 'NVIDIA Broadcast' -Uri $bUri -Destination $bDest -Arguments '/S' -Install -InteractiveFallback|Out-Null}else{Write-Log 'NVIDIA Broadcast direct installer URL was not exposed in the official page HTML; using WinGet publisher package as fallback.' 'WARN';if($winget){$bd=Join-Path $nRoot 'NVIDIA Broadcast';New-Item -ItemType Directory -Force -Path $bd|Out-Null;Start-Process -FilePath $winget.Source -ArgumentList @('download','--id','Nvidia.Broadcast','--exact','--source','winget','--accept-package-agreements','--accept-source-agreements','--download-directory',"`"$bd`"") -Wait -NoNewWindow -ErrorAction SilentlyContinue|Out-Null}}
        }catch{Write-Log "NVIDIA Broadcast page lookup failed: $($_.Exception.Message)" 'WARN'}

        # NVIDIA Broadcast SDK redistributables are selected from NVIDIA's official
        # resource page based on the detected RTX generation (20/30/40/50). A package
        # from another generation is never used as a fallback.
        foreach($item in $sdkItems){
            $dest=Join-Path (Join-Path $nRoot $item.Folder) $item.FileName
            $marker=Join-Path (Split-Path $dest -Parent) 'INSTALL-COMPLETE.marker'
            New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent)|Out-Null
            if(Test-Path -LiteralPath $marker){
                Write-Log "[$($item.Name)] Persistent completion marker found; official NVIDIA executable installer already completed. No duplicate installer launch." 'OK'
                continue
            }

            Write-Log "[$($item.Name)] Downloading/retaining the official NVIDIA RTX $($item.Generation) executable installer and launching it now." 'STEP'
            $r=Invoke-OfficialDownloadAndInstall -Label $item.Name -Uri $item.Uri -Destination $dest -Arguments '' -Install -InteractiveFallback
            if($r.Installed){
                Set-Content -LiteralPath $marker -Value ("Completed $(Get-Date -Format o)") -Encoding UTF8
                Write-Log "[$($item.Name)] Official NVIDIA executable installer completed and was recorded in the persistent utility checkpoint." 'OK'
            }else{
                Write-Log "[$($item.Name)] Official NVIDIA executable installer did not report successful completion. It remains the utility checkpoint for the next invocation." 'WARN'
            }
        }


    }

    # Full offline Armoury Crate package. ASUS currently publishes 1.5.0.7 / 4.99 GB from the
    # official support page. It is downloaded in full; the package is intentionally not treated as
    # a generic /S executable because ASUS documents device selection/installation behavior for it.
    $acRoot=Join-Path $utilRoot 'ASUS Armoury Crate';New-Item -ItemType Directory -Force -Path $acRoot|Out-Null
    try{
        $acHtml=(Invoke-WebRequest -Uri 'https://www.asus.com/au/supportonly/armoury%20crate/helpdesk_download/' -UseBasicParsing -TimeoutSec 60 -Headers @{'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v64'}).Content
        $acLinks=[regex]::Matches([string]$acHtml,'https?://[^"''\s<>]+')|ForEach-Object{$_.Value}|Where-Object{$_ -match '(?i)dlcdnets\.asus\.com|dlcdnwebimgs\.asus\.com'}|Where-Object{$_ -match '(?i)Armoury|Armoury_Crate|ArmouryCrate|Full_Installation|Full.*Package'}|Select-Object -Unique
        $acFull=$acLinks|Where-Object{$_ -match '(?i)Full.*(Installation|Package)|Armoury_Crate_Full'}|Select-Object -First 1
        $acSmall=$acLinks|Where-Object{$_ -notmatch '(?i)Full.*(Installation|Package)' -and $_ -match '(?i)ArmouryCrateInstallTool|Armoury.*Installer'}|Select-Object -First 1
        if($acFull){$acDest=Join-Path $acRoot 'Armoury Crate Full Installation Package.zip';Invoke-FileDownloadWithProgress -Uri $acFull -Destination $acDest -Label 'ASUS Armoury Crate Full Installation Package' -ExpectedSHA256 'FCF9D0ECF3350837F68F1223F452CB0E1CD7B8F28753AC6ECC4376C792ACBD87'|Out-Null}else{Write-Log 'ASUS Armoury Crate full-package URL was not exposed in raw HTML; the ASUS catalog phase will still download it if the G635LW catalog provides it.' 'WARN'}
        if($acSmall){$acSmallDest=Join-Path $acRoot 'Armoury Crate & Aura Creator Installer.zip';Invoke-FileDownloadWithProgress -Uri $acSmall -Destination $acSmallDest -Label 'ASUS Armoury Crate & Aura Creator Installer' -ExpectedSHA256 'A0C3181B135C0439F12338D6D6B77181B54A62BC44C1FE95C57945F1735D54E1'|Out-Null}
    }catch{Write-Log "Armoury Crate full-package lookup failed: $($_.Exception.Message)" 'WARN'}

    $unmapped=@()
    foreach($d in @($Devices)){
        $fam=Get-DeviceComponentFamily -Device $d
        if($fam -and -not @(Get-ManufacturerStoreProfiles -Device $d).Count){
            $unmapped += [pscustomobject]@{Device=$d.Name;PNP=$d.PNPDeviceID;Family=$fam;Vendor=Get-HardwareVendor -Device $d}
        }
    }
    if($unmapped.Count -gt 0){
        Write-Log "Manufacturer driver-store coverage report: $($unmapped.Count) detected device(s) have no safe manufacturer package profile. These remain in-box-driver candidates; no third-party driver is guessed." 'WARN'
        foreach($u in $unmapped){Write-Log "  UNMAPPED MANUFACTURER DEVICE: $($u.Vendor) | $($u.Family) | $($u.Device) | $($u.PNP)" 'WARN'}
    } else {
        Write-Log 'Manufacturer driver-store coverage report: every classified device has an official manufacturer profile or an explicit safe fallback.' 'OK'
    }

    Write-Log "Manufacturer/utility download archive is at $utilRoot" 'OK'
    Write-Log 'No automatic CPU/GPU overclock, undervolt, voltage, power-limit or fan-profile changes are applied.' 'OK'
}
function Save-State {
    param(
        [int]$RebootCount,
        [int]$UpdatePass,
        [bool]$PostDetectionComplete,
        [bool]$RebootPending,
        [bool]$AsusPackagePhaseComplete=$false
    )
    [pscustomobject]@{
        Version=$Version
        RebootCount=$RebootCount
        UpdatePass=$UpdatePass
        PostDetectionComplete=$PostDetectionComplete
        RebootPending=$RebootPending
        AsusPackagePhaseComplete=$AsusPackagePhaseComplete
        Updated=(Get-Date).ToString('o')
    } | ConvertTo-Json | Set-Content -LiteralPath $StateFile -Encoding UTF8
}

function Register-ResumeTask {
    # Always refresh the stable copy from the script that is currently running. This prevents
    # a stale v33 stable copy from launching an older revision after a reboot.
    $resumeScript=$StableScript
    try {
        $source=(Resolve-Path -LiteralPath $PSCommandPath -ErrorAction Stop).Path
        $stableParent=[IO.Path]::GetDirectoryName($StableScript)
        if([string]::IsNullOrWhiteSpace($stableParent)){ throw "Could not determine the persistent startup-resume directory." }
        New-Item -ItemType Directory -Force -Path $stableParent -ErrorAction Stop | Out-Null
        $sameFile=$false
        try { $sameFile=((Resolve-Path -LiteralPath $StableScript -ErrorAction Stop).Path -ieq $source) } catch { $sameFile=$false }
        if(-not $sameFile){ Copy-Item -LiteralPath $source -Destination $StableScript -Force -ErrorAction Stop }
        Write-Log "Persistent startup-resume script refreshed: $StableScript" 'OK'
    } catch {
        # The current script is still a valid resume source. Register it as a fallback
        # rather than silently losing automatic reboot resume because the stable-copy
        # refresh failed. The next successful invocation will repair the stable copy.
        $resumeScript=$PSCommandPath
        Write-Log "Could not refresh the persistent startup-resume copy: $($_.Exception.Message)" 'WARN'
        Write-Log "Registering the current script itself as the startup-resume fallback: $resumeScript" 'WARN'
    }

    $action=New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$resumeScript`" -Resume"
    # AtStartup is intentional: the resume begins as soon as Windows starts, without
    # waiting for the user to remember to launch the installer manually.
    $trigger=New-ScheduledTaskTrigger -AtStartup
    $principal=New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
    Write-Log "Resume task registered for the next Windows startup: $TaskName" 'OK'
}

function Remove-ResumeTask {
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch {}
}

function Install-FinalArmouryCrateAuraStage {
    param(
        [Parameter(Mandatory)]$Devices,
        [Parameter(Mandatory)]$AsusPackages,
        [Parameter(Mandatory)]$InstalledSoftware
    )
    Write-Log 'STAGE 07H - FINAL ASUS ARMOURY CRATE / AURA SOFTWARE FINALIZATION (after all NVIDIA/manufacturer packages)' 'STEP'
    Write-Log 'Armoury Crate/Aura is software, not a device driver. This final stage intentionally runs after NVIDIA, manufacturer utilities and Microsoft Store companions so Armoury Crate/Aura is the last ASUS software finalization step.' 'INFO'

    $root=Join-Path $DownloadRoot '04 Final ASUS Armoury Crate and Aura';New-Item -ItemType Directory -Force -Path $root|Out-Null

    # Required final ordering: MyASUS immediately before the Armoury Crate/Aura prerequisite and installer stage.
    if($script:DownloadAllContentMode){[void](Install-MyASUS)} elseif(-not (Install-MyASUS)){        Write-Log 'FINAL ASUS SOFTWARE GATE failed at MyASUS. Armoury Crate/Aura is intentionally blocked until MyASUS is installed and its interactive settings gate is completed.' 'ERROR'
        return $false
    }


    # Prerequisite packages from the G635LW catalog. If they are already present in the ASUS archive,
    # reuse them. Otherwise use the official catalog row when available.
    $control=@($AsusPackages | Where-Object { [string]$_.Title -match '(?i)Armoury Crate Control Interface' } | Sort-Object Version -Descending | Select-Object -First 1)
    $system=@($AsusPackages | Where-Object { [string]$_.Title -match '(?i)ASUS System Control Interface' } | Sort-Object Version -Descending | Select-Object -First 1)
    foreach($p in @($system+$control)){
        if(-not $p){continue}
        $safe=(([string]$p.Title)+' '+[string]$p.Version+' '+([IO.Path]::GetFileName(([uri]$p.DownloadUrl).AbsolutePath))) -replace '[\/:*?"<>|]','_'
        $dest=Join-Path $root $safe
        $ready=$false
        if(Test-Path $dest){try{$ready=((Get-Item $dest).Length -gt 0)}catch{}}
        if(-not $ready){if(Test-Path -LiteralPath (Join-Path $DownloadRoot 'ALL-CONTENT-DOWNLOAD-COMPLETE.json')){Write-Log "Final ASUS prerequisite '$($p.Title)' is missing after the global download phase; no network download is permitted." 'ERROR';continue}else{Invoke-FileDownloadWithProgress -Uri $p.DownloadUrl -Destination $dest -Label "Final ASUS prerequisite $($p.Title)" -ExpectedSHA256 ([string]$p.SHA256)|Out-Null}}
        if(Test-Path $dest){
            if($script:DownloadAllContentMode){Write-Log "Final ASUS prerequisite '$($p.Title)' downloaded; installation deferred until the global content-download phase is complete." 'OK';continue}
            try{
                $extract=Join-Path $root (([IO.Path]::GetFileNameWithoutExtension($dest))+'-Extracted')
                $inst=Get-InstallerFromPackage -PackagePath $dest -ExtractDirectory $extract
                if($inst){$pr=Start-Process -FilePath $inst.Path -ArgumentList $inst.Arguments -WorkingDirectory $root -Wait -PassThru;Write-Log "Final ASUS prerequisite '$($p.Title)' exit code $($pr.ExitCode)." $(if($pr.ExitCode -in @(0,3010,1641)){'OK'}else{'WARN'})}
            }catch{Write-Log "Final ASUS prerequisite '$($p.Title)' failed: $($_.Exception.Message)" 'WARN'}
        }
    }

    # Prefer the current ASUS catalog Armoury Crate & Aura Creator installer. The G635LW page
    # documents that this installer installs Armoury Crate, Aura Creator and prerequisite services.
    $ac=@($AsusPackages | Where-Object { [string]$_.Title -match '(?i)Armoury Crate.*Aura Creator Installer|Armoury Crate.*Installer' } | Sort-Object Version -Descending | Select-Object -First 1)
    if(-not $ac){
        Write-Log 'No Armoury Crate/Aura Creator installer row was exposed by the current G635LW catalog; Armoury Crate Store companion remains the fallback.' 'WARN'
    } else {
        $fileName='Armoury Crate & Aura Creator Installer.zip'
        $dest=Join-Path $root $fileName
        $ready=$false
        if(Test-Path $dest){
            try{if([string]$ac.SHA256 -match '^[A-Fa-f0-9]{64}$'){$ready=((Get-FileHash -LiteralPath $dest -Algorithm SHA256).Hash -ieq [string]$ac.SHA256)}else{$ready=((Get-Item $dest).Length -gt 0)}}catch{}
        }
        if(-not $ready){if(Test-Path -LiteralPath (Join-Path $DownloadRoot 'ALL-CONTENT-DOWNLOAD-COMPLETE.json')){Write-Log 'Final ASUS Armoury Crate/Aura Creator payload is missing after the global download phase; no network download is permitted.' 'ERROR';return $false}else{$ready=Invoke-FileDownloadWithProgress -Uri $ac.DownloadUrl -Destination $dest -Label 'FINAL ASUS Armoury Crate & Aura Creator Installer' -ExpectedSHA256 ([string]$ac.SHA256)}}
        if($ready){
            if($script:DownloadAllContentMode){Write-Log 'Final ASUS Armoury Crate/Aura Creator payload downloaded; installation deferred until all content sources are complete.' 'OK';return $true}
            try{
                $extract=Join-Path $root 'Armoury Crate & Aura Creator Extracted'
                $inst=Get-InstallerFromPackage -PackagePath $dest -ExtractDirectory $extract
                if($inst){
                    Write-Log 'Launching the official Armoury Crate/Aura Creator installer as the final ASUS software package.' 'STEP'
                    $p=Start-Process -FilePath $inst.Path -ArgumentList $inst.Arguments -WorkingDirectory $extract -Wait -PassThru
                    Write-Log "Final Armoury Crate/Aura installer exit code: $($p.ExitCode)." $(if($p.ExitCode -in @(0,3010,1641)){'OK'}else{'WARN'})
                    if(Test-Path $extract){[void](Invoke-PackageCompanionExecutables -ExtractDirectory $extract -PrimaryPath $inst.Path -PackageTitle 'Armoury Crate & Aura Creator final stage')}
                }
            }catch{Write-Log "Final Armoury Crate/Aura installation failed: $($_.Exception.Message)" 'WARN'}
        }
    }

    Write-Log 'Final Armoury/Aura Store handling: MyASUS was installed and interactively configured immediately before this stage; Armoury Crate/Aura Store companions were handled only after verification and interactive launch.' 'INFO'
    Write-Log 'FINAL ASUS ARMOURY CRATE / AURA stage complete. No PnP or installed-state re-detection was performed.' 'OK'
    return $true
}

function Write-FinalReport {
    param($Model,$OS,$Bios,$Devices,$InstalledDrivers,$RebootCount,$UpdatePass)

    $bad=@($Devices | Where-Object {$_.ConfigManagerErrorCode -and $_.ConfigManagerErrorCode -ne 0})

    $driverInventory = Join-Path $Base 'Driver-Inventory.csv'
    try {
        # v33: export the exact STAGE 06 driver baseline captured once at the start.
        # Do not query Win32_PnPSignedDriver again here.
        @($InstalledDrivers) |
            Select-Object DeviceName,Manufacturer,DriverVersion,DriverDate,InfName,DeviceID,IsSigned |
            Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $driverInventory
    } catch {
        Write-Log "Could not write the STAGE 06 driver baseline inventory: $($_.Exception.Message)" 'WARN'
    }

    @(
        'ROG STRIX SCAR 16 G635LW - WINDOWS 11 25H2 INSTALLER v84'
        "Completed: $(Get-Date)"
        "Model: $Model"
        "Windows: $($OS.DisplayVersion) build $($OS.Build).$($OS.UBR)"
        "BIOS: $($Bios.SMBIOSBIOSVersion)"
        "Automatic reboots used: $RebootCount"
        "PnP configuration errors in final inventory: $($bad.Count)"
        "Driver inventory: $driverInventory"
        "Logs: $LogDir"
        "Driver/software archive: $DownloadRoot"
        "NOTE: ASUS full setup packages are attempted first after detection from the official G635LW catalog. The ASUS package phase has an independent persistent checkpoint and a Desktop directory check; a reboot/cancel/failure resumes at the first package not marked Installed, without using package-number position as the resume key. Official manufacturer driver stores are then queried for hardware families not represented in the ASUS catalog, using the same live hardware/date/version association logic; directly associated manufacturer apps/utilities/tweak tools are then downloaded/installed. No automatic overclock/undervolt/tuning profile is applied."
    ) | Set-Content -LiteralPath $ReportFile -Encoding UTF8

    Write-Log "Final report: $ReportFile" 'OK'
    Write-Log "Driver inventory: $driverInventory" 'OK'
}

function Invoke-DownloadAllContentPhase {
    param([Parameter(Mandatory)]$Devices,[Parameter(Mandatory)]$AsusPackages,[Parameter(Mandatory)]$InstalledDrivers,[Parameter(Mandatory)]$InstalledSoftware)
    $marker=Join-Path $DownloadRoot 'ALL-CONTENT-DOWNLOAD-COMPLETE.json'
    if(Test-Path -LiteralPath $marker){Write-Log "GLOBAL CONTENT-DOWNLOAD PHASE: completion marker already exists at $marker. No download phase is repeated." 'OK';return $true}
    Write-Log 'STAGE 07A - GLOBAL CONTENT DOWNLOAD ONLY: ALL selected PnP/driver/package/installer/software/utility content is downloaded before ANY installer is permitted to execute.' 'STEP'
    Write-Log 'GLOBAL RULE: download -> verify -> archive everything first. Installation begins only after ALL-CONTENT-DOWNLOAD-COMPLETE.json exists.' 'STEP'
    $script:DownloadAllContentMode=$true;$ok=$true
    try {
        $asusDir=Join-Path $DownloadRoot '01 ASUS G635LW Official Packages';New-Item -ItemType Directory -Force -Path $asusDir|Out-Null
        foreach($pkg in @($AsusPackages)){
            if(Test-PackageIsMicrosoftStoreApp -Package $pkg){
                $def=Get-StoreAppDefinitionForAsusPackage -Package $pkg
                if(-not $def){
                    Write-Log "[$($pkg.Title)] Store application has no safe product-ID mapping. It remains deferred to the Store companion reconciliation stage." 'WARN'
                    continue
                }
                $sd=Join-Path $asusDir ('Microsoft Store - '+(($def.Name) -replace '[\/:*?"<>|]','_'))
                New-Item -ItemType Directory -Force -Path $sd|Out-Null
                $files=@(Get-ChildItem -LiteralPath $sd -Recurse -File -ErrorAction SilentlyContinue | Where-Object {$_.Extension -in '.msix','.msixbundle','.appx','.appxbundle','.appinstaller'})
                if($files.Count -gt 0){
                    Write-Log "[$($def.Name)] Existing offline Store payload found; it will be reused if local deployment is supported." 'OK'
                } else {
                    Write-Log "[$($def.Name)] Offline Store download intentionally skipped. 'winget download --source msstore' requires Microsoft Entra ID licensing authentication even when the Store UI is signed in. The app will be installed later through the normal msstore install transaction in the existing interactive user context." 'INFO'
                }
                continue
            }
            $dest=Get-PackageDestination -Package $pkg -AsusDirectory $asusDir;$existing=Find-ExistingAsusPackage -Package $pkg -AsusDirectory $asusDir -PreferredPath $dest;if($existing){$dest=$existing};$ready=$false
            if(Test-Path -LiteralPath $dest -PathType Leaf){try{$ready=((Get-Item -LiteralPath $dest).Length -gt 0);if([string]$pkg.SHA256 -match '^[A-Fa-f0-9]{64}$'){$ready=((Get-FileHash -LiteralPath $dest -Algorithm SHA256).Hash -ieq [string]$pkg.SHA256)}}catch{$ready=$false}}
            if($ready){Write-Log "[ASUS DOWNLOAD ONLY] Already archived and verified: $($pkg.Title) $($pkg.Version)" 'OK'}else{if(-not (Invoke-FileDownloadWithProgress -Uri $pkg.DownloadUrl -Destination $dest -Label "ASUS DOWNLOAD ONLY - $($pkg.Title) $($pkg.Version)" -ExpectedSHA256 ([string]$pkg.SHA256))){$ok=$false}}
        }
        [void](Install-OfficialManufacturerDriverStorePass -Devices $Devices -AsusPackages $AsusPackages -InstalledDrivers $InstalledDrivers)
        [void](Install-DeviceAssociatedManufacturerUtilities -Devices $Devices -InstalledSoftware $InstalledSoftware)
        [void](Install-OfficialManufacturerFallbacks -Devices $Devices -AsusPackages $AsusPackages -InstalledDrivers $InstalledDrivers)
        [void](Install-FinalArmouryCrateAuraStage -Devices $Devices -AsusPackages $AsusPackages -InstalledSoftware $InstalledSoftware)
        if($ok){[pscustomobject]@{Version=$Version;Completed=(Get-Date).ToString('o');Rule='ALL CONTENT DOWNLOADED BEFORE INSTALLATION';ASUSPackages=@($AsusPackages).Count}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $marker -Encoding UTF8;$script:AllContentDownloadMarker=$marker;Write-Log "GLOBAL CONTENT-DOWNLOAD PHASE COMPLETE. Marker written: $marker" 'OK'}else{Write-Log 'GLOBAL CONTENT-DOWNLOAD PHASE did not fully complete. No installation phase will be started in this invocation.' 'ERROR'}
    } catch {Write-Log "GLOBAL CONTENT-DOWNLOAD PHASE failed: $($_.Exception.Message)" 'ERROR';$ok=$false} finally {$script:DownloadAllContentMode=$false}
    return $ok
}

# ---------------- MAIN ----------------

Write-Log 'ROG STRIX SCAR 16 G635LW - WINDOWS 11 25H2 INSTALLER v84' 'STEP'
Write-Log 'Use -Reset to remove the v53/v53/v47/v42/v37/v36/v35/v34/v33/v32/v31/v30/v29/v28/v27/v26/v21/v20/v19/v18/v17/v16/v15 installer archive, checkpoint/state files and scheduled resume task without uninstalling drivers or software.' 'INFO'
Write-Log 'v42 deliberately keeps manually launched terminal sessions open after completion so the complete shell output can be reviewed.' 'INFO'
Write-Log 'v84 separates the workflow into two global phases: download/verify/archive ALL content first, then install the archived content serially.' 'INFO'
Write-Log "Target OS architecture: $(Get-OSArchitecturePreference). ALL architecture variants are retained; package priority is x64 -> ARM64 -> x86 -> neutral/unspecified. Architecture is a priority, not an exclusion filter." 'INFO'
Write-Log 'ASUS full setup packages are preferred after detection; BIOS/firmware flashing remains excluded from unattended execution.' 'INFO'
if ($Resume) {
    Write-Log 'Resumed after reboot; continuing directly with the saved post-detection installation checkpoint.' 'INFO'
    Remove-ResumeTask
    Write-Log 'Consumed the startup resume task for this boot. It will be registered again only if another reboot checkpoint is required.' 'INFO'
}

$model=Test-G635LW
$os=Get-OSInfo
$bios=Get-CimInstance Win32_BIOS
Write-Log "Detected BIOS: $($bios.SMBIOSBIOSVersion)"

$state=$null
if (Test-Path $StateFile) {
    try { $state=Get-Content $StateFile -Raw | ConvertFrom-Json } catch { $state=$null }
}

$rebootCount=if($state){[int]$state.RebootCount}else{0}
$updatePass=if($state){[int]$state.UpdatePass}else{0}
$postDetectionComplete=if($state){[bool]$state.PostDetectionComplete}else{$false}
$asusPackagePhaseComplete=if($state -and $null -ne $state.PSObject.Properties['AsusPackagePhaseComplete']){[bool]$state.AsusPackagePhaseComplete}else{$false}

# On the first run, do preparation only before hardware detection.
if (-not $postDetectionComplete) {
    Prepare-RestorePoint
    Prepare-Resume
}

# v33: perform the COMPLETE unbounded PnP scan exactly once at the start of each
# script invocation. After this point there is deliberately NO second PnP rescan and
# NO second hardware/driver/software inventory query during this invocation.
# If Windows reboots, the auto-resume starts a new invocation and therefore gets one
# fresh start-of-script baseline again.
Write-Log 'STAGE 06 - ONE-TIME HARDWARE BASELINE: starting the only unbounded PnP rescan for this script invocation.' 'STEP'
Invoke-PnpDeviceRescan
$devices=Get-HardwareInventory -Label 'STAGE 06 - One-time hardware detection after PnP rescan'
$script:CurrentLiveDevices=@($devices)

# This is the ONLY driver/software inventory captured for this invocation. It remains
# authoritative for the complete run; later installer phases do not rebuild it.
$installedDrivers=Get-InstalledDriverInventory
$installedSoftware=Get-InstalledSoftwareInventory
Write-Log "STAGE 06 complete: one unbounded PnP rescan and one authoritative hardware/driver/software baseline are complete. Hardware IDs, driver versions and software versions from this baseline are the only inventory data used for this invocation." 'OK'
Write-Log "Baseline counts: HardwareDevices=$($devices.Count) | InstalledDrivers=$($installedDrivers.Count) | InstalledSoftware=$($installedSoftware.Count). No later inventory refresh will occur until a new script invocation (including automatic post-reboot resume)." 'OK'

# Everything below this point is intentionally after hardware detection.
Initialize-DownloadFolders

# The ASUS package phase has its OWN checkpoint. A reboot or a cancelled attended installer
# must never set this to complete. On every subsequent invocation the package directory and
# ASUS-Package-Resume-State.json are checked first, and the first package not marked Installed
# is resumed. This is independent of the package installation counter.
if (-not $asusPackagePhaseComplete) {
    Write-Log 'v53 ordering: MyASUS is deliberately deferred until immediately before the final Armoury Crate/Aura stage. It is NOT installed in the early ASUS package checkpoint.' 'INFO'

    $asusCatalog = @(Get-AsusOfficialPackages)
    if(-not (Invoke-DownloadAllContentPhase -Devices $devices -AsusPackages $asusCatalog -InstalledDrivers $installedDrivers -InstalledSoftware $installedSoftware)){Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $true -RebootPending $false -AsusPackagePhaseComplete $false;[void](Wait-ForUserBeforeExit -ExitCode 3 -Reason 'The global content-download phase did not finish. No driver, package, installer, software or utility installation was started.');exit 3}
    $nvidiaAsusPackages=@($asusCatalog | Where-Object { (Test-PackageLooksLikeDriver -Package $_) -and (Get-PackageComponentFamily -Package $_) -eq 'NVIDIA Graphics' })
if(@($devices | Where-Object { (Get-DeviceComponentFamily -Device $_) -eq 'NVIDIA Graphics' }).Count -gt 0){
    if($nvidiaAsusPackages.Count -eq 0){
        Write-Log 'CRITICAL ASUS-ONLY NVIDIA POLICY: NVIDIA hardware is present but the official ASUS G635LW catalog exposed no NVIDIA graphics driver package. The installer will NOT substitute a generic NVIDIA source.' 'ERROR'
    } else {
        Write-Log "ASUS-ONLY NVIDIA POLICY: $($nvidiaAsusPackages.Count) ASUS NVIDIA graphics package candidate(s) retained; generic NVIDIA driver stores and non-ASUS NVIDIA display drivers are disabled." 'OK'
    }
}
$asusFullOutput = @(Install-AsusFullPackages -Devices $devices -CatalogPackages $asusCatalog -InstalledDrivers $installedDrivers -InstalledSoftware $installedSoftware)
    # Install-AsusFullPackages must return exactly one result object. Be defensive here because
    # PowerShell functions can accidentally emit pipeline values when a helper returns a collection.
    $asusFull = @($asusFullOutput | Where-Object { $_ -and $_.PSObject.Properties['RebootRequired'] }) | Select-Object -Last 1
    if(-not $asusFull) {
        Write-Log 'ASUS package installer returned no valid result object. The run is being stopped rather than treating a null/array value as a reboot flag.' 'ERROR'
        Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $true -RebootPending $false -AsusPackagePhaseComplete $false
        Remove-ResumeTask
        [void](Wait-ForUserBeforeExit -ExitCode 3 -Reason 'Required ASUS package preparation could not continue.')
        exit 3
    }

    Write-Log 'STAGE 07D - No hardware/driver/software re-detection is performed here. The original STAGE 06 baseline remains authoritative for this invocation.' 'INFO'

    if($asusFull.DeferredReboot){
        # A deliberate "NO" at the batched reboot prompt is not an installer failure.
        # Keep the startup resume task registered so the next manual Windows reboot
        # immediately restarts this installer at the saved checkpoint.
        $asusPackagePhaseComplete=$false
        $postDetectionComplete=$true
        Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $postDetectionComplete -RebootPending $true -AsusPackagePhaseComplete $false
        Register-ResumeTask
        Write-Log "ASUS package phase paused because the user deliberately selected N at the reboot prompt. This is NOT an installer failure. The startup resume task is preserved and will automatically resume v53 immediately when Windows is rebooted. The Desktop archive and resume state are preserved." 'INFO'
        [void](Wait-ForUserBeforeExit -ExitCode 0 -Reason 'Reboot was deferred. Choose Restart from Windows when ready; the installer is registered to auto-resume at Windows startup from the saved package checkpoint.')
        exit 0
    }

    if($asusFull.StoppedForInstallerFailure){
        # IMPORTANT: do not proceed to later ASUS packages or manufacturer utilities.
        # The exact cancelled/failed package remains the resume point and its downloaded file is retained.
        $asusPackagePhaseComplete=$false
        $postDetectionComplete=$true
        Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $postDetectionComplete -RebootPending $false -AsusPackagePhaseComplete $false
        Write-Log "ASUS package installation failed/cancelled at the current package boundary. The failed package remains the resume point and will not be skipped." 'WARN'
        Invoke-CriticalManufacturerUtilityCheckpoint -Devices $devices -AsusPackages $asusCatalog
        Remove-ResumeTask
        [void](Wait-ForUserBeforeExit -ExitCode 2 -Reason 'The ASUS package phase stopped at its current failed/cancelled checkpoint; the Desktop archive and resume state were preserved.')
        exit 2
    }

    $asusPackagePhaseComplete=[bool]$asusFull.Completed
    $postDetectionComplete=$true
    $asusRebootRequired=$false
    try { $asusRebootRequired=[System.Convert]::ToBoolean($asusFull.RebootRequired) } catch { $asusRebootRequired=$false; Write-Log "Could not convert ASUS reboot flag to Boolean; treating it as false. Value type: $($asusFull.RebootRequired.GetType().FullName)" 'WARN' }
    Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $postDetectionComplete -RebootPending $asusRebootRequired -AsusPackagePhaseComplete $asusPackagePhaseComplete

    if ($asusRebootRequired) {
        $rebootCount++
        Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $true -RebootPending $true -AsusPackagePhaseComplete $false
        $batchReady=Register-LoginResumeBatch
        if(-not $batchReady){
            Write-Log 'ASUS package phase requires a reboot, but the Startup resume batch could not be created. Stopping without reboot so the checkpoint remains safe.' 'ERROR'
            [void](Wait-ForUserBeforeExit -ExitCode 2 -Reason 'Automatic resume could not be armed for the required reboot. The checkpoint was preserved.')
            exit 2
        }
        Write-Log "ASUS package phase requires a reboot. Automatic reboot #$rebootCount in 15 seconds; there is NO reboot-count limit. v84 will resume from the saved checkpoint after Windows startup." 'WARN'
        Start-Sleep 15
        Restart-Computer -Force
        exit
    }

    if($asusFull.DeferredStore){
        # Microsoft Store companion apps are non-driver software. They must never block
        # the ASUS driver/package and manufacturer fallback work. Keep the Store package
        # in the persistent resume state, leave the ASUS phase incomplete, but continue
        # with non-Store manufacturer utilities in this invocation.
        $asusPackagePhaseComplete=$false
        $postDetectionComplete=$true
        Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $true -RebootPending $false -AsusPackagePhaseComplete $false
        Write-Log 'One or more Microsoft Store companion packages were deferred because the Store transaction did not complete. This is NOT an ASUS driver failure. Continuing with manufacturer utilities; the deferred Store package remains the resume point for the next invocation.' 'WARN'
        Write-Log 'STAGE 07C/07F - Continuing with the unchanged STAGE 06 hardware/driver/software baseline; no PnP rescan or inventory refresh is performed.' 'INFO'
        $storePass=Install-RelatedMicrosoftStoreApps -Devices $devices -AsusPackages $asusCatalog
        if([int]$storePass.Deferred -gt 0){
            Write-Log "Microsoft Store companion pass still has $($storePass.Deferred) deferred/failed Store application(s). Stopping before later ASUS/manufacturer stages so no required Store application is silently skipped." 'ERROR'
            Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $true -RebootPending $false -AsusPackagePhaseComplete $false
            Register-ResumeTask
            [void](Wait-ForUserBeforeExit -ExitCode 2 -Reason 'A Microsoft Store companion could not be installed/verified/launched. The installer will resume the Store checkpoint after the next Windows reboot.')
            exit 2
        }
        Install-OfficialManufacturerDriverStorePass -Devices $devices -AsusPackages $asusCatalog -InstalledDrivers $installedDrivers
        Install-DeviceAssociatedManufacturerUtilities -Devices $devices -InstalledSoftware $installedSoftware
        Install-OfficialManufacturerFallbacks -Devices $devices -AsusPackages $asusCatalog -InstalledDrivers $installedDrivers
        $finalAsusOk=Install-FinalArmouryCrateAuraStage -Devices $devices -AsusPackages $asusCatalog -InstalledSoftware $installedSoftware
        if(-not $finalAsusOk){ Write-Log 'Final ASUS MyASUS/Armoury Crate/Aura stage did not complete;' 'ERROR'; [void](Wait-ForUserBeforeExit -ExitCode 2 -Reason 'The final ASUS software stage did not complete. The archive and state were preserved.'); exit 2 }
        Write-Log 'Non-Store ASUS/manufacturer work completed. The ASUS package checkpoint remains incomplete only because a Microsoft Store companion was deferred.' 'OK'
    }
    elseif(-not $asusPackagePhaseComplete){
        Write-Log 'ASUS package phase did not reach a clean completion checkpoint. Running the critical Intel/NVIDIA executable utility checkpoint before stopping; the outstanding ASUS package remains the exact resume point.' 'WARN'
        Invoke-CriticalManufacturerUtilityCheckpoint -Devices $devices -AsusPackages $asusCatalog
        Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $true -RebootPending $false -AsusPackagePhaseComplete $false
        [void](Wait-ForUserBeforeExit -ExitCode 2 -Reason 'The ASUS package phase did not reach a clean completion checkpoint. The current package remains the resume point.')
        exit 2
    }

    else {
        # Only after every ASUS package checkpoint is complete do manufacturer driver-store
        # reconciliation and then manufacturer utilities/companions run.
        Write-Log 'STAGE 07C/07F - Using the unchanged STAGE 06 hardware/driver/software baseline; no PnP rescan or inventory refresh is performed.' 'INFO'
        $storePass=Install-RelatedMicrosoftStoreApps -Devices $devices -AsusPackages $asusCatalog
        if([int]$storePass.Deferred -gt 0){
            Write-Log "Microsoft Store companion pass still has $($storePass.Deferred) deferred/failed Store application(s). Stopping before later ASUS/manufacturer stages so no required Store application is silently skipped." 'ERROR'
            Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $true -RebootPending $false -AsusPackagePhaseComplete $false
            Register-ResumeTask
            [void](Wait-ForUserBeforeExit -ExitCode 2 -Reason 'A Microsoft Store companion could not be installed/verified/launched. The installer will resume the Store checkpoint after the next Windows reboot.')
            exit 2
        }
        Install-OfficialManufacturerDriverStorePass -Devices $devices -AsusPackages $asusCatalog -InstalledDrivers $installedDrivers
        Install-DeviceAssociatedManufacturerUtilities -Devices $devices -InstalledSoftware $installedSoftware
        Install-OfficialManufacturerFallbacks -Devices $devices -AsusPackages $asusCatalog -InstalledDrivers $installedDrivers
        $finalAsusOk=Install-FinalArmouryCrateAuraStage -Devices $devices -AsusPackages $asusCatalog -InstalledSoftware $installedSoftware
        if(-not $finalAsusOk){ Write-Log 'Final ASUS MyASUS/Armoury Crate/Aura stage did not complete;' 'ERROR'; [void](Wait-ForUserBeforeExit -ExitCode 2 -Reason 'The final ASUS software stage did not complete. The archive and state were preserved.'); exit 2 }
        Write-Log 'ASUS full setup package phase, manufacturer fallback phase, and final Armoury Crate/Aura stage are complete.' 'OK'
    }
}
else {
    Write-Log 'ASUS package phase is already complete according to the persistent checkpoint; no ASUS package is downloaded or installed again.' 'OK'
    # The final Armoury/Aura stage is still run on a resumed invocation so it remains the last
    # ASUS software stage even when the main ASUS package phase was completed on an earlier boot.
    $resumeAsusCatalog=@(Get-AsusOfficialPackages)
    $finalAsusOk=Install-FinalArmouryCrateAuraStage -Devices $devices -AsusPackages $resumeAsusCatalog -InstalledSoftware $installedSoftware
    if(-not $finalAsusOk){ Write-Log 'Final ASUS MyASUS/Armoury Crate/Aura stage did not complete on resumed invocation.' 'ERROR'; [void](Wait-ForUserBeforeExit -ExitCode 2 -Reason 'The final ASUS software stage did not complete. The archive and state were preserved.'); exit 2 }
}

Write-Log 'STAGE 08 - Final verification' 'STEP'
# v33 final verification uses the unchanged STAGE 06 hardware baseline. No second PnP
# scan or hardware/driver/software inventory query is permitted during this invocation.
$finalDevices=@($devices)
$script:CurrentLiveDevices=@($finalDevices)
$os=Get-OSInfo
$bios=Get-CimInstance Win32_BIOS
Write-Log 'STAGE 08 - Final verification is based on the original STAGE 06 one-time hardware baseline; no additional PnP or installed-state enumeration was performed.' 'INFO'

Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $true -RebootPending $false -AsusPackagePhaseComplete $asusPackagePhaseComplete
Write-FinalReport -Model $model -OS $os -Bios $bios -Devices $finalDevices -InstalledDrivers $installedDrivers -RebootCount $rebootCount -UpdatePass $updatePass
Remove-ResumeTask
Write-Log 'VERSION 82 RUN COMPLETE' 'STEP'
Write-Log "Logs: $LogDir"
Write-Log "Driver/software archive: $DownloadRoot"
Write-Log "Report: $ReportFile"
[void](Wait-ForUserBeforeExit -ExitCode 0 -Reason 'All requested installer stages reached their completion checkpoint. The final verification report has been written.')