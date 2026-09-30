#requires -RunAsAdministrator

[CmdletBinding()]
param(
    [switch]$Resume,
    [switch]$Reset
)
# ROG Strix SCAR 16 G635LW - Windows 11 25H2
# Driver & Update Installer v44
#
# v12 adds controlled multi-package reboot batching: a vendor installer that reports
# 3010 (reboot required) is treated as a successful install with a deferred reboot.
# The script can continue through up to six attended packages before requesting a full
# system reboot.  Exit code 1641 remains a hard stop because the installer has actually
# initiated a reboot.  Persistent batch state is keyed to the Windows boot marker so a
# real reboot resets the six-package counter. User-selected 'restart later' choices are
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
# - STAGE 07 performs the main Microsoft Update/driver installation AFTER detection.
# - MyASUS is installed/updated only after STAGE 06.
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
# - Reboots resume at the post-detection update cycle rather than repeating pre-detection work.
# - After hardware detection, ASUS's official G635LW catalog is queried for complete
#   driver/setup packages. Packages are downloaded as provided by ASUS and installed
#   through their setup.exe/install.exe/install.cmd/install.bat when available.
# - BIOS/firmware packages are deliberately excluded from unattended execution.
# - v14 explicitly prioritizes the ASUS Intel Graphics package for the Core Ultra 9 275HX before all other Intel platform packages and before Intel XTU.
# - v16 retains ALL architecture variants and orders them x64 -> ARM64 -> x86 -> neutral/unspecified.
# - The ASUS platform bundle includes chipset, processor power management (PPM), Serial IO, DTT, CSME/ME, VPU, PMT, GNA and Intel SST components when present in the G635LW catalog.
# - No generic Intel CPU package is allowed to overwrite ASUS OEM platform packages merely to make XTU appear; Windows supplies the processor-class driver and ASUS platform packages remain authoritative.

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'Continue'
$script:MicrosoftStorePrepared = $false

# v29 fatal-runtime diagnostics. If a true terminating PowerShell exception occurs,
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
$Version = 'v44'
$Base = Join-Path $env:ProgramData 'ROG-G635LW-Installer-v45'
$LogDir = Join-Path $Base 'Logs'
$StateFile = Join-Path $Base 'state.json'
$ReportFile = Join-Path $Base 'Final-Report.txt'
$StableScript = Join-Path $Base 'ROG-G635LW-Windows11-25H2-Installer-v45.ps1'
$TaskName = 'ROG-G635LW-Installer-v45-Resume'
$MaxPackagesBetweenReboots = 6

New-Item -ItemType Directory -Force -Path $Base,$LogDir | Out-Null
$LogFile = Join-Path $LogDir ("Installer-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$TargetUser = 'possesd'
$TargetUserRoot = Join-Path 'C:\Users' $TargetUser
$MasterDownloadRoot = Join-Path $TargetUserRoot 'ASUS Rog Strix Scar 16 2025 G635LW Drivers and Softwares'
$DesktopDownloadRoot = Join-Path (Join-Path $TargetUserRoot 'Desktop') 'ASUS Rog Strix Scar 16 2025 G635LW Drivers and Softwares'
$DownloadRoot = $MasterDownloadRoot
# The visible archive remains under the user's Desktop as requested. The script never assumes
# that the Desktop path itself is a special object; it is treated as an ordinary filesystem path.
# If OneDrive Files On-Demand marks a downloaded installer online-only, the installer is deferred
# until the file is physically available locally.

function Wait-ForUserBeforeExit {
    param(
        [int]$ExitCode = 0,
        [string]$Reason = 'The installer has finished this run.'
    )

    # Automatic startup-resume runs must never block waiting for keyboard input.
    # Manual/interactive runs remain open so the complete terminal output can be reviewed.
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host ' ROG G635LW INSTALLER v41 - RUN FINISHED' -ForegroundColor Cyan
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
    # It deliberately does NOT uninstall drivers, software, Windows Updates, or change device state.
    Write-Host 'ROG G635LW Installer v42 RESET requested.' -ForegroundColor Yellow
    Write-Host 'This will remove the v42 download archive, resume/checkpoint state, logs/report, and scheduled reboot task (and clean legacy v21/v20/v19/v18/v17/v16/v15 resume state).' -ForegroundColor Yellow
    Write-Host 'Installed drivers and installed software will NOT be removed or changed.' -ForegroundColor Green

    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Scheduled resume task removed: $TaskName" -ForegroundColor Green
    } catch {
        Write-Host "Scheduled task removal warning: $($_.Exception.Message)" -ForegroundColor Yellow
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
    Write-Host 'v42 reset complete. No installed drivers or software were removed.' -ForegroundColor Green
    Write-Host 'A normal v41 run can now be started from the Desktop script.' -ForegroundColor Green
    [void](Wait-ForUserBeforeExit -ExitCode 0 -Reason 'The installer reset completed. Installed drivers and software were not changed.')
    exit 0
}

function Write-Log {
    param([string]$Message,[ValidateSet('INFO','OK','WARN','ERROR','STEP')][string]$Level='INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'),$Level,$Message
    Write-Host $line
    Add-Content -LiteralPath $LogFile -Value $line
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
    Write-Host ' ROG G635LW INSTALLER v41 - TERMINATING ERROR' -ForegroundColor Red
    Write-Host " $trapMessage" -ForegroundColor Red
    Write-Host ' The terminal will remain open so the error and preceding output can be reviewed.' -ForegroundColor Yellow
    Write-Host '============================================================' -ForegroundColor Red
    if(-not $Resume){
        try { [void](Read-Host 'Press ENTER when you have finished reviewing the error output') } catch { Start-Sleep -Seconds 30 }
    }
    exit 1
}

if ($Reset) { Invoke-InstallerReset }

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

        Checkpoint-Computer -Description 'ROG G635LW Installer v41' -RestorePointType MODIFY_SETTINGS -ErrorAction Stop
        Write-Log 'Restore point created successfully.' 'OK'
    } catch {
        Write-Log "Restore point was not created: $($_.Exception.Message). Continuing without aborting the installer." 'WARN'
    }
}

function Enable-MicrosoftUpdateService {
    try {
        $sm = New-Object -ComObject Microsoft.Update.ServiceManager
        $sm.ClientApplicationID = 'ROG-G635LW-Installer-v45'
        $guid = '7971f918-a847-4430-9279-4a52d1efe18d'
        $exists = $false
        foreach ($s in $sm.Services) {
            if ($s.ServiceID -eq $guid) { $exists=$true; break }
        }
        if (-not $exists) {
            $null = $sm.AddService2($guid,7,'')
            Write-Log 'Microsoft Update service registered.' 'OK'
        } else {
            Write-Log 'Microsoft Update service already registered.' 'OK'
        }
    } catch {
        Write-Log "Microsoft Update service registration skipped: $($_.Exception.Message)" 'WARN'
    }
}

function Prepare-UpdateServices {
    Write-Log 'STAGE 04 - Windows Update service preparation (no package download/install)' 'STEP'
    Enable-MicrosoftUpdateService
    try {
        $svc = Get-Service -Name wuauserv -ErrorAction Stop
        if ($svc.Status -ne 'Running') {
            Start-Service -Name wuauserv -ErrorAction Stop
            Write-Log 'Windows Update service started.' 'OK'
        } else {
            Write-Log 'Windows Update service already running.' 'OK'
        }
    } catch {
        Write-Log "Could not start Windows Update service: $($_.Exception.Message)" 'WARN'
    }
    Write-Log 'No Windows Update packages are downloaded or installed before hardware detection.' 'OK'
}

function Prepare-Resume {
    Write-Log 'STAGE 05 - Resume/reboot preparation (no package download/install)' 'STEP'
    Write-Log 'Main driver and update installation is intentionally deferred until after STAGE 06.' 'OK'
}

function Get-HardwareInventory {
    param([string]$Label='STAGE 06 - Hardware detection')
    Write-Log $Label 'STEP'

    $devices = @(Get-CimInstance Win32_PnPEntity | Where-Object {$_.Present -eq $true})

    # Capture the device's real hardware-ID list once per full inventory. This is the
    # association source used later for every ASUS driver package and Windows Update
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

    # Microsoft Store AppX/MSIX transactions must not be left waiting for UI in an
    # unattended driver installer. Earlier versions combined --interactive with
    # --disable-interactivity; v42 deliberately removes that contradiction.
    # WinGet's msstore source is still used, but the transaction is forced silent,
    # non-interactive, user-scoped, and bounded by a hard timeout.
    $certError = -1978335138 # 0x8A15005E
    $timeoutCode = 1460      # ERROR_TIMEOUT
    $args = @('install','--id',$PackageId)
    if($Exact){ $args += '--exact' }
    $args += @('--source','msstore','--accept-package-agreements','--accept-source-agreements')
    $args += '--silent','--disable-interactivity','--scope','user'
    if($Interactive){
        Write-Log "[$PackageId] An interactive Store install was requested by the caller, but v42 forces it non-interactive to prevent a Store UI/backend stall." 'INFO'
    }

    $run = {
        param($a)
        # Keep WinGet's normal stdout/progress visible, but capture stderr so a transient
        # certificate/source diagnostic from an expected first attempt cannot paint the
        # PowerShell console red when the certificate-pinning retry subsequently succeeds.
        $stderrFile = Join-Path $env:TEMP ("ROG-G635LW-WinGet-Stderr-{0}.log" -f ([guid]::NewGuid().ToString('N')))
        try {
            $p=Start-Process -FilePath $WingetPath -ArgumentList $a -PassThru -NoNewWindow -RedirectStandardError $stderrFile
            if(-not $p.WaitForExit([Math]::Max(1,$TimeoutSeconds)*1000)){
                Write-Log "WinGet Store operation exceeded the $TimeoutSeconds-second timeout; terminating WinGet and its child processes." 'WARN'
                try { & taskkill.exe /PID $p.Id /T /F 2>&1 | Out-Null } catch { try { $p.Kill($true) } catch {} }
                return $timeoutCode
            }
            return [int]$p.ExitCode
        } finally {
            if(Test-Path -LiteralPath $stderrFile){
                try { Remove-Item -LiteralPath $stderrFile -Force -ErrorAction SilentlyContinue } catch {}
            }
        }
    }

    $code=& $run $args
    if($code -ne $certError){
        return [pscustomobject]@{ExitCode=$code;UsedCertificateBypass=$false}
    }

    Write-Log 'WinGet returned 0x8A15005E (server certificate did not match the expected Microsoft Store certificate). Retrying the Store operation once with Microsoft Store certificate-pinning bypass temporarily enabled.' 'WARN'
    Write-Log 'The bypass is temporary and will be restored immediately after this Store operation; it is not left enabled by the installer.' 'INFO'

    $previousEnabled=$false
    try {
        $info=@(& $WingetPath --info 2>&1 | Out-String)
        if(($info -join '') -match '(?im)BypassCertificatePinningForMicrosoftStore\s+Enabled'){$previousEnabled=$true}
    } catch {}

    $enabled=$false
    try {
        $setArgs=@('settings','--enable','BypassCertificatePinningForMicrosoftStore','--disable-interactivity')
        $sp=Start-Process -FilePath $WingetPath -ArgumentList $setArgs -Wait -PassThru -NoNewWindow
        if($sp.ExitCode -eq 0){
            $enabled=$true
            Write-Log 'Temporarily enabled WinGet Microsoft Store certificate-pinning bypass for the retry.' 'INFO'
        } else { Write-Log "WinGet could not enable the temporary Microsoft Store certificate-pinning bypass (exit code $($sp.ExitCode))." 'WARN' }
        if($enabled){ $code=& $run $args }
    } finally {
        if($enabled -and -not $previousEnabled){
            try {
                $restoreArgs=@('settings','--disable','BypassCertificatePinningForMicrosoftStore','--disable-interactivity')
                $rp=Start-Process -FilePath $WingetPath -ArgumentList $restoreArgs -Wait -PassThru -NoNewWindow
                if($rp.ExitCode -eq 0){ Write-Log 'Microsoft Store certificate-pinning validation restored after the temporary retry.' 'OK' }
                else { Write-Log "WARNING: WinGet could not restore Microsoft Store certificate-pinning validation (exit code $($rp.ExitCode))." 'ERROR' }
            } catch { Write-Log "WARNING: Exception while restoring Microsoft Store certificate-pinning validation: $($_.Exception.Message)" 'ERROR' }
        }
    }
    return [pscustomobject]@{ExitCode=$code;UsedCertificateBypass=$enabled}
}

function Install-MyASUS {
    Write-Log 'STAGE 07A - REQUIRED ASUS MyASUS Microsoft Store installation/update (after hardware detection)' 'STEP'
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) {
        Write-Log 'winget.exe was not found. MyASUS is REQUIRED and the ASUS package phase cannot be marked complete without it.' 'ERROR'
        return $false
    }

    $id='9N7R5S6B0ZZH'

    # If the user has Microsoft Store open, close only the Store UI before any Store
    # registration repair or deployment. This avoids ERROR_PACKAGE_IN_USE.
    # This avoids ERROR_PACKAGE_IN_USE for Microsoft.WindowsStore while leaving the rest
    # of the interactive session untouched.
    try {
        $storeProcesses=@(Get-Process -Name 'WinStore.App','MicrosoftStore' -ErrorAction SilentlyContinue)
        if($storeProcesses.Count -gt 0){
            Write-Log 'Microsoft Store UI is open. Closing Store before the MyASUS AppX deployment to prevent package-in-use errors.' 'INFO'
            foreach($sp in $storeProcesses){ try { $sp.CloseMainWindow() | Out-Null } catch {} }
            Start-Sleep -Seconds 2
            $storeProcesses=@(Get-Process -Name 'WinStore.App','MicrosoftStore' -ErrorAction SilentlyContinue)
            foreach($sp in $storeProcesses){ try { Stop-Process -Id $sp.Id -Force -ErrorAction SilentlyContinue } catch {} }
            Start-Sleep -Seconds 2
        }
    } catch {}

    if(-not (Initialize-MicrosoftStoreForCurrentUser)){
        Write-Log 'Microsoft Store is not registered/available for the current interactive user. MyASUS remains REQUIRED; stopping before the ASUS full-package checkpoint.' 'ERROR'
        return $false
    }

    $maxAttempts=3
    for($attempt=1;$attempt -le $maxAttempts;$attempt++){
        try {
            Write-Log "Installing MyASUS from Microsoft Store using exact package ID $id (attempt $attempt/$maxAttempts)." 'INFO'
            Write-Log 'MyASUS WinGet download/install operation starting; WinGet owns the package progress display for this operation.' 'STEP'
            $result=Invoke-WingetMsStoreInstall -WingetPath $winget.Source -PackageId $id -Exact -Silent -TimeoutSeconds 180
            $code=[int]$result.ExitCode
            Write-Log "MyASUS winget exit code: $code"
            if ($code -eq 0) {
                Write-Log 'MyASUS installed successfully from the Microsoft Store.' 'OK'
                return $true
            }
            if ($code -eq -1978335189) {
                Write-Log 'MyASUS is already at the available version; WinGet reported UPDATE_NOT_APPLICABLE.' 'OK'
                return $true
            }
            if ($code -eq 1460) {
                Write-Log 'MyASUS Microsoft Store transaction timed out after 180 seconds.' 'WARN'
            } elseif ($code -eq -1978335138) {
                Write-Log 'MyASUS Microsoft Store installation returned 0x8A15005E after the certificate-pinning retry.' 'WARN'
            } else {
                Write-Log "MyASUS returned exit code $code." 'WARN'
            }
        } catch {
            Write-Log "MyASUS installation attempt $attempt failed: $($_.Exception.Message)" 'WARN'
        }
        if($attempt -lt $maxAttempts){
            Write-Log 'MyASUS is REQUIRED and was not confirmed installed. Waiting 10 seconds before the next Store installation attempt.' 'WARN'
            Start-Sleep -Seconds 10
        }
    }
    Write-Log 'MyASUS could not be installed/confirmed after 3 Microsoft Store attempts. The ASUS package phase will remain incomplete and the installer will resume MyASUS on the next invocation.' 'ERROR'
    return $false
}

function Get-UpdateCategories {
    param($Update)
    $names=@()
    try {
        foreach ($c in $Update.Categories) {
            if ($c.Name) { $names += [string]$c.Name }
        }
    } catch {}
    return ($names -join ', ')
}


function Get-WindowsUpdateDriverMetadata {
    param([Parameter(Mandatory)]$Update)
    $props=@{IsDriver=$false;HardwareId='';Manufacturer='';Model='';Provider='';Version='';Date=$null;Class=''}
    try{$props.IsDriver=([string](Get-UpdateCategories $Update) -match '(?i)Driver|Drivers')}catch{}
    foreach($map in @(
        @('DriverHardwareID','HardwareId'),@('DriverManufacturer','Manufacturer'),@('DriverModel','Model'),@('DriverProvider','Provider'),@('DriverVerVersion','Version'),@('DriverClass','Class')
    )){
        try{$v=$Update.($map[0]);if($null -ne $v){$props[$map[1]]=[string]$v}}catch{}
        if(-not $props[$map[1]]){try{$p=$Update.PSObject.Properties[$map[0]];if($p){$props[$map[1]]=[string]$p.Value}}catch{}}
    }
    try{$v=$Update.DriverVerDate;if($v){$props.Date=[datetime]$v}}catch{try{$p=$Update.PSObject.Properties['DriverVerDate'];if($p -and $p.Value){$props.Date=[datetime]$p.Value}}catch{}}
    return [pscustomobject]$props
}

function Test-WindowsUpdateDriverAlreadySatisfied {
    param([Parameter(Mandatory)]$Update,[Parameter(Mandatory)]$Devices,[Parameter(Mandatory)]$InstalledDrivers)
    $m=Get-WindowsUpdateDriverMetadata -Update $Update
    if(-not $m.IsDriver){return [pscustomobject]@{Satisfied=$false;Reason='Not a driver-category Windows Update';Metadata=$m}}
    $hid=[string]$m.HardwareId
    $title=[string]$Update.Title
    foreach($d in @($InstalledDrivers)){
        $ids=@(Get-DeviceHardwareIds -Device $d)
        $idMatch=$false
        if($hid){foreach($id in $ids){if($id -ieq $hid -or $id -imatch ('^'+[regex]::Escape($hid)+'$')){$idMatch=$true;break}}}
        if(-not $idMatch -and $hid -and ($hid -match '(?i)VEN_[0-9A-F]{4}&DEV_[0-9A-F]{4}')){
            $venDev=$Matches[0];foreach($id in $ids){if($id -match ('(?i)'+[regex]::Escape($venDev))){$idMatch=$true;break}}
        }
        if(-not $hid){
            $idMatch=("$($d.DeviceName) $($d.Manufacturer) $($d.DeviceID) $($d.HardwareID)" -match [regex]::Escape($m.Manufacturer)) -or
                     ("$($d.DeviceName) $($d.DeviceID) $($d.HardwareID)" -match [regex]::Escape(($m.Model -replace '[^A-Za-z0-9 _-]',' ').Trim()))
        }
        if(-not $idMatch){continue}
        $have=Convert-ToComparableVersion ([string]$d.DriverVersion);$want=Convert-ToComparableVersion ([string]$m.Version)
        $haveDate=$null;try{if([string]$d.DriverDate){$haveDate=[datetime]$d.DriverDate}}catch{}
        if($want -and $have -and $have -ge $want){return [pscustomobject]@{Satisfied=$true;Reason="Installed driver '$($d.DeviceName)' version $($d.DriverVersion) is >= Windows Update driver $($m.Version)";Metadata=$m}}
        if($m.Date -and $haveDate -and $haveDate -ge $m.Date){return [pscustomobject]@{Satisfied=$true;Reason="Installed driver '$($d.DeviceName)' date $($haveDate.ToString('yyyy-MM-dd')) is >= Windows Update driver date $($m.Date.ToString('yyyy-MM-dd'))";Metadata=$m}}
    }
    return [pscustomobject]@{Satisfied=$false;Reason='No matching installed hardware driver at or above the Windows Update driver date/version';Metadata=$m}
}

function Install-WindowsUpdatePass {
    param([int]$PassNumber=1,[Parameter(Mandatory)]$Devices,[Parameter(Mandatory)]$InstalledDrivers)
    Write-Log "STAGE 07E - Full Windows/Microsoft Update + official driver pass $PassNumber" 'STEP'
    Enable-MicrosoftUpdateService
    $maxSearchAttempts=3
    for($attempt=1;$attempt -le $maxSearchAttempts;$attempt++){
        try{
            $session=New-Object -ComObject Microsoft.Update.Session
            $session.ClientApplicationID='ROG-G635LW-Installer-v45'
            $searcher=$session.CreateUpdateSearcher();$searcher.IncludePotentiallySupersededUpdates=$true
            $criteria='IsInstalled=0 and IsHidden=0'
            Write-Log "Searching Windows Update with criteria: $criteria (attempt $attempt/$maxSearchAttempts)"
            $result=$searcher.Search($criteria)
            Write-Log "Windows Update returned $($result.Updates.Count) applicable update(s)."
            if($result.Updates.Count -eq 0){Write-Log 'No applicable Windows/Microsoft updates or driver packages were found.' 'OK';return [pscustomobject]@{RebootRequired=$false;InstalledCount=0;DriverCount=0;SearchFailed=$false}}
            $updates=New-Object -ComObject Microsoft.Update.UpdateColl;$driverCount=0
            for($i=0;$i -lt $result.Updates.Count;$i++){
                $u=$result.Updates.Item($i);$categories=Get-UpdateCategories $u;$isDriver=($categories -match '(?i)Driver|Drivers')
                Write-Log ("[{0}/{1}] {2} | Categories: {3}" -f ($i+1),$result.Updates.Count,$u.Title,$categories)
                if($isDriver){
                    $driverCount++
                    $gate=Test-WindowsUpdateDriverAlreadySatisfied -Update $u -Devices $Devices -InstalledDrivers $InstalledDrivers
                    $wm=$gate.Metadata
                    if($wm.HardwareId -or $wm.Version -or $wm.Date){Write-Log "    Driver association: HardwareID='$($wm.HardwareId)' | Manufacturer='$($wm.Manufacturer)' | Model='$($wm.Model)' | Version='$($wm.Version)' | Date='$($wm.Date)'" 'INFO'}
                    if($gate.Satisfied){Write-Log "    [SKIP] Windows Update driver already satisfied: $($gate.Reason)" 'OK';continue}
                    Write-Log "    [QUEUE] Windows Update driver is not yet satisfied by the matching hardware/date/version; it remains eligible." 'INFO'
                }
                try{if(-not $u.EulaAccepted){$u.AcceptEula()};[void]$updates.Add($u)}catch{Write-Log "Could not stage update '$($u.Title)': $($_.Exception.Message)" 'WARN'}
            }
            if($updates.Count -eq 0){Write-Log 'No updates could be staged for installation.' 'WARN';return [pscustomobject]@{RebootRequired=$false;InstalledCount=0;DriverCount=$driverCount;SearchFailed=$false}}
            # WUA stores its payload in the Windows Update cache rather than exposing a supported
            # per-update arbitrary output path. Save a complete manifest here so every update that
            # Windows attempted to download/install is still accounted for in the user's archive.
            try {
                $wuManifest=Join-Path $DownloadRoot ("WindowsUpdate-Pass-{0}-Manifest.csv" -f $PassNumber)
                $manifestRows=@()
                for($mi=0;$mi -lt $updates.Count;$mi++){
                    $mu=$updates.Item($mi)
                    $manifestRows += [pscustomobject]@{Title=$mu.Title;UpdateID=$mu.Identity.UpdateID;Revision=$mu.Identity.RevisionNumber;Categories=(Get-UpdateCategories $mu);IsDownloaded=[bool]$mu.IsDownloaded}
                }
                $manifestRows | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $wuManifest
                Write-Log "Windows Update manifest: $wuManifest" 'OK'
            } catch { Write-Log "Could not write Windows Update manifest: $($_.Exception.Message)" 'WARN' }
            Write-Log "Downloading $($updates.Count) official Microsoft Update package(s), including $driverCount driver update(s) where offered..."
            Write-Log "Microsoft Update API is starting $($updates.Count) package download(s). Per-file HTTP byte progress is not exposed by the Windows Update COM API, so this phase reports the start and completion/result code rather than inventing a percentage or speed." 'INFO';$downloader=$session.CreateUpdateDownloader();$downloader.Updates=$updates;$dl=$downloader.Download();Write-Log "Microsoft Update download phase complete. Result code: $($dl.ResultCode)" 'OK'
            if($dl.ResultCode -notin @(2,3)){Write-Log "The update download operation did not report a successful result. Code: $($dl.ResultCode)" 'WARN'}
            $installer=$session.CreateUpdateInstaller();$installer.Updates=$updates;Write-Log "Installing $($updates.Count) downloaded package(s)...";$ir=$installer.Install();Write-Log "Install result code: $($ir.ResultCode)"
            for($i=0;$i -lt $updates.Count;$i++){try{Write-Log ("Result: {0} => {1}" -f $updates.Item($i).Title,$ir.GetUpdateResult($i).ResultCode)}catch{}}
            return [pscustomobject]@{RebootRequired=[bool]$ir.RebootRequired;InstalledCount=$updates.Count;DriverCount=$driverCount;SearchFailed=$false}
        }catch{
            $msg=$_.Exception.Message;$hr=$null;try{$hr=$_.Exception.HResult}catch{}
            if($msg -match '0x80240438' -or "$hr" -eq '-2145107960'){Write-Log "Windows Update Agent assessment failed with 0x80240438 on attempt $attempt. Retrying the WUA assessment." 'WARN'}else{Write-Log "Windows Update pass $PassNumber assessment failed on attempt ${attempt}: $msg" 'WARN'}
            if($attempt -lt $maxSearchAttempts){try{Restart-Service wuauserv -Force -ErrorAction SilentlyContinue}catch{};try{Start-Service bits -ErrorAction SilentlyContinue}catch{};Start-Sleep -Seconds (5*$attempt)}
        }
    }
    try{$uso=Join-Path $env:SystemRoot 'System32\UsoClient.exe';if(Test-Path $uso){Write-Log 'WUA assessment remained unavailable; starting native USOClient StartScan as a fallback.' 'WARN';Start-Process -FilePath $uso -ArgumentList 'StartScan' -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue}}catch{}
    return [pscustomobject]@{RebootRequired=$false;InstalledCount=0;DriverCount=0;SearchFailed=$true}
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
        $script:DownloadRoot=Join-Path ([Environment]::GetFolderPath('Desktop')) 'ASUS Rog Strix Scar 16 2025 G635LW Drivers and Softwares'
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

function Invoke-FileDownloadWithProgress {
    param([Parameter(Mandatory)][string]$Uri,[Parameter(Mandatory)][string]$Destination,[string]$Label='File',[string]$ExpectedSHA256='')
    $tmp="$Destination.partial";$resp=$null;$stream=$null;$file=$null
    try{
        if((Test-Path $Destination)-and((Get-Item $Destination).Length -gt 0)){
            if($ExpectedSHA256 -match '^[A-Fa-f0-9]{64}$'){$existing=(Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash;if($existing -eq $ExpectedSHA256.ToUpperInvariant()){Write-Log "[$Label] Existing file matches SHA-256; download skipped." 'OK';return $true};Remove-Item $Destination -Force -ErrorAction SilentlyContinue}else{Write-Log "[$Label] Existing file retained because no vendor SHA-256 was exposed." 'OK';return $true}
        }
        if(Test-Path $tmp){Remove-Item $tmp -Force -ErrorAction SilentlyContinue};New-Item -ItemType Directory -Force -Path (Split-Path $Destination -Parent)|Out-Null
        [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
        $req=[Net.HttpWebRequest]::Create($Uri);$req.Method='GET';$req.AllowAutoRedirect=$true;$req.Timeout=60000;$req.ReadWriteTimeout=60000;$req.UserAgent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v45';$req.Accept='*/*'
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

function Invoke-OfficialDownloadAndInstall {
    param([Parameter(Mandatory)][string]$Label,[Parameter(Mandatory)][string]$Uri,[Parameter(Mandatory)][string]$Destination,[string]$SHA256='',[string]$Arguments='/S',[switch]$Install)
    $ok=Invoke-FileDownloadWithProgress -Uri $Uri -Destination $Destination -Label $Label -ExpectedSHA256 $SHA256;if(-not$ok){return [pscustomobject]@{Downloaded=$false;Installed=$false;ExitCode=$null;Path=$Destination}}
    if(-not$Install){return [pscustomobject]@{Downloaded=$true;Installed=$false;ExitCode=$null;Path=$Destination}}
    try{$p=Start-Process -FilePath $Destination -ArgumentList $Arguments -WorkingDirectory (Split-Path $Destination -Parent) -Wait -PassThru -WindowStyle Hidden;$good=($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010 -or $p.ExitCode -eq 1641);if($good){Write-Log "[$Label] Silent installer exit code $($p.ExitCode)." 'OK'}else{Write-Log "[$Label] Installer exit code $($p.ExitCode)." 'WARN'};return [pscustomobject]@{Downloaded=$true;Installed=$good;ExitCode=$p.ExitCode;Path=$Destination}}catch{Write-Log "[$Label] Silent installation failed: $($_.Exception.Message)" 'WARN';return [pscustomobject]@{Downloaded=$true;Installed=$false;ExitCode=$null;Path=$Destination}}
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
    # Intel Graphics Command Center is a Microsoft Store companion application, not a driver.
    if($t -match '(?i)Intel.*Graphics.*Command\s*Center|Graphics\s*Command\s*Center|Command\s*Center\s*Application' -or $t -match '(?i)apps\.microsoft\.com/store/apps/9PLFNLNT3G5G') { return $false }
    return ($t -match '(?i)driver|chipset|graphics|display|audio|sound|bluetooth|wireless|wlan|lan|ethernet|serial.?io|dynamic.?tuning|management.?engine|csme|vpu|gna|neural|smart.?sound|sst|touchpad|camera|card.?reader|storage|rst|rapid.?storage|thunderbolt|usb|firmware.?driver')
}

function Test-PackageIsMicrosoftStoreApp {
    param([Parameter(Mandatory)]$Package)
    $t = "$($Package.Category) $($Package.Title) $($Package.DownloadUrl)"
    return ($t -match '(?i)Intel.*Graphics.*Command\s*Center' -or $t -match '(?i)9PLFNLNT3G5G')
}

function Test-StoreAppInstalled {
    param([Parameter(Mandatory)]$Package)
    $t = "$($Package.Title) $($Package.DownloadUrl)"
    if($t -notmatch '(?i)Intel.*Graphics.*Command\s*Center|9PLFNLNT3G5G') { return $false }
    try {
        $apps=@(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object {
            ([string]$_.Name -match '(?i)IntelGraphicsExperience|AppUp\.IntelGraphicsExperience') -or
            ([string]$_.PackageFullName -match '(?i)IntelGraphicsExperience')
        })
        return ($apps.Count -gt 0)
    } catch { return $false }
}

function Test-CompanionPrerequisitesSatisfied {
    param([Parameter(Mandatory)]$Package,[Parameter(Mandatory)]$InstalledDrivers)
    $t="$($Package.Title) $($Package.Category) $($Package.DownloadUrl)"
    if($t -match '(?i)Intel.*Graphics.*Command\s*Center|9PLFNLNT3G5G') {
        $gfx=@($InstalledDrivers | Where-Object {
            $id=[string]$_.DeviceID
            $name=[string]$_.DeviceName
            $ver=[string]$_.DriverVersion
            $status=[string]$_.Status
            (($id -match '(?i)VEN_8086&DEV_7D67') -or ($name -match '(?i)Intel.*(Graphics|Display|Arc)')) -and
            ($ver -match '\d') -and ($status -notmatch '(?i)error|unknown')
        })
        if($gfx.Count -gt 0) {
            Write-Log "COMPANION GATE: Intel Graphics Command Center is allowed because the Intel graphics driver/device is present in the live installed-driver inventory." 'OK'
            return $true
        }
        Write-Log 'COMPANION GATE: Intel Graphics Command Center is deferred because the Intel graphics driver/device is not yet present. No Store application will be installed before its driver prerequisite.' 'WARN'
        return $false
    }
    return $true
}

function Sort-PackagesForDependencyOrder {
    param([Parameter(Mandatory)][object[]]$Packages)
    $driver=@($Packages | Where-Object { Test-PackageLooksLikeDriver -Package $_ })
    $software=@($Packages | Where-Object { -not (Test-PackageLooksLikeDriver -Package $_) })
    # Within the driver phase, keep the required Intel graphics/platform ordering first.
    $ig=@($driver | Where-Object {
        $t="$($_.Category) $($_.Title)"; $raw=''; try{$raw=$_.Raw|ConvertTo-Json -Depth 20 -Compress}catch{}
        (($t -match '(?i)Intel.*(Graphic|Graphics|Display)') -or ($raw -match '(?i)ArrowLake.?HX|7D67|Intel.*Graphics.*Driver|Intel Graphic driver')) -and $t -notmatch '(?i)Command\s*Center|Application'
    })
    $ip=@($driver | Where-Object {
        $t="$($_.Category) $($_.Title)"; $t -match '(?i)Intel.*(Chipset|Platform Power Management|Serial IO|Dynamic Tuning|Converged Security|Management Engine|VPU|Platform Monitoring|Gaussian|Neural|SST|Smart Sound)|\b(Chipset|Platform Power Management|Intel\(R\) Serial IO|Intel Dynamic Tuning|Intel Converged Security|Intel VPU|Intel Platform Monitoring|Intel Gaussian|Intel\(R\) SST)\b'
    }) | Where-Object { $ig -notcontains $_ }
    # v45: storage and wired LAN are explicit high-priority driver groups. They are
    # removed from the generic bucket and processed immediately after Intel platform
    # prerequisites, so IRST/VMD and Realtek LAN cannot be buried behind later software.
    $irst=@($driver | Where-Object { "$($_.Category) $($_.Title)" -match '(?i)Intel.*Rapid Storage|\bIRST\b|Rapid Storage|\bRST\b|VMD|Volume Management' })
    $lan=@($driver | Where-Object { "$($_.Category) $($_.Title)" -match '(?i)Realtek.*LAN|LAN Driver|RTL8111H|RTL8125D' }) | Where-Object { $irst -notcontains $_ }
    $other=@($driver | Where-Object { $ig -notcontains $_ -and $ip -notcontains $_ -and $irst -notcontains $_ -and $lan -notcontains $_ })
    $ig=@(Sort-PackagesByArchitecture -Packages $ig)
    $ip=@(Sort-PackagesByArchitecture -Packages $ip)
    $irst=@(Sort-PackagesByArchitecture -Packages $irst)
    $lan=@(Sort-PackagesByArchitecture -Packages $lan)
    $other=@(Sort-PackagesByArchitecture -Packages $other)
    $software=@(Sort-PackagesByArchitecture -Packages $software)
    return @($ig+$ip+$irst+$lan+$other+$software)
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
    param([Parameter(Mandatory)][object[]]$Packages)
    # ALL architecture variants are retained and sorted strictly by priority:
    # x64 first, ARM64 second, x86 third, then neutral/unspecified packages.
    return @($Packages | Sort-Object @{Expression={-[int]$_.ArchitectureRank};Descending=$false}, @{Expression={[string]$_.Title};Descending=$false}, @{Expression={ $v=Convert-ToComparableVersion ([string]$_.Version); if($null -eq $v){[version]'0.0'}else{$v} };Descending=$true})
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
                $r=Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 60 -Headers @{'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v45';'Accept'='application/json,text/plain,*/*';'Referer'='https://rog.asus.com/support/'}
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
    $packages=@(Sort-PackagesByArchitecture -Packages $packages)
    Write-Log "ASUS catalog parser committed $($packages.Count) package objects after architecture filtering/priority." 'INFO'
    if($packages.Count -eq 0){Write-Log 'ASUS catalog contained no current Windows package installer URLs after filtering.' 'WARN';return @()}
    Write-Log "ASUS catalog yielded $($packages.Count) DISTINCT full-package candidates after excluding BIOS/firmware/documentation. No title/category deduplication was applied." 'OK'
    $realtekCatalog=@($packages | Where-Object { "$($_.Category) $($_.Title)" -match '(?i)Realtek.*LAN|LAN Driver|RTL8111H|RTL8125D' })
    $irstCatalog=@($packages | Where-Object { "$($_.Category) $($_.Title)" -match '(?i)Intel.*Rapid Storage|\bIRST\b|Rapid Storage|\bRST\b|VMD|Volume Management' })
    $auditLevel=if($realtekCatalog.Count -gt 0 -and $irstCatalog.Count -gt 0){'OK'}else{'WARN'}
    Write-Log "CRITICAL CATALOG AUDIT: Realtek LAN candidates=$($realtekCatalog.Count); Intel IRST/VMD candidates=$($irstCatalog.Count). These packages are explicitly retained and separately prioritized for installation." $auditLevel
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
        'VPU' {$patterns+=@('VPU|Vision Processing');break}
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

function Request-BatchedReboot {
    param([Parameter(Mandatory)]$BatchState)

    Write-Log "Six attended ASUS packages have been processed since the last full reboot while a reboot is pending. The package files and checkpoints are saved." 'WARN'

    # A scheduled -Resume process runs at Windows startup under SYSTEM, so it cannot
    # reliably receive interactive keyboard input. If that automatic resume reaches a
    # batched-reboot checkpoint, continue automatically: save the checkpoint, reboot,
    # and let the startup task resume again. Manual runs always receive an explicit Y/N
    # choice below.
    if($Resume){
        $BatchState.PendingReboot=$true
        Save-AsusPackageBatchState -State $BatchState
        Write-Log 'Automatic -Resume run reached a batched reboot checkpoint. No interactive prompt is possible under the startup task, so the saved checkpoint will be followed by an automatic reboot.' 'WARN'
        return $true
    }

    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Yellow
    Write-Host ' A SYSTEM REBOOT IS NOW RECOMMENDED' -ForegroundColor Yellow
    Write-Host " $($BatchState.PackagesSinceReboot) packages have been installed since the last reboot." -ForegroundColor Yellow
    Write-Host '============================================================' -ForegroundColor Yellow
    Write-Host ' Y = reboot Windows now' -ForegroundColor Green
    Write-Host ' N = do not reboot now; save the checkpoint and resume automatically after you reboot later' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor Yellow

    do {
        $answer=Read-Host 'Restart Windows now? [Y/N]'
        if($answer -match '^(?i:y|yes)$'){
            $BatchState.PendingReboot=$true
            Save-AsusPackageBatchState -State $BatchState
            Write-Log 'User selected Y: the batched reboot checkpoint is saved and Windows will reboot.' 'WARN'
            return $true
        }
        if($answer -match '^(?i:n|no)$'){
            $BatchState.PendingReboot=$true
            Save-AsusPackageBatchState -State $BatchState
            # Keep the startup resume task registered. The user may reboot manually later;
            # Windows startup will then invoke this exact v33 checkpoint automatically.
            Register-ResumeTask
            Write-Log 'User selected N: no reboot was initiated. The checkpoint and startup resume task were preserved so the installer will automatically resume after the next Windows reboot.' 'WARN'
            return $false
        }
        Write-Host 'Please enter Y or N.' -ForegroundColor Red
    } while($true)
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
            $url=[string]$Package.DownloadUrl

            $identity=($title+' '+$category+' '+$url)
            $identity=$identity -replace '(?i)\b(v?\d+(?:\.\d+){1,4})\b',' '
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

        # Pick the highest comparable version. If versions tie or are unavailable,
        # use release date. No Sort-Object custom comparer is used.
        $winner=$candidates[0]
        $winnerVersion=$null
        try{$winnerVersion=Convert-ToComparableVersion ([string]$winner.Version)}catch{}
        $winnerDate=$null
        try{if([string]$winner.ReleaseDate){$winnerDate=[datetime]$winner.ReleaseDate}}catch{}

        if($candidates.Count -gt 1){
            for($i=1;$i -lt $candidates.Count;$i++){
                $candidate=$candidates[$i]
                $cv=$null
                try{$cv=Convert-ToComparableVersion ([string]$candidate.Version)}catch{}
                $cd=$null
                try{if([string]$candidate.ReleaseDate){$cd=[datetime]$candidate.ReleaseDate}}catch{}

                $take=$false
                if($cv -and -not $winnerVersion){$take=$true}
                elseif($cv -and $winnerVersion -and $cv -gt $winnerVersion){$take=$true}
                elseif(-not $cv -and -not $winnerVersion -and $cd -and -not $winnerDate){$take=$true}
                elseif($cd -and $winnerDate -and $cd -gt $winnerDate){$take=$true}

                if($take){
                    $winner=$candidate
                    $winnerVersion=$cv
                    $winnerDate=$cd
                }
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

        if(-not ($hardwareMatched -or $familyMatched -or $tokenMatched)){continue}

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
        foreach($d in @($InstalledDrivers)){
            $dev = [pscustomobject]@{
                Name = [string]$d.DeviceName
                Manufacturer = [string]$d.Manufacturer
                PNPDeviceID = [string]$d.DeviceID
                HardwareID = [string]$d.HardwareID
            }
            $df = Get-DeviceComponentFamily -Device $dev
            if(-not $family -or -not $df -or $df -ne $family){ continue }

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


function Install-RelatedMicrosoftStoreApps {
    param(
        [Parameter(Mandatory=$true)]$Devices,
        [Parameter(Mandatory=$true)]$AsusPackages
    )

    Write-Log 'STAGE 07G - Installing related Microsoft Store hardware/driver companion applications through the official msstore source' 'STEP'
    Write-Log 'The Store companion pass is additive: it does not replace ASUS executable packages, and Store failures never suppress the official driver archive or Windows Update fallback.' 'INFO'

    $apps=@(
        [pscustomobject]@{Name='MyASUS';Id='9N7R5S6B0ZZH';Publisher='ASUS';Reason='ASUS System Control Interface / full MyASUS experience';Patterns='ASUS System Control Interface|MyASUS';Required=$true},
        [pscustomobject]@{Name='Intel Graphics Command Center';Id='9PLFNLNT3G5G';Publisher='Intel';Reason='Intel graphics control companion';Patterns='Intel.*Graphics|Graphics.*Command.*Center|Intel Graphics Command Center';Required=$false},
        [pscustomobject]@{Name='Realtek Audio Control';Id='9P2B8MCSVPLN';Publisher='Realtek Semiconductor Corp.';Reason='Realtek Audio/Codec Console companion';Patterns='Realtek.*Audio|Realtek.*Codec|ALC3288|Audio';Required=$false},
        [pscustomobject]@{Name='NVIDIA Control Panel';Id='9NF8H0H7WMLT';Publisher='NVIDIA Corp.';Reason='NVIDIA DCH graphics control companion';Patterns='NVIDIA.*(Graphic|Graphics|Display)|GeForce|RTX 5080';Required=$false},
        [pscustomobject]@{Name='Dolby Access';Id='9N0866FS04W8';Publisher='Dolby Laboratories';Reason='Dolby Atmos companion for the ASUS Dolby Atmos package';Patterns='Dolby Atmos|Dolby';Required=$false},
        [pscustomobject]@{Name='Armoury Crate';Id='9PM9DFQRDH3F';Publisher='ASUS';Reason='ROG/ASUS device-control companion';Patterns='Armoury Crate|ArmouryCrate|Aura Creator';Required=$false},
        [pscustomobject]@{Name='Aura Creator';Id='9MWZ3VLW5HBW';Publisher='ASUS';Reason='Aura/lighting companion associated with Armoury Crate';Patterns='Armoury Crate|Aura Creator|Aura|Lighting|RGB';Required=$false}
    )

    # Only install apps whose relationship is represented by the detected hardware or
    # the ASUS package catalog. MyASUS is always applicable to this ASUS notebook.
    $hardwareText=''
    try { $hardwareText=(@($Devices)|ForEach-Object { "$($_.Name) $($_.Manufacturer) $($_.PNPDeviceID) $($_.HardwareID)" }) -join ' ' } catch {}
    $catalogText=''
    try { $catalogText=(@($AsusPackages)|ForEach-Object { "$($_.Title) $($_.Category) $($_.DownloadUrl)" }) -join ' ' } catch {}
    $allText="$hardwareText $catalogText"

    $winget=Get-Command winget.exe -ErrorAction SilentlyContinue
    if(-not $winget){
        Write-Log 'Related Microsoft Store pass: winget.exe is unavailable; all Store companion apps are retained as deferred Store targets for the next invocation.' 'WARN'
        return [pscustomobject]@{Attempted=0;Installed=0;Deferred=$apps.Count}
    }
    if(-not (Initialize-MicrosoftStoreForCurrentUser)){
        Write-Log 'Related Microsoft Store pass: Microsoft Store is not ready for the interactive user. All Store companion apps are deferred without blocking drivers.' 'WARN'
        return [pscustomobject]@{Attempted=0;Installed=0;Deferred=$apps.Count}
    }

    $attempted=0;$installed=0;$deferred=0
    foreach($app in $apps){
        $appApplicable=($app.Required -or ($allText -match $app.Patterns))
        if(-not $appApplicable){
            Write-Log "[MSSTORE] Skipping unrelated Store app '$($app.Name)' because no matching hardware/ASUS package relationship was found in the STAGE 06 baseline/catalog." 'INFO'
            continue
        }

        # AppX detection is intentionally broad so package-family/name changes do not
        # cause a reinstall merely because Microsoft changed the Store package identity.
        $already=$false
        try {
            $needle=($app.Name -replace '[^A-Za-z0-9]','')
            $already=@(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object {
                $blob="$($_.Name) $($_.PackageFullName) $($_.DisplayName)"
                ($blob -replace '[^A-Za-z0-9]','') -match [regex]::Escape($needle)
            }).Count -gt 0
        } catch {}
        if($already){
            Write-Log "[MSSTORE] $($app.Name) is already installed; no Store transaction required." 'OK'
            continue
        }

        $attempted++
        Write-Log "[MSSTORE] Related app selected: $($app.Name) | Product ID $($app.Id) | Publisher $($app.Publisher) | Reason: $($app.Reason)" 'STEP'
        Write-Log "[MSSTORE] Installing '$($app.Name)' with WinGet using the official msstore source and exact Product ID $($app.Id)." 'INFO'
        try {
            $wp=Invoke-WingetMsStoreInstall -WingetPath $winget.Source -PackageId $app.Id -Exact -Silent -TimeoutSeconds 180
            $code=[int]$wp.ExitCode
            if($code -eq 0 -or $code -eq 3010){
                $installed++
                Write-Log "[MSSTORE] $($app.Name) installation completed with exit code $code." 'OK'
            } elseif($code -eq 1460){
                $deferred++
                Write-Log "[MSSTORE] $($app.Name) Store transaction timed out after 180 seconds. It is deferred and will be retried on a later invocation; this is not a driver failure." 'WARN'
            } else {
                $deferred++
                Write-Log "[MSSTORE] $($app.Name) Store transaction returned exit code $code. It is deferred for retry; official driver/package processing continues." 'WARN'
            }
        } catch {
            $deferred++
            Write-Log "[MSSTORE] $($app.Name) Store transaction raised an exception: $($_.Exception.Message). Deferred for retry." 'WARN'
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
            # Only run executables that look like a same-package configuration/control/utility component.
            return ($n -match '(?i)(control|console|config|configuration|settings?|utility|panel|assistant|manager|tuning|audio|sound|graphics|display|camera|touchpad|keyboard|numberpad|numpad|wireless|bluetooth|lighting|aura|armoury|dolby|realtek|intel|nvidia|logitech|mediatek|cmedia|yamaha|synapse|icue|steelseries)')
        } | Sort-Object FullName)
    if($candidates.Count -eq 0){ return 0 }
    Write-Log "  COMPANION PASS: '$PackageTitle' contains $($candidates.Count) selected same-package configuration/control/utility executable(s). They will be run sequentially before the next package." 'STEP'
    $ran=0
    foreach($exe in $candidates){
        try{
            Write-Log "    Companion executable: $($exe.FullName)" 'INFO'
            $proc=Start-Process -FilePath $exe.FullName -WorkingDirectory $exe.DirectoryName -PassThru -WindowStyle Normal -ErrorAction Stop
            # Configuration utilities are allowed to be GUI processes. Wait up to 120 seconds so a
            # utility that remains open does not permanently block the driver pipeline.
            $finished=$proc.WaitForExit(120000)
            if($finished){Write-Log "    Companion '$($exe.Name)' exited with code $($proc.ExitCode)." $(if($proc.ExitCode -in @(0,3010,1641)){'OK'}else{'WARN'})}
            else{Write-Log "    Companion '$($exe.Name)' remained running after 120 seconds; continuing to the next companion without killing the user's utility." 'WARN'}
            $ran++
        }catch{Write-Log "    Companion '$($exe.Name)' could not be launched: $($_.Exception.Message)" 'WARN'}
    }
    return $ran
}

function Install-AsusFullPackages {
    param(
        [Parameter(Mandatory)]$Devices,
        [Parameter(Mandatory)]$Packages,
        [Parameter(Mandatory)]$InstalledDrivers,
        [Parameter(Mandatory)]$InstalledSoftware
    )
    $script:CurrentLiveDevices=@($Devices)
    Write-Log 'STAGE 07B - Downloading/installing ALL current ASUS G635LW Windows driver/software packages with persistent per-package resume checkpoints' 'STEP'
    $packages=@($Packages)
    if($packages.Count -eq 0){Write-Log 'No ASUS full setup packages were available from the official catalog. The catalog request may have failed due to temporary DNS/network availability; the persistent resume checkpoint remains unchanged.' 'WARN';return [pscustomobject]@{Installed=0;Downloaded=0;RebootRequired=$false;StoppedForInstallerFailure=$false}}
    $legacyInstalledTitles=@(Get-LegacyInstalledAsusTitles)
    if($legacyInstalledTitles.Count -gt 0){Write-Log "Resume migration: found $($legacyInstalledTitles.Count) ASUS package(s) previously completed by v10; those packages will be marked installed and skipped." 'OK'}

    # Dependency-aware serial order: ALL device-driver packages are processed before
    # companion applications/utilities. Within the driver phase the Intel graphics
    # driver is first, followed by Intel platform prerequisites, then other drivers.
    # Companion software such as Intel Graphics Command Center is deliberately placed
    # after the driver phase and is individually gated against its live driver prerequisite.
    $packages=@(Sort-PackagesForDependencyOrder -Packages $packages)
    Write-Log 'Dependency order: driver packages first (Intel Graphics -> Intel platform -> remaining drivers), followed by companion software/utilities. A companion package cannot run until its required device driver is detected as installed.' 'INFO'

    # v41: RETAIN THE COMPLETE ASUS CATALOG. Do not collapse revisions, alternate
    # hardware variants, or companion entries into one logical candidate. The user
    # requested a complete official-package archive, so every distinct ASUS catalog
    # package survives into the download bundle. Hardware applicability is evaluated
    # later at installation time; it must never prevent an official package from being
    # archived.
    Write-Log "v42 lossless ASUS catalog policy: retaining all $($packages.Count) distinct official package candidates. No title/category/logical-target deduplication is applied." 'OK'

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
    Write-Log "PRE-DOWNLOAD: building the complete official ASUS G635LW package bundle before any installer is launched. $($packages.Count) catalog candidates will be processed; Microsoft Store entries remain dynamic and are installed through the official Store source." 'STEP'
    $bundleDownloadFailures=0
    $bundleIndex=0
    foreach($bundlePkg in @($packages)){
        $bundleIndex++
        if(Test-PackageIsMicrosoftStoreApp -Package $bundlePkg){
            Write-Log "PRE-DOWNLOAD [$bundleIndex/$($packages.Count)] $($bundlePkg.Title): Microsoft Store package; no static binary is available to archive." 'INFO'
            continue
        }
        try{
            $bundleDest=Get-PackageDestination -Package $bundlePkg -AsusDirectory $asusDir
            $bundleExisting=Find-ExistingAsusPackage -Package $bundlePkg -AsusDirectory $asusDir -PreferredPath $bundleDest
            if($bundleExisting){$bundleDest=$bundleExisting}
            $bundlePkg.LocalPath=$bundleDest
            $bundleKey=Get-PackageKey -Package $bundlePkg
            $bundleRow=@($resumeRows|Where-Object{$_.PackageKey -eq $bundleKey})|Select-Object -First 1
            if($null -ne $bundleRow){$bundleRow=Normalize-PackageResumeRow -Row $bundleRow}
            if(-not $bundleRow){$bundleRow=[pscustomobject]@{PackageKey=$bundleKey;Title=[string]$bundlePkg.Title;Version=[string]$bundlePkg.Version;DownloadUrl=[string]$bundlePkg.DownloadUrl;LocalPath=$bundleDest;DownloadStatus='Pending';InstallStatus='Pending';LastAction='';Updated=(Get-Date).ToString('o')}}
            $bundleReady=$false
            if(Test-Path $bundleDest){
                try{
                    $bundleLen=(Get-Item $bundleDest).Length
                    if($bundleLen -gt 0){
                        if([string]$bundlePkg.SHA256 -match '^[A-Fa-f0-9]{64}$'){
                            $bundleActual=(Get-FileHash -LiteralPath $bundleDest -Algorithm SHA256).Hash
                            if($bundleActual -eq ([string]$bundlePkg.SHA256).ToUpperInvariant()){$bundleReady=$true;Write-Log "PRE-DOWNLOAD [$bundleIndex/$($packages.Count)] existing SHA-256-valid archive retained: $($bundlePkg.Title)" 'OK'}
                            else{Write-Log "PRE-DOWNLOAD [$bundleIndex/$($packages.Count)] stale archive detected for '$($bundlePkg.Title)'; downloading current ASUS revision to its version/SHA-specific path." 'INFO';Remove-Item $bundleDest -Force -ErrorAction SilentlyContinue}
                        }else{$bundleReady=$true;Write-Log "PRE-DOWNLOAD [$bundleIndex/$($packages.Count)] existing archive retained (ASUS did not expose SHA-256): $($bundlePkg.Title)" 'OK'}
                    }
                }catch{Write-Log "PRE-DOWNLOAD archive check failed for '$($bundlePkg.Title)': $($_.Exception.Message)" 'WARN'}
            }
            if(-not $bundleReady){
                $bundleRow.DownloadStatus='Downloading';$bundleRow.LastAction='Complete catalog pre-download started';$bundleRow.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $bundleKey})+$bundleRow;Save-PackageResumeState -Rows $resumeRows
                if(Invoke-FileDownloadWithProgress -Uri $bundlePkg.DownloadUrl -Destination $bundleDest -Label("ASUS BUNDLE $bundleIndex/$($packages.Count) - $($bundlePkg.Title)") -ExpectedSHA256([string]$bundlePkg.SHA256)){
                    $bundleRow.DownloadStatus='Downloaded';$bundleRow.LastAction='Complete catalog pre-download finished';$bundleRow.LocalPath=$bundleDest;$bundleRow.Updated=(Get-Date).ToString('o');Write-Log "PRE-DOWNLOAD [$bundleIndex/$($packages.Count)] complete: $($bundlePkg.Title)" 'OK'
                }else{
                    $bundleDownloadFailures++;$bundleRow.DownloadStatus='Failed';$bundleRow.LastAction='Complete catalog pre-download failed';$bundleRow.LocalPath=$bundleDest;$bundleRow.Updated=(Get-Date).ToString('o');Write-Log "PRE-DOWNLOAD [$bundleIndex/$($packages.Count)] FAILED: $($bundlePkg.Title). The installation phase will retry this official URL." 'WARN'
                }
                $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $bundleKey})+$bundleRow;Save-PackageResumeState -Rows $resumeRows
            }
        }catch{
            $bundleDownloadFailures++
            Write-Log "PRE-DOWNLOAD [$bundleIndex/$($packages.Count)] exception for '$($bundlePkg.Title)': $($_.Exception.Message). The installation phase will retry it." 'WARN'
        }
    }
    if($bundleDownloadFailures -eq 0){Write-Log "PRE-DOWNLOAD COMPLETE: all static ASUS catalog packages are now archived before installation. Store packages remain dynamic." 'OK'}else{Write-Log "PRE-DOWNLOAD COMPLETE WITH $bundleDownloadFailures failure(s): failed packages remain checkpointed and will be retried by the installation phase." 'WARN'}

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
            Write-Log "  DRIVER ASSOCIATION: Family='$($assoc.Family)' | Applicable=$($assoc.Applicable) | $($assoc.Reason)" 'INFO'
        }
        # v41: hardware association is an INSTALL gate, never a DOWNLOAD gate. Every
        # official ASUS catalog package is archived first so the complete 57-entry
        # bundle is available even when a package targets an alternate hardware variant.

        # SECOND HARD GATE: reconcile the actual installed component immediately before
        # any download/installer launch. This catches packages that became installed
        # between the initial inventory and this package boundary.
        $installedGate=Test-InstalledComponentSatisfiesAsusPackage -Package $pkg -InstalledDrivers $InstalledDrivers -InstalledSoftware $InstalledSoftware
        if($installedGate.Satisfied){
            $row.DownloadStatus=if(Test-Path $dest){'Downloaded'}else{'NotRequired'}
            $row.InstallStatus='Installed(Detected)'
            $row.LastAction=$installedGate.Reason
            $row.LocalPath=$dest
            $row.Updated=(Get-Date).ToString('o')
            $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row
            Save-PackageResumeState -Rows $resumeRows
            Write-Log "  HARD ANTI-REPEAT GATE: '$($pkg.Title)' is already installed at a sufficient version; NO download or installer launch. $($installedGate.Reason)" 'OK'
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
            Write-Log "  Resume directory/state check: this package is already recorded as installed; skipping it and continuing with the next package." 'OK'
            continue
        }

        # v14 authoritative post-reboot reconciliation: if the package was installed
        # manually, by Windows Update, or in a different order, the live driver/software
        # inventory wins over the historical package counter.
        $script:CurrentLiveDevices=@($Devices)
        $live=Test-AsusPackageAlreadySatisfied -Package $pkg -InstalledDrivers $InstalledDrivers -InstalledSoftware $InstalledSoftware
        if($live.Satisfied){
            $row.DownloadStatus=if(Test-Path $dest){'Downloaded'}else{'NotRequired'}
            $row.InstallStatus='Installed(Detected)'
            $row.LastAction=$live.Reason
            $row.LocalPath=$dest
            $row.Updated=(Get-Date).ToString('o')
            $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row
            Save-PackageResumeState -Rows $resumeRows
            Write-Log "  LIVE RECONCILIATION: package already satisfied by installed system component: $($live.Reason). No installer launched." 'OK'
            continue
        }
        $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row
        Save-PackageResumeState -Rows $resumeRows

        # Microsoft Store packages are not stable downloadable files. In particular,
        # the Intel Graphics Command Center is distributed through the Microsoft Store
        # rather than as part of the Intel DCH graphics-driver package. Do not feed its
        # dynamic Store URL to the binary downloader or compare it with a stale static
        # catalog hash. Install it through the official msstore source only after the
        # corresponding Intel graphics driver has been detected.
        if(Test-PackageIsMicrosoftStoreApp -Package $pkg){
            Write-Log ("ASUS package {0}/{1}: {2} {3} [MICROSOFT STORE COMPANION]"-f $index,$packages.Count,$pkg.Title,$pkg.Version) 'INFO'
            if(Test-StoreAppInstalled -Package $pkg){
                $row.DownloadStatus='NotRequired';$row.InstallStatus='Installed(Detected)';$row.LastAction='Microsoft Store app already installed';$row.Updated=(Get-Date).ToString('o')
                $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                Write-Log "[$($pkg.Title)] Microsoft Store application is already installed; no Store download required." 'OK'
                continue
            }
            if(-not (Test-CompanionPrerequisitesSatisfied -Package $pkg -InstalledDrivers $InstalledDrivers)){
                $row.DownloadStatus='Deferred';$row.InstallStatus='DeferredPrerequisite';$row.LastAction='Companion deferred until driver prerequisite is installed';$row.Updated=(Get-Date).ToString('o')
                $resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                Write-Log "[$($pkg.Title)] STOPPING before companion software because its device driver prerequisite is not installed. This package remains the resume point; no companion software is launched ahead of its driver." 'WARN'
                $stopped=$true;break
            }
            $winget=Get-Command winget.exe -ErrorAction SilentlyContinue
            if(-not $winget){
                $deferredStore=$true
                $row.InstallStatus='Deferred(StoreInstallerUnavailable)';$row.LastAction='WinGet/App Installer unavailable; Store companion retained for next invocation';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                Write-Log "[$($pkg.Title)] WinGet/App Installer is unavailable. Deferring this Microsoft Store companion so it cannot block the remaining ASUS driver/software packages." 'WARN'
                continue
            }
            if(-not (Initialize-MicrosoftStoreForCurrentUser)){
                $deferredStore=$true
                $row.InstallStatus='Deferred(StoreNotReady)';$row.LastAction='Microsoft Store not registered/available for current user; retained for next invocation';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                Write-Log "[$($pkg.Title)] Microsoft Store is not ready for the current interactive user. Deferring the Store companion and continuing with the remaining ASUS packages." 'WARN'
                continue
            }
            $row.InstallStatus='Installing';$row.LastAction='Microsoft Store installation launched';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
            # ASUS currently exposes the Intel GCC Store product as 9PLFNLNT3G5G.
            # WinGet supports the msstore source; v42 deliberately uses silent/non-interactive mode.
            try{
                Write-Log "[$($pkg.Title)] Installing from the official Microsoft Store source via WinGet. The ASUS catalog SHA-256 is intentionally NOT used because Store delivery is dynamic." 'STEP'
                $wp=Invoke-WingetMsStoreInstall -WingetPath $winget.Source -PackageId '9PLFNLNT3G5G' -Silent -TimeoutSeconds 180
                $wcode=[int]$wp.ExitCode
                if($wcode -eq 0 -or $wcode -eq 3010){
                    # v33: do not rebuild installed-state inventories after installation; STAGE 06 baseline is authoritative.
                    $installLedger=Add-PackageInstallLedgerEntry -Ledger $installLedger -Package $pkg -Fingerprint $fingerprint -ExitCode $wcode
                    Save-PackageInstallLedger -Rows $installLedger
                    $row.InstallStatus="Installed($wcode)";$row.LastAction='Microsoft Store installation completed';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                    Write-Log "[$($pkg.Title)] Microsoft Store installation completed with exit code $wcode." 'OK'
                    if($wcode -eq 3010){$batchState.PendingReboot=$true;$batchState.PackagesSinceReboot++;Save-AsusPackageBatchState -State $batchState}
                    continue
                }
                $deferredStore=$true
                if($wcode -eq 1460){
                    $row.InstallStatus='Deferred(StoreTimeout)';$row.LastAction='Microsoft Store transaction timed out; package retained for next invocation';
                    Write-Log "[$($pkg.Title)] Microsoft Store transaction timed out after 180 seconds. This is a recoverable Store condition, NOT a driver/package installer failure. Continuing with the remaining ASUS packages." 'WARN'
                } else {
                    $row.InstallStatus="Deferred(StoreExit:$wcode)";$row.LastAction='Microsoft Store installation deferred; package retained for next invocation';
                    Write-Log "[$($pkg.Title)] Microsoft Store installation returned exit code $wcode. This Store companion is being deferred so it cannot block the remaining ASUS driver/software packages." 'WARN'
                }
                $row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                continue
            }catch{
                $row.InstallStatus='InstallException';$row.LastAction='Microsoft Store installation exception';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
                Write-Log "[$($pkg.Title)] Microsoft Store installation failed: $($_.Exception.Message)" 'WARN';$stopped=$true;break
            }
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

        # Checkpoint BEFORE launching because an installer may reboot the machine before PowerShell
        # gets control back. On the next run, this exact package is reconsidered first.
        $row.InstallStatus='Installing';$row.LastAction='Installer launched';$row.Updated=(Get-Date).ToString('o');$resumeRows=@($resumeRows|Where-Object{$_.PackageKey -ne $key})+$row;Save-PackageResumeState -Rows $resumeRows
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
                    Write-Log "  COMPANION PASS complete for '$($pkg.Title)': $companionCount executable(s) processed before the next package." 'OK'
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
                    $batchState.PackagesSinceReboot++
                    Write-Log "Full ASUS setup completed for $($pkg.Title) (exit 1641). The installer initiated a reboot; no subsequent package will be launched." 'WARN'
                } elseif($code -eq 3010 -or $pendingAfter -or $batchState.PendingReboot){
                    # 3010 explicitly means successful installation with reboot deferred/required.
                    # This is the key v14 behavior: continue to the next attended package instead of forcing a reboot.
                    if(-not $batchState.PendingReboot){$batchState.PendingReboot=$true;$batchState.PackagesSinceReboot=0}
                    $batchState.PackagesSinceReboot++
                    Save-AsusPackageBatchState -State $batchState
                    Write-Log "Full ASUS setup completed for $($pkg.Title) (exit $code). Reboot is deferred; continuing to the next package ($($batchState.PackagesSinceReboot)/$MaxPackagesBetweenReboots before a batched reboot prompt)." 'OK'
                    if($batchState.PackagesSinceReboot -ge $MaxPackagesBetweenReboots){
                        if(Request-BatchedReboot -BatchState $batchState){$reboot=$true}else{$deferredReboot=$true;$stopped=$true}
                    }
                } else {
                    Write-Log "Full ASUS setup completed for $($pkg.Title) (exit $code). No reboot is currently pending; continuing." 'OK'
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
                        $batchState.PackagesSinceReboot++
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
            'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v45'
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
        'Intel.*(Dynamic Tuning|DTT)' { return 'Intel Dynamic Tuning' }
        'Intel.*(Smart Sound|SST)' { return 'Intel Smart Sound' }
        'Intel.*(Platform Monitoring|PMT)' { return 'Intel Platform Monitoring' }
        'Intel.*VPU|Vision Processing|VPU' { return 'Intel VPU' }
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
            # Do not guess at a third-party VCOM package. Windows Update and the phone's own
            # MTP/ADB interface remain the authoritative fallback.
        }
        'USB Audio' {
            # USB Audio Class devices commonly use Microsoft's in-box USB Audio driver. No
            # generic vendor driver is forced because doing so can replace device-specific DSP.
        }
        'USB Yamaha Audio' {
            # Yamaha USB audio interfaces vary by model/region; without a model-specific official
            # package URL, retain the device and allow Windows Update to supply the matching driver.
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
            'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v45'
            'Accept'='text/html,application/xhtml+xml,*/*'
        }
        $html=[string]$r.Content
        $hrefs=[regex]::Matches($html,'(?i)href\s*=\s*["'']([^"'']+)["'']')|ForEach-Object{$_.Groups[1].Value}
        $rawLinks=New-Object System.Collections.Generic.List[string]
        foreach($h in @($hrefs)){
            if([string]$h){[void]$rawLinks.Add($h)}
        }
        foreach($m in [regex]::Matches($html,'(?i)https?://[^"''\s<>]+\.(?:exe|msi|zip)(?:\?[^"''\s<>]*)?')){[void]$rawLinks.Add($m.Value)}
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
            'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v45'
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
        if(-not $family){Write-Log "[$vendor] No manufacturer-store family mapping for '$($device.Name)' ($($device.PNPDeviceID)); Windows Update remains the final driver source." 'INFO';continue}

        $asusExact=Test-AsusDeviceCovered -Device $device -AsusPackages $AsusPackages
        $asusFamily=Test-AsusCatalogFamilyCovered -Device $device -AsusPackages $AsusPackages
        if($asusExact -or $asusFamily){
            Write-Log "[$vendor/$family] ASUS catalog coverage exists for '$($device.Name)'; manufacturer generic driver store will NOT replace the ASUS OEM package." 'INFO'
            continue
        }

        $profiles=@(Get-ManufacturerStoreProfiles -Device $device)
        if($profiles.Count -eq 0){
            Write-Log "[$vendor/$family] No direct manufacturer driver-store profile is available for '$($device.Name)'; Windows Update remains the final fallback." 'INFO'
            continue
        }

        foreach($profile in $profiles){
            $key="$($profile.Vendor)|$($profile.Family)|$($profile.PageUrl)"
            if($seen.ContainsKey($key)){continue};$seen[$key]=$true
            Write-Log "[$($profile.Vendor)/$($profile.Family)] Manufacturer driver store selected for '$($device.Name)': $($profile.PageUrl)" 'STEP'
            $meta=Get-PagePackageMetadata -PageUrl $profile.PageUrl
            $links=@(Get-OfficialPageCandidateLinks -PageUrl $profile.PageUrl -LinkPatterns @($profile.LinkPatterns))
            if($links.Count -eq 0){
                Write-Log "[$($profile.Vendor)/$($profile.Family)] No direct x64 installer link was exposed by the official manufacturer page. No third-party mirror will be used." 'WARN'
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
            $gate=Test-ManufacturerPackageAlreadySatisfied -Device $device -Candidate $candidate -InstalledDrivers $InstalledDrivers
            Write-Log "  MANUFACTURER ASSOCIATION: Device='$($device.Name)' | PNP='$($device.PNPDeviceID)' | Family='$family' | Candidate='$name' | Version='$($candidate.Version)' | Date='$($candidate.ReleaseDate)' | AlreadySatisfied=$($gate.Satisfied)" 'INFO'
            if($gate.Satisfied){
                Write-Log "  [SKIP] Manufacturer package already satisfied: $($gate.Reason)" 'OK'
                continue
            }

            $ok=Invoke-FileDownloadWithProgress -Uri $uri -Destination $dest -Label "$($profile.Vendor) $($profile.Family) official driver-store package" -ExpectedSHA256 ''
            if(-not $ok){continue}
            Write-Log "  Manufacturer package persisted before installation: $dest" 'OK'

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
                Write-Log "[$vendor/$family] Using Intel's direct official download endpoint: $uri" 'OK'
            }else{
                $links=@(Get-OfficialPageCandidateLinks -PageUrl $job.Page -LinkPatterns @($job.LinkPatterns))
                if($links.Count -eq 0){
                    Write-Log "[$vendor/$family] No direct official installer link was exposed for '$($job.Label)'; utility remains available from the official vendor page and is not sourced from a third party." 'WARN'
                    continue
                }
                $uri=$links|Select-Object -First 1
            }
            $name=[IO.Path]::GetFileName(([uri]$uri).AbsolutePath)
            if(-not $name){$name=($job.Key -replace '[\\/:*?"<>|]','_')+'.exe'}
            $dest=Join-Path (Join-Path $root $job.Folder) $name
            $r=Invoke-OfficialDownloadAndInstall -Label $job.Label -Uri $uri -Destination $dest -Arguments $job.Args -Install
            if($r.Installed){Write-Log "[$($job.Label)] Installed and associated with live device family '$family'." 'OK'}
            else{Write-Log "[$($job.Label)] Downloaded/retained for the live device family '$family'; automatic installation did not complete." 'WARN'}
        }
    }
    Write-Log 'Device-associated manufacturer app/utility/tuning-tool pass complete. No automatic overclock, undervolt, voltage, power-limit or fan-profile changes are applied.' 'OK'
}

function Install-OfficialManufacturerFallbacks {
    param([Parameter(Mandatory)]$Devices,[Parameter(Mandatory)]$AsusPackages,[Parameter(Mandatory)]$InstalledDrivers)
    Write-Log 'STAGE 07E - Official manufacturer fallback packages and remaining ASUS/vendor utilities' 'STEP'
    $utilRoot=Join-Path $DownloadRoot '02 Official Manufacturer Utilities';New-Item -ItemType Directory -Force -Path $utilRoot|Out-Null
    $allVendors=@($Devices|ForEach-Object{Get-HardwareVendor $_}|Sort-Object -Unique);$winget=Get-Command winget.exe -ErrorAction SilentlyContinue

    # Intel XTU: current official Intel page lists 10.0.1.188 for Windows 11 25H2. The
    # fixed vendor URL is used below so an Intel webpage anti-bot 403 cannot block the download.
    if($allVendors -contains 'Intel'){
        $cpu=(Get-CimInstance Win32_Processor|Select-Object -First 1);$cpuName=[string]$cpu.Name
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
            $xtuPage='https://www.intel.com/content/www/us/en/download/17881/intel-extreme-tuning-utility-intel-xtu.html'
            $xtuUri=$null
            try{
                $xtuHtml=(Invoke-WebRequest -Uri $xtuPage -UseBasicParsing -TimeoutSec 60 -Headers @{'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v45'}).Content
                $xtuUri=[regex]::Matches([string]$xtuHtml,'https?://[^"''\s<>]+XTUSetup_10\.0\.1\.188\.exe')|ForEach-Object{$_.Value}|Select-Object -First 1
            }catch{}
            if(-not$xtuUri){$xtuUri='https://downloadmirror.intel.com/877126/XTUSetup_10.0.1.188.exe'}
            $xtuDest=Join-Path $utilRoot 'Intel XTU\XTUSetup_10.0.1.188.exe';New-Item -ItemType Directory -Force -Path (Split-Path $xtuDest -Parent)|Out-Null
            $xtu=Invoke-OfficialDownloadAndInstall -Label 'Intel XTU 10.0.1.188' -Uri $xtuUri -Destination $xtuDest -SHA256 '0118C89C059D348FF3D45E72F3CB88241EEEB7BE1FC27D284ACA8EECE5F1C7ED' -Arguments '/S' -Install
            if(-not$xtu.Downloaded -and $winget){Write-Log '[Intel] Direct XTU URL failed; using WinGet as a secondary official-publisher retrieval method.' 'WARN';Start-Process -FilePath $winget.Source -ArgumentList @('download','--id','Intel.ExtremeTuningUtility','--exact','--source','winget','--accept-package-agreements','--accept-source-agreements','--download-directory',(Split-Path $xtuDest -Parent),'--silent') -Wait -NoNewWindow -ErrorAction SilentlyContinue|Out-Null}
        }else{Write-Log "Intel XTU not auto-installed because CPU '$cpuName' was not positively matched to Intel's current XTU support family." 'INFO'}
    }

    # RAM timing/diagnostic utilities. CPU-Z exposes memory timings/SPD and HWiNFO exposes
    # detailed memory/controller telemetry. They are installed silently only after their
    # complete installers have been downloaded to the user's archive.
    $ramRoot=Join-Path $utilRoot 'RAM Timing and Hardware Utilities';New-Item -ItemType Directory -Force -Path $ramRoot|Out-Null
    $cpuzUri='https://download.cpuid.com/cpu-z/cpu-z_3.01-en.exe';$cpuzDest=Join-Path $ramRoot 'CPU-Z 3.01\cpu-z_3.01-en.exe';New-Item -ItemType Directory -Force -Path (Split-Path $cpuzDest -Parent)|Out-Null
    Invoke-OfficialDownloadAndInstall -Label 'CPUID CPU-Z 3.01 (RAM timings/SPD)' -Uri $cpuzUri -Destination $cpuzDest -SHA256 '' -Arguments '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -Install|Out-Null
    $hwUri='https://www.hwinfo.com/files/hwi64_852.exe';$hwDest=Join-Path $ramRoot 'HWiNFO64 8.52\hwi64_852.exe';New-Item -ItemType Directory -Force -Path (Split-Path $hwDest -Parent)|Out-Null
    Invoke-OfficialDownloadAndInstall -Label 'HWiNFO64 8.52 (memory timings/telemetry)' -Uri $hwUri -Destination $hwDest -Arguments '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -Install|Out-Null

    # NVIDIA RTX 50-series Broadcast application and the official Broadcast SDK redistributables.
    # NVIDIA's resource page currently lists Audio Effects 1.6.1, Video Effects 0.7.6 and AR 0.8.7
    # specifically for RTX 50-series. The SDK redistributables are downloaded/extracted; NVIDIA says
    # they are intended for applications that integrate the SDK, so the script does not force-install
    # DLLs that are not standalone end-user installers.
    if($allVendors -contains 'NVIDIA'){
        $nRoot=Join-Path $utilRoot 'NVIDIA RTX 50 Series Broadcast';New-Item -ItemType Directory -Force -Path $nRoot|Out-Null
        $broadcastPage='https://www.nvidia.com/en-au/geforce/broadcasting/broadcast-app/'
        try{
            $html=(Invoke-WebRequest -Uri $broadcastPage -UseBasicParsing -TimeoutSec 60 -Headers @{'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v45'}).Content
            $links=[regex]::Matches([string]$html,'https?://[^"''\s<>]+')|ForEach-Object{$_.Value}|Where-Object{$_ -match '(?i)(NVIDIA.?Broadcast|broadcast).*(\.exe|download)'}|Select-Object -First 10
            $bUri=$links|Where-Object{$_ -match '(?i)\.exe'}|Select-Object -First 1
            if($bUri){$bDest=Join-Path $nRoot 'NVIDIA Broadcast\NVIDIA-Broadcast.exe';New-Item -ItemType Directory -Force -Path (Split-Path $bDest -Parent)|Out-Null;Invoke-OfficialDownloadAndInstall -Label 'NVIDIA Broadcast' -Uri $bUri -Destination $bDest -Arguments '/S' -Install|Out-Null}else{Write-Log 'NVIDIA Broadcast direct installer URL was not exposed in the official page HTML; using WinGet publisher package as fallback.' 'WARN';if($winget){$bd=Join-Path $nRoot 'NVIDIA Broadcast';New-Item -ItemType Directory -Force -Path $bd|Out-Null;Start-Process -FilePath $winget.Source -ArgumentList @('download','--id','Nvidia.Broadcast','--exact','--source','winget','--accept-package-agreements','--accept-source-agreements','--download-directory',$bd) -Wait -NoNewWindow -ErrorAction SilentlyContinue|Out-Null}}
        }catch{Write-Log "NVIDIA Broadcast page lookup failed: $($_.Exception.Message)" 'WARN'}

        # SDK page: parse direct links and select RTX 50 variants for Audio, Video and AR.
        $sdkPage='https://www.nvidia.com/en-au/geforce/broadcasting/broadcast-sdk/resources/'
        try{
            $sdkHtml=(Invoke-WebRequest -Uri $sdkPage -UseBasicParsing -TimeoutSec 60 -Headers @{'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v45'}).Content
            $allLinks=[regex]::Matches([string]$sdkHtml,'https?://[^"''\s<>]+')|ForEach-Object{$_.Value}|Where-Object{$_ -match '(?i)(\.zip|\.exe|\.msi)'}|Select-Object -Unique
            $sdkItems=@(
                [pscustomobject]@{Name='Audio Effects SDK RTX 50';Pattern='(?i)audio.*(50|blackwell).*\.(zip|exe|msi)';Folder='Audio Effects SDK'},
                [pscustomobject]@{Name='Video Effects SDK RTX 50';Pattern='(?i)video.*(50|blackwell).*\.(zip|exe|msi)';Folder='Video Effects SDK'},
                [pscustomobject]@{Name='AR SDK RTX 50';Pattern='(?i)(^|/|_)ar.*(50|blackwell).*\.(zip|exe|msi)';Folder='AR SDK'}
            )
            foreach($item in $sdkItems){$uri=$allLinks|Where-Object{$_ -match $item.Pattern}|Select-Object -First 1;if($uri){$ext=[IO.Path]::GetExtension(([uri]$uri).AbsolutePath);$dest=Join-Path (Join-Path $nRoot $item.Folder) ([IO.Path]::GetFileName(([uri]$uri).AbsolutePath));New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent)|Out-Null;Invoke-FileDownloadWithProgress -Uri $uri -Destination $dest -Label $item.Name|Out-Null;try{if($ext -ieq '.zip'){Expand-Archive -LiteralPath $dest -DestinationPath (Join-Path (Split-Path $dest -Parent) 'Extracted') -Force -ErrorAction Stop;Write-Log "[$($item.Name)] SDK archive extracted; redistributable files retained for applications that integrate NVIDIA Broadcast." 'OK'}}catch{Write-Log "[$($item.Name)] SDK archive could not be extracted: $($_.Exception.Message)" 'WARN'}}else{Write-Log "[$($item.Name)] RTX 50-series direct SDK download link was not exposed by NVIDIA's current page response." 'WARN'}}
        }catch{Write-Log "NVIDIA Broadcast SDK page lookup failed: $($_.Exception.Message)" 'WARN'}
    }

    # Full offline Armoury Crate package. ASUS currently publishes 1.5.0.7 / 4.99 GB from the
    # official support page. It is downloaded in full; the package is intentionally not treated as
    # a generic /S executable because ASUS documents device selection/installation behavior for it.
    $acRoot=Join-Path $utilRoot 'ASUS Armoury Crate';New-Item -ItemType Directory -Force -Path $acRoot|Out-Null
    try{
        $acHtml=(Invoke-WebRequest -Uri 'https://www.asus.com/au/supportonly/armoury%20crate/helpdesk_download/' -UseBasicParsing -TimeoutSec 60 -Headers @{'User-Agent'='Mozilla/5.0 (Windows NT 10.0; Win64; x64) ROG-G635LW-Installer-v45'}).Content
        $acLinks=[regex]::Matches([string]$acHtml,'https?://[^"''\s<>]+')|ForEach-Object{$_.Value}|Where-Object{$_ -match '(?i)dlcdnets\.asus\.com|dlcdnwebimgs\.asus\.com'}|Where-Object{$_ -match '(?i)Armoury|Armoury_Crate|ArmouryCrate|Full_Installation|Full.*Package'}|Select-Object -Unique
        $acFull=$acLinks|Where-Object{$_ -match '(?i)Full.*(Installation|Package)|Armoury_Crate_Full'}|Select-Object -First 1
        $acSmall=$acLinks|Where-Object{$_ -notmatch '(?i)Full.*(Installation|Package)' -and $_ -match '(?i)ArmouryCrateInstallTool|Armoury.*Installer'}|Select-Object -First 1
        if($acFull){$acDest=Join-Path $acRoot 'Armoury Crate Full Installation Package.zip';Invoke-FileDownloadWithProgress -Uri $acFull -Destination $acDest -Label 'ASUS Armoury Crate Full Installation Package' -ExpectedSHA256 'FCF9D0ECF3350837F68F1223F452CB0E1CD7B8F28753AC6ECC4376C792ACBD87'|Out-Null}else{Write-Log 'ASUS Armoury Crate full-package URL was not exposed in raw HTML; the ASUS catalog phase will still download it if the G635LW catalog provides it.' 'WARN'}
        if($acSmall){$acSmallDest=Join-Path $acRoot 'Armoury Crate & Aura Creator Installer.zip';Invoke-FileDownloadWithProgress -Uri $acSmall -Destination $acSmallDest -Label 'ASUS Armoury Crate & Aura Creator Installer' -ExpectedSHA256 'A0C3181B135C0439F12338D6D6B77181B54A62BC44C1FE95C57945F1735D54E1'|Out-Null}
    }catch{Write-Log "Armoury Crate full-package lookup failed: $($_.Exception.Message)" 'WARN'}

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
        if(-not $ready){Invoke-FileDownloadWithProgress -Uri $p.DownloadUrl -Destination $dest -Label "Final ASUS prerequisite $($p.Title)" -ExpectedSHA256 ([string]$p.SHA256)|Out-Null}
        if(Test-Path $dest){
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
        if(-not $ready){$ready=Invoke-FileDownloadWithProgress -Uri $ac.DownloadUrl -Destination $dest -Label 'FINAL ASUS Armoury Crate & Aura Creator Installer' -ExpectedSHA256 ([string]$ac.SHA256)}
        if($ready){
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

    # Armoury Crate/Aura Store companions are handled in the earlier Store stage. Do not
    # query AppX again here: the STAGE 06 installed-software baseline remains authoritative.
    Write-Log 'Final Armoury/Aura Store handling: Microsoft Store companions were handled by the earlier Store stage; no second AppX inventory query is performed.' 'INFO'
    Write-Log 'FINAL ASUS ARMOURY CRATE / AURA stage complete. No PnP or installed-state re-detection was performed.' 'OK'
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
        'ROG STRIX SCAR 16 G635LW - WINDOWS 11 25H2 INSTALLER v42'
        "Completed: $(Get-Date)"
        "Model: $Model"
        "Windows: $($OS.DisplayVersion) build $($OS.Build).$($OS.UBR)"
        "BIOS: $($Bios.SMBIOSBIOSVersion)"
        "Update passes completed after hardware detection: $UpdatePass"
        "Automatic reboots used: $RebootCount"
        "PnP configuration errors in final inventory: $($bad.Count)"
        "Driver inventory: $driverInventory"
        "Logs: $LogDir"
        "Driver/software archive: $DownloadRoot"
        "NOTE: ASUS full setup packages are attempted first after detection from the official G635LW catalog. The ASUS package phase has an independent persistent checkpoint and a Desktop directory check; a reboot/cancel/failure resumes at the first package not marked Installed, without using package-number position as the resume key. Official manufacturer driver stores are then queried for hardware families not represented in the ASUS catalog, using the same live hardware/date/version association logic; directly associated manufacturer apps/utilities/tweak tools are then downloaded/installed; Windows/Microsoft Update remains the final fallback. No automatic overclock/undervolt/tuning profile is applied."
    ) | Set-Content -LiteralPath $ReportFile -Encoding UTF8

    Write-Log "Final report: $ReportFile" 'OK'
    Write-Log "Driver inventory: $driverInventory" 'OK'
}

# ---------------- MAIN ----------------

Write-Log 'ROG STRIX SCAR 16 G635LW - WINDOWS 11 25H2 INSTALLER v42' 'STEP'
Write-Log 'Use -Reset to remove the v42/v37/v36/v35/v34/v33/v32/v31/v30/v29/v28/v27/v26/v21/v20/v19/v18/v17/v16/v15 installer archive, checkpoint/state files and scheduled resume task without uninstalling drivers or software.' 'INFO'
Write-Log 'v42 deliberately keeps manually launched terminal sessions open after completion so the complete shell output can be reviewed.' 'INFO'
Write-Log 'v42 deliberately defers the main downloading/installing work until AFTER STAGE 06 hardware detection.' 'INFO'
Write-Log 'Uses the built-in Windows Update Agent; no PSWindowsUpdate module is required.' 'INFO'
Write-Log "Target OS architecture: $(Get-OSArchitecturePreference). ALL architecture variants are retained; package priority is x64 -> ARM64 -> x86 -> neutral/unspecified. Architecture is a priority, not an exclusion filter." 'INFO'
Write-Log 'ASUS full setup packages are preferred after detection; BIOS/firmware flashing remains excluded from unattended execution.' 'INFO'
if ($Resume) {
    Write-Log 'Resumed after reboot; continuing directly with post-detection hardware/update processing.' 'INFO'
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
    Prepare-UpdateServices
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
# is resumed. This is independent of the Windows Update pass counter.
if (-not $asusPackagePhaseComplete) {
    # MyASUS is a required Microsoft Store package. The only authoritative software
    # inventory is the STAGE 06 baseline, so if that baseline did not contain MyASUS,
    # attempt its Store installation on every invocation until it succeeds. We do not
    # perform a later software inventory just to decide whether it installed.
    $myAsusBaselineInstalled=@($installedSoftware | Where-Object { [string]$_.DisplayName -match '(?i)^ASUS MyASUS$|^MyASUS$' })
    if($myAsusBaselineInstalled.Count -gt 0){
        Write-Log 'MyASUS is present in the authoritative STAGE 06 installed-software baseline; no Store installation is required for this invocation.' 'OK'
    } else {
        $myAsusResult=Install-MyASUS
        if(-not $myAsusResult){
            $asusPackagePhaseComplete=$false
            $postDetectionComplete=$true
            Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $true -RebootPending $false -AsusPackagePhaseComplete $false
            Register-ResumeTask
            Write-Log 'Required MyASUS installation did not complete. Stopping before ASUS full packages; the startup resume task and Desktop archive are preserved.' 'ERROR'
            [void](Wait-ForUserBeforeExit -ExitCode 2 -Reason 'Required MyASUS Microsoft Store installation did not complete. The installer will automatically resume at the saved checkpoint after the next Windows reboot.')
            exit 2
        }
    }

    $asusCatalog = @(Get-AsusOfficialPackages)
    $asusFullOutput = @(Install-AsusFullPackages -Devices $devices -Packages $asusCatalog -InstalledDrivers $installedDrivers -InstalledSoftware $installedSoftware)
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
        Write-Log "ASUS package phase paused because the user deliberately selected N at the reboot prompt. This is NOT an installer failure. The startup resume task is preserved and will automatically resume v44 immediately when Windows is rebooted. The Desktop archive and resume state are preserved." 'INFO'
        [void](Wait-ForUserBeforeExit -ExitCode 0 -Reason 'Reboot was deferred. Choose Restart from Windows when ready; the installer is registered to auto-resume at Windows startup from the saved package checkpoint.')
        exit 0
    }

    if($asusFull.StoppedForInstallerFailure){
        # IMPORTANT: do not proceed to later ASUS packages, manufacturer utilities, or Windows Update.
        # The exact cancelled/failed package remains the resume point and its downloaded file is retained.
        $asusPackagePhaseComplete=$false
        $postDetectionComplete=$true
        Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $postDetectionComplete -RebootPending $false -AsusPackagePhaseComplete $false
        Write-Log "ASUS package installation failed/cancelled at the current package boundary. Re-run v44 to resume from the Desktop archive/checkpoint; it will not jump ahead." 'WARN'
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
        if ($rebootCount -lt 4) {
            $rebootCount++
            Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $true -RebootPending $true -AsusPackagePhaseComplete $false
            Register-ResumeTask
            Write-Log "A reboot is required/was explicitly selected. Package checkpoint is saved. Automatic reboot $rebootCount of 4 in 15 seconds; startup will re-run the Desktop directory check before continuing." 'WARN'
            Start-Sleep 15
            Restart-Computer -Force
            exit
        }
        Write-Log 'ASUS full installers requested a reboot but the automatic reboot limit has been reached. Package phase remains incomplete so a later manual run can resume it.' 'WARN'
        [void](Wait-ForUserBeforeExit -ExitCode 2 -Reason 'The automatic reboot limit was reached before the ASUS package phase completed. The checkpoint was preserved.')
        exit 2
    }

    if($asusFull.DeferredStore){
        # Microsoft Store companion apps are non-driver software. They must never block
        # the ASUS driver/package and manufacturer fallback work. Keep the Store package
        # in the persistent resume state, leave the ASUS phase incomplete, but continue
        # with non-Store manufacturer utilities and Windows Update in this invocation.
        $asusPackagePhaseComplete=$false
        $postDetectionComplete=$true
        Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $true -RebootPending $false -AsusPackagePhaseComplete $false
        Write-Log 'One or more Microsoft Store companion packages were deferred because the Store transaction did not complete. This is NOT an ASUS driver failure. Continuing with manufacturer utilities and Windows Update; the deferred Store package remains the resume point for the next invocation.' 'WARN'
        Write-Log 'STAGE 07C/07F - Continuing with the unchanged STAGE 06 hardware/driver/software baseline; no PnP rescan or inventory refresh is performed.' 'INFO'
        Install-RelatedMicrosoftStoreApps -Devices $devices -AsusPackages $asusCatalog | Out-Null
        Install-OfficialManufacturerDriverStorePass -Devices $devices -AsusPackages $asusCatalog -InstalledDrivers $installedDrivers
        Install-DeviceAssociatedManufacturerUtilities -Devices $devices -InstalledSoftware $installedSoftware
        Install-OfficialManufacturerFallbacks -Devices $devices -AsusPackages $asusCatalog -InstalledDrivers $installedDrivers
        Install-FinalArmouryCrateAuraStage -Devices $devices -AsusPackages $asusCatalog -InstalledSoftware $installedSoftware
        Write-Log 'Non-Store ASUS/manufacturer work completed. The ASUS package checkpoint remains incomplete only because a Microsoft Store companion was deferred.' 'OK'
    }
    elseif(-not $asusPackagePhaseComplete){
        Write-Log 'ASUS package phase did not reach a clean completion checkpoint. Stopping before manufacturer utilities/Windows Update so the next run resumes the exact outstanding package.' 'WARN'
        Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $true -RebootPending $false -AsusPackagePhaseComplete $false
        [void](Wait-ForUserBeforeExit -ExitCode 2 -Reason 'The ASUS package phase did not reach a clean completion checkpoint. The current package remains the resume point.')
        exit 2
    }

    else {
        # Only after every ASUS package checkpoint is complete do manufacturer driver-store
        # reconciliation and then manufacturer utilities/companions run.
        Write-Log 'STAGE 07C/07F - Using the unchanged STAGE 06 hardware/driver/software baseline; no PnP rescan or inventory refresh is performed.' 'INFO'
        Install-RelatedMicrosoftStoreApps -Devices $devices -AsusPackages $asusCatalog | Out-Null
        Install-OfficialManufacturerDriverStorePass -Devices $devices -AsusPackages $asusCatalog -InstalledDrivers $installedDrivers
        Install-DeviceAssociatedManufacturerUtilities -Devices $devices -InstalledSoftware $installedSoftware
        Install-OfficialManufacturerFallbacks -Devices $devices -AsusPackages $asusCatalog -InstalledDrivers $installedDrivers
        Install-FinalArmouryCrateAuraStage -Devices $devices -AsusPackages $asusCatalog -InstalledSoftware $installedSoftware
        Write-Log 'ASUS full setup package phase, manufacturer fallback phase, and final Armoury Crate/Aura stage are complete; continuing to Windows/Microsoft Update.' 'OK'
    }
}
else {
    Write-Log 'ASUS package phase is already complete according to the persistent checkpoint; no ASUS package is downloaded or installed again.' 'OK'
    # The final Armoury/Aura stage is still run on a resumed invocation so it remains the last
    # ASUS software stage even when the main ASUS package phase was completed on an earlier boot.
    $resumeAsusCatalog=@(Get-AsusOfficialPackages)
    Install-FinalArmouryCrateAuraStage -Devices $devices -AsusPackages $resumeAsusCatalog -InstalledSoftware $installedSoftware
}

# Run an extended repeated post-detection Windows/Microsoft Update cycle. This allows
# newly installed hardware drivers and cumulative updates to expose additional applicable
# updates on later scans.
$MaxUpdatePasses = 8
Write-Log 'Windows/Microsoft Update is retained as the final fallback and will run up to 8 post-detection passes, retaining the one-time STAGE 06 hardware/driver/software baseline for every pass.' 'INFO'

while ($updatePass -lt $MaxUpdatePasses) {
    $updatePass++
    $r = Install-WindowsUpdatePass -PassNumber $updatePass -Devices $devices -InstalledDrivers $installedDrivers

    # v33 deliberately does NOT re-enumerate after Windows Update. The one-time STAGE 06
    # baseline remains the only hardware/driver/software inventory for this invocation.
    Write-Log "Update pass $updatePass completed; retaining the original STAGE 06 hardware/driver/software baseline. No PnP rescan or inventory refresh is performed." 'INFO'

    if ($r.RebootRequired) {
        if ($rebootCount -lt 4) {
            $rebootCount++
            Save-State -RebootCount $rebootCount -UpdatePass $updatePass -PostDetectionComplete $true -RebootPending $true -AsusPackagePhaseComplete $asusPackagePhaseComplete
            Register-ResumeTask
            Write-Log "Update pass $updatePass requires a reboot. Automatic reboot $rebootCount of 4 in 15 seconds." 'WARN'
            Start-Sleep 15
            Restart-Computer -Force
            exit
        } else {
            Write-Log 'The maximum automatic reboot count (4) has been reached. Continuing to final verification without another automatic reboot.' 'WARN'
            break
        }
    }

    # A pass with no installed updates means Windows Update has reached a stable state.
    # If packages were installed, do one more scan so newly exposed drivers/updates are found.
    if ($r.SearchFailed) {
        Write-Log "Windows Update pass $updatePass could not complete an online WUA assessment. The original hardware/driver/software baseline is retained; continuing to the next pass." 'WARN'
        continue
    }
    if ($r.InstalledCount -eq 0) {
        Write-Log "Update pass $updatePass found nothing new. Post-detection update cycle is complete." 'OK'
        break
    }

    Write-Log "Update pass $updatePass installed $($r.InstalledCount) package(s), including $($r.DriverCount) driver-category update(s). Running another full scan." 'OK'
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
Write-Log 'VERSION 41 RUN COMPLETE' 'STEP'
Write-Log "Logs: $LogDir"
Write-Log "Driver/software archive: $DownloadRoot"
Write-Log "Report: $ReportFile"
[void](Wait-ForUserBeforeExit -ExitCode 0 -Reason 'All requested installer stages reached their completion checkpoint. The final verification report has been written.')
