# ASUS uPnP device lookup and associated drivers automated installer v1.10.0

Automated Windows 11 25H2 driver and associated-software installer for the ASUS ROG Strix SCAR 16 G635LW.

## Release contents

- `ROG-G635LW-Windows11-25H2-Installer-v45.ps1` — supplied PowerShell installer.
- `ROG-G635LW-Windows11-25H2-Installer-v45.zip` — supplied archive containing the script, release changelog, and checksum manifest.
- `CHANGELOG-ROG-G635LW-Windows11-25H2-Installer-v45.txt` — supplied v45 release notes.
- `SHA256SUMS-ROG-G635LW-v45.txt` — supplied SHA-256 manifest.

## v45 scope

The supplied v45 release covers Realtek RTL8111H / RTL8125D LAN handling, Intel Rapid Storage Technology (IRST) / VMD handling, critical catalogue auditing, USB device detection and association, USB audio and HID handling, MediaTek/Android USB handling, NVIDIA RTX 5080-related processing, Microsoft Store companions, same-package companion executables, persistent downloads and resume support, Windows Update fallback, and final Armoury Crate/Aura installation after NVIDIA processing.

The installer deliberately does **not** flash BIOS or firmware automatically.

## Verification

The supplied SHA-256 manifest covers the PowerShell script only. Its expected and calculated SHA-256 value is:

```text
9e5b50ba5bd7c792f2670da73205875cb1a3f0237123c4a6c32b44c61506c99e  ROG-G635LW-Windows11-25H2-Installer-v45.ps1
```

The supplied ZIP is not listed in that manifest. Its calculated SHA-256 is:

```text
35f73cb33ef24e3f4be597d0404fb16507384de48de39ba487ac679757175c56  ROG-G635LW-Windows11-25H2-Installer-v45.zip
```

## Supplied-as-is note

This release is published exactly as supplied. Although the file names and release notes identify v45, portions of the script's embedded display/version text retain earlier v44 and v41 labels. Those source labels were not altered, preserving the verified script hash and supplied archive.

## Usage

Run from an elevated PowerShell session:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\ROG-G635LW-Windows11-25H2-Installer-v45.ps1
```

Use `-Resume` to continue from the installer checkpoint and `-Reset` to clear its persisted state before starting again.
