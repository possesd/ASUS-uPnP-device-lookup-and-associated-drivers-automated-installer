# Consolidated changelog

This is the complete version progression that can be established from the supplied v45 installer and v45 release notes. Entries marked “No separate change note in supplied artifacts” preserve the version sequence without inventing undocumented changes.

## v4

No separate change note in supplied artifacts.

## v5

No separate change note in supplied artifacts.

## v6

No separate change note in supplied artifacts.

## v7

No separate change note in supplied artifacts.

## v8

No separate change note in supplied artifacts.

## v9

No separate change note in supplied artifacts.

## v10

No separate change note in supplied artifacts; later source comments note that earlier deduplication behaviour could create architecture/version gaps.

## v11

No separate change note in supplied artifacts.

## v12

- Added controlled multi-package reboot batching.
- Treats exit code 3010 as a successful install with a deferred reboot, while 1641 remains a hard stop.
- Persists batch state by Windows boot marker and respects a user choice to restart later.

## v13

No separate change note in supplied artifacts.

## v14

- Explicitly prioritised the ASUS Intel Graphics package for Core Ultra 9 275HX before other Intel platform packages and Intel XTU.

## v15

No separate change note in supplied artifacts.

## v16

- Retained all architecture variants and ordered them x64, ARM64, x86, then neutral/unspecified.

## v17

No separate change note in supplied artifacts.

## v18

No separate change note in supplied artifacts.

## v19

No separate change note in supplied artifacts.

## v20

No separate change note in supplied artifacts.

## v21

No separate change note in supplied artifacts.

## v22

No separate change note in supplied artifacts.

## v23

No separate change note in supplied artifacts.

## v24

No separate change note in supplied artifacts.

## v25

No separate change note in supplied artifacts.

## v26

No separate change note in supplied artifacts.

## v27

No separate change note in supplied artifacts.

## v28

No separate change note in supplied artifacts.

## v29

No separate change note in supplied artifacts.

## v30

- Normalised older resume-state JSON rows before property assignment.
- Fixed resume failures caused by missing `LocalPath` and related fields.
- Kept the Desktop archive location, added local-path diagnostics, and retained unbounded PnP rescan and persistent package ordering.

## v31

No separate change note in supplied artifacts.

## v32

No separate change note in supplied artifacts.

## v33

- Runs one unbounded PnP device rescan at invocation start.
- Captures one authoritative hardware, driver, and software baseline after detection.
- Avoids later rescans or inventory rebuilds and resumes after reboot at the post-detection update cycle.

## v34

No separate change note in supplied artifacts.

## v35

- Added resilient ASUS catalogue/network retries.
- Added a narrowly scoped Microsoft Store certificate-pinning retry for error `0x8A15005E` on fresh Windows 11 25H2 installations.

## v36

- Repairs or re-registers the existing Microsoft Store package for the interactive administrator before Store companion installation.

## v37

No separate change note in supplied artifacts.

## v38

No separate change note in supplied artifacts.

## v39

No separate change note in supplied artifacts.

## v40

No separate change note in supplied artifacts.

## v41

- Does not remove or reinstall Microsoft Store and does not alter DNS.
- Repairs Store registration only when required and logs an absent Store without treating it as a driver failure.

## v42

- Forces Microsoft Store companion installs into non-interactive mode.
- Adds a hard WinGet timeout and restores the temporary certificate-pinning bypass immediately after Store operations.

## v43

No separate change note in supplied artifacts.

## v44

- Adds explicit USB hardware inventory and classification.
- Retains USB hardware association and manufacturer/Windows fallback behaviour.

## v45

- Treats Realtek RTL8111H (PCI DEV_8168) and RTL8125D (PCI DEV_8125) as distinct LAN variants, and evaluates their associations before generic catalogue metadata.
- Prioritises Intel IRST/VMD immediately after Intel platform prerequisites and recognises applicable controllers by name, class, and known G635LW Arrow Lake-HX VMD PCI IDs.
- Audits the ASUS catalogue for Realtek LAN and Intel IRST/VMD candidates, logging explicit counts and warnings/errors when either critical family is absent.
- Retains complete ASUS package pre-download, persistent Desktop archive, SHA-256-aware resume, USB inventory, same-package companion executable processing, NVIDIA-before-final-Armoury ordering, Microsoft Store companions, and Windows Update fallback.
- Keeps BIOS and firmware flashing disabled.
