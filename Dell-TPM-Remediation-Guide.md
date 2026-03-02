# Dell TPM Remote Enablement via Intune — Practical Guide

---

## Step 1 — Package Dell Command | Configure (DCC) for Intune

### Download DCC
Get the latest version from Dell: https://www.dell.com/support/kbdoc/en-us/000177325

You'll get an `.exe` installer. To deploy via Intune as a Win32 app, you need to wrap it in an `.intunewin` file using the **Win32 Content Prep Tool**.

```powershell
# Download the Microsoft Win32 Content Prep Tool if you don't have it
# https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool

IntuneWinAppUtil.exe `
  -c "C:\Sources\DellCommandConfigure" `  # folder with the installer
  -s "Dell-Command-Configure.exe" `        # installer filename
  -o "C:\Output"                           # output folder for .intunewin
```

### Intune Win32 App Settings

| Field | Value |
|---|---|
| **Install command** | `Dell-Command-Configure.exe /s` |
| **Uninstall command** | `msiexec /x {PRODUCT_GUID} /quiet` |
| **Install behaviour** | System |
| **Detection rule** | File exists: `C:\Program Files\Dell\Command Configure\X86_64\cctk.exe` |

Deploy this to your **Dell device group** and make sure it's installed before the remediation runs.

---

## Step 2 — Create the Remediation Scripts

You'll need two scripts for **Intune Remediations** (Devices → Remediations).

### Detection Script
Save as `Detect-TPM.ps1`

```powershell
try {
    $tpm = Get-Tpm
    if ($tpm.TpmPresent -and $tpm.TpmReady) {
        # TPM is present and enabled - no remediation needed
        Write-Host "TPM enabled and ready."
        exit 0
    } else {
        Write-Host "TPM not ready. Remediation required."
        exit 1
    }
} catch {
    Write-Host "Error checking TPM: $_"
    exit 1
}
```

### Remediation Script
Save as `Remediate-TPM.ps1`

```powershell
$cctkPath = "C:\Program Files\Dell\Command Configure\X86_64\cctk.exe"

if (-not (Test-Path $cctkPath)) {
    Write-Host "Dell Command Configure not found. Exiting."
    exit 1
}

try {
    # Enable TPM - add --ValSetupPwd=YourPassword if you have a BIOS password set
    $result = & $cctkPath --tpm=on
    Write-Host "CCTK result: $result"

    # Schedule a reboot in 5 minutes to apply BIOS change
    shutdown /r /t 300 /c "Restarting to apply TPM BIOS settings" /f

    exit 0
} catch {
    Write-Host "Remediation failed: $_"
    exit 1
}
```

> ⚠️ **If you have a BIOS password set**, change the cctk line to:
> ```powershell
> $result = & $cctkPath --tpm=on --ValSetupPwd=YourBIOSPassword
> ```
> Be mindful of securing this — see Step 4 for options.

---

## Step 3 — Deploy in Intune

1. Go to **Intune Admin Centre → Devices → Remediations → + Create**
2. Fill in the basics (name it something like `Dell - Enable TPM`)
3. Upload your two scripts
4. Set:

| Setting | Value |
|---|---|
| **Run this script using logged-on credentials** | No (run as System) |
| **Enforce script signature check** | No (unless you sign your scripts) |
| **Run script in 64-bit PowerShell** | Yes |

5. On the **Assignments** tab, target your **Dell devices group**
6. Set the schedule — **Once** is fine initially, or daily if you want ongoing detection

---

## Step 4 — Handling the BIOS Password Securely (If Applicable)

Hardcoding a BIOS password in a script is not ideal. A cleaner approach:

- Store the password as an **Intune custom OMA-URI** or retrieve it from **Azure Key Vault** at runtime
- Alternatively, store it in an **encrypted local file** deployed separately
- For most internal IT environments, deploying via System context with Intune's secure channel is an acceptable risk if access to the script is controlled

---

## Step 5 — Verify It's Working

After the remediation runs and devices reboot, you can confirm TPM status with:

```powershell
Get-Tpm
```

Or check in Intune under **Devices → your device → Hardware** — the TPM version should populate once enabled.

You can also monitor remediation success/failure rates under **Devices → Remediations → your policy → Device status**.

---

## Summary Flow

```
Intune detects device needs remediation
        ↓
DCC already installed (Win32 app)
        ↓
Remediation script runs cctk.exe --tpm=on
        ↓
Device reboots (5 min warning)
        ↓
TPM enabled — Autopilot/BitLocker can proceed
```
