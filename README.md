# StoragePolicyManager

A PowerShell utility to export and import vCenter VM Storage Policies (SPBM) — useful for migrating policies between vCenter instances or backing them up before upgrades.

| Script | Version | Purpose |
|---|---|---|
| `Invoke-StoragePolicyManager.ps1` | 1.0.1 | Export / import vCenter storage policies — **vSphere 9 only** |

---

## What it does

- **Export** — connects to a vCenter, lists all storage policies with rule set and rule counts, and lets you select one or more to export. Each policy is saved as a portable JSON file containing the full rule structure.
- **Import** — reads previously exported JSON files and recreates the policies on a target vCenter. Capabilities and tags are resolved against the target vCenter; any that cannot be found are reported and skipped.

### Supported rule types

| Rule type | Export | Import |
|---|---|---|
| Capability-based (vSAN, vVols, etc.) | Full | Full — capabilities matched by ID |
| Tag-based | Full | Full — tags matched by category and name |
| Range values | Full | Full |
| Unknown types | Exported as-is | Skipped with warning |

## Requirements

| Requirement | Notes |
|---|---|
| PowerShell 5.1+ | Included with Windows 10 / Server 2016 and later |
| VCF.PowerCLI 9.0+ or VMware.PowerCLI 13+ | `Install-Module -Name VCF.PowerCLI -Scope CurrentUser` |
| Network access | HTTPS to vCenter Server |

## Usage

```powershell
# Interactive export — lists all policies, prompts for selection
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc01.vcf.lab -Mode Export

# Non-interactive export — export a specific policy directly
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc01.vcf.lab -Mode Export -PolicyName "vSAN Gold"

# Export to a specific directory
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc01.vcf.lab -Mode Export -FilePath C:\PolicyBackups

# Interactive import — scans script directory for JSON files
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc02.vcf.lab -Mode Import

# Import from a specific directory
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc02.vcf.lab -Mode Import -FilePath C:\PolicyBackups

# Import a specific file
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc02.vcf.lab -Mode Import -FilePath .\vSAN_Gold.json

# Import and rename the policy
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc02.vcf.lab -Mode Import -FilePath .\vSAN_Gold.json -NewPolicyName "vSAN Gold v2"

# Lab environment (self-signed certificate)
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc01.vcf.lab -Mode Export -SkipCertificateValidation

# Reset saved credentials
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc01.vcf.lab -Mode Export -ResetCredentials
```

Credentials are encrypted and saved as `<hostname>.cred` next to the script using `Export-Clixml` (DPAPI-protected, tied to the current Windows user).

## Interactive picker

**Export** — lists all storage policies with rule set and rule counts:

```
  Storage policies on vc01.vcf.lab:

   [ 1]  vSAN Default Storage Policy              (1 rule set(s), 4 rule(s))
   [ 2]  vSAN Gold                                (1 rule set(s), 3 rule(s))
   [ 3]  vSAN Silver                              (1 rule set(s), 2 rule(s))

  Enter number(s) to export (comma-separated, or 'all'):
```

**Import** — lists JSON files found in the directory with policy name and rule set count:

```
  JSON files in C:\PolicyBackups:

   [ 1]  vSAN_Gold.json          policy: vSAN Gold                    (1 rule set(s))
   [ 2]  vSAN_Silver.json        policy: vSAN Silver                  (1 rule set(s))

  Enter number(s) to import (comma-separated, or 'all'):
```

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-vCenterServer` | `string` | *(required)* | FQDN or IP of the vCenter Server |
| `-Mode` | `Export\|Import` | *(required)* | Operation mode |
| `-PolicyName` | `string` | *(interactive)* | Policy to export — omit to use the interactive picker |
| `-FilePath` | `string` | Next to script | Export: output directory. Import: JSON file or directory to scan |
| `-NewPolicyName` | `string` | *(from file)* | Rename the policy on import (single file only) |
| `-CredentialPath` | `string` | Next to script | Path to the encrypted credential file |
| `-SkipCertificateValidation` | `switch` | — | Skip TLS validation — for lab use |
| `-ResetCredentials` | `switch` | — | Force a new credential prompt |

## Export file format

```json
{
  "ExportedFrom": "vc01.vcf.lab",
  "ExportedAt": "2026-03-31 14:00:00",
  "ScriptVersion": "1.0.0",
  "PolicyName": "vSAN Gold",
  "Description": "High performance vSAN policy",
  "RuleSetCount": 1,
  "RuleSets": [
    {
      "Rules": [
        {
          "Type": "Capability",
          "CapabilityId": "VSAN.hostFailuresToTolerate",
          "CapabilityName": "Failures to tolerate",
          "ValueType": "Int32",
          "Value": "1"
        }
      ]
    }
  ],
  "CommonRules": []
}
```

---

## Author

Paul van Dieen — [hollebollevsan.nl](https://www.hollebollevsan.nl)
