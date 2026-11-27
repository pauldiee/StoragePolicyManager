# StoragePolicyManager

A PowerShell utility to export and import vCenter VM Storage Policies (SPBM) — useful for migrating policies between vCenter instances or backing them up before upgrades.

| Script | Version | Purpose |
|---|---|---|
| `Invoke-StoragePolicyManager.ps1` | 1.5.0 | Export / import vCenter storage policies — **vSphere 9 only** |
| `tests/New-SPMTestSet.ps1` | 1.0.1 | Builds, shows and removes a throwaway set of storage policies in a **lab** vCenter |
| `tests/New-SPMScaleSet.ps1` | 0.1.0 | Builds and removes a large set of storage policies in a **lab** vCenter, to test at scale |

---

## Step-by-step guide

StoragePolicyManager is step 4 of carrying the vCenter configuration over to a new VCF 9 vCenter in a side-by-side migration. The guide on docs.hollebollevsan.nl walks through it, with the commands and a check per step:

- [Carrying over the vCenter configuration, step 4: storage policies](https://docs.hollebollevsan.nl/docs/30-carry-over-vcenter-configuration/#step-4-storage-policies): export the policies from the source, import or map them on the target, and review the rules that were skipped.
- [The order, and why](https://docs.hollebollevsan.nl/docs/30-carry-over-vcenter-configuration/#the-order-and-why): storage policies come after the tags (step 2, VITagManager), because tag-based rules need them.
- [Side-by-side migration: build new, then move the workloads](https://docs.hollebollevsan.nl/docs/29-side-by-side-migration/#toolkits): the whole route, and how this toolkit fits with the others.

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
| PowerShell 7.4+ | VCF.PowerCLI 9.x needs PowerShell 7 ([KB 409981](https://knowledge.broadcom.com/external/article/409981)); Windows ships 5.1, so start the script from `pwsh` |
| VCF.PowerCLI 9.0+ or VMware.PowerCLI 13+ | `Install-Module -Name VCF.PowerCLI -Scope CurrentUser` |
| Network access | HTTPS to vCenter Server |

Setting up PowerShell 7 and VCF.PowerCLI on a jump host, with or without internet access, and keeping one PowerCLI version per host: [Jump Host: PowerShell 7 and VCF PowerCLI](https://vcf-planning.hollebollevsan.nl/docs/42-jump-host-powershell/).

## Usage

```powershell
# Interactive export — lists all policies, prompts for selection
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc01.vcf.lab -Mode Export

# Fully interactive, including the vCenter (saved ones as numbers)
.\Invoke-StoragePolicyManager.ps1

# Fully interactive for one vCenter: mode menu, folder prompt, picker
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc01.vcf.lab

# Non-interactive export — one policy, or several names and wildcards
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc01.vcf.lab -Mode Export -PolicyName "vSAN Gold"
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc01.vcf.lab -Mode Export -PolicyName 'vSAN*','Gold*' -FilePath C:\PolicyBackups

# Import every policy export in a folder, without the picker
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc02.vcf.lab -Mode Import -FilePath 'C:\PolicyBackups\*.json'

# Preview the same import: what would be created, nothing changed
.\Invoke-StoragePolicyManager.ps1 -vCenterServer vc02.vcf.lab -Mode Import -FilePath 'C:\PolicyBackups\*.json' -DryRun

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

The first run asks for the vCenter login and saves it; see [Credentials](#credentials).

## How a run goes

Without `-Mode` a menu asks for Export or Import. Without `-FilePath` the script then asks for the folder: **Enter** keeps the script directory (the default), **B** opens the Windows folder browser (only where a GUI is available), or type a path. Import shows how many JSON files the folder holds and asks again when it holds none. An export with `-PolicyName` doesn't ask. Everything the console shows, the per-rule warnings included, also goes to `Logs\StoragePolicyManager_<timestamp>.log` next to the script (`-LogDirectory` to change it). Long steps show a progress bar with "n of N: item", the percentage and the time left; it disappears when the step ends.

## Interactive picker

**Export** — lists all storage policies with rule set and rule counts:

```
  Policies on this vCenter: select the ones to EXPORT
  [1]  vSAN Default Storage Policy                   (1 rule set(s), 4 rule(s))
  [2]  vSAN Gold                                     (1 rule set(s), 3 rule(s))
  [3]  vSAN Silver                                   (1 rule set(s), 2 rule(s))
```

**Import** — lists JSON files found in the directory with policy name and rule set count:

```
  JSON files in C:\PolicyBackups: select the ones to IMPORT
  [1]  vSAN_Gold.json                      policy: vSAN Gold                      (1 rule set(s))
  [2]  vSAN_Silver.json                    policy: vSAN Silver                    (1 rule set(s))
```

The picker accepts single numbers, comma-separated lists, ranges (`2-6`), combinations (`1,3-5,8`), `all`, and `cancel` (or `exit`, `quit`, `q`).

### Filtering a picker

The picker filters by name. While a filter is on, the prompt shows it, for example `Your selection (filter: /vsan, 4 shown)`.

| To | Type |
|---|---|
| Narrow the list | A wildcard (`*vsan*`) or `/text` for a substring |
| Change the filter | The new filter straight away; it searches the full list again, not within the previous filter |
| Clear the filter | `reset`, or just Enter |
| See everything in the current list | `list` |

The list is renumbered after every filter and after clearing it, so pick the number from the list shown just above the prompt. `all` takes what the filter shows. A filter that matches nothing leaves the list as it was. Above 50 items the picker shows a few examples and asks for a filter instead of printing the whole list.

## Failures and exit codes

A policy that fails to export or import is logged in red as `FAILED …`, with vCenter's own message. The summary is red when something failed, yellow when something was only skipped (the policy already exists, rule parts left out because a capability or tag is missing on the target), green otherwise, and lists every failure.

| Code | Meaning |
|---|---|
| 0 | Finished without failures (skips are allowed) |
| 1 | Something failed: see the red `FAILED` lines and the summary, or the script could not start (PowerCLI, connection, no policy matches `-PolicyName`, `-FilePath` not found) |
| 2 | Stopped because a prompt needed console input in an unattended run; the message names the parameters to pass. Save the credential first with one interactive run |

From a scheduled task or another script, start it with `pwsh -Command` rather than `pwsh -File`: with `-File`, `-PolicyName 'a','b'` arrives as one string.

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-vCenterServer` | `string` | *(asked)* | FQDN or IP of the vCenter Server. Left out: the script asks for it and lists the vCenters with a saved credential as numbers (1.4.0) |
| `-Mode` | `Export\|Import` | *(menu)* | Operation mode; a menu asks when omitted |
| `-PolicyName` | `string[]` | *(interactive)* | Export: policy names or wildcards — omit to use the interactive picker |
| `-FilePath` | `string` | Next to script | Export: output directory. Import: a JSON file, a wildcard (every match is imported) or a folder to pick from. Asked after the mode when omitted |
| `-NewPolicyName` | `string` | *(from file)* | Rename the policy on import (single file only) |
| `-LogDirectory` | `string` | `Logs\` next to script | Where the log file goes |
| `-CredentialPath` | `string` | Next to script | Path to the encrypted credential file |
| `-SkipCertificateValidation` | `switch` | — | Accept a certificate this machine does not trust (lab use). Without it the script asks whether to go on without the check for this run; an unattended run stops with exit code 2 (before W38 such a vCenter was connected silently) (1.4.0) |
| `-ResetCredentials` | `switch` | — | Force a new credential prompt |
| `-DryRun` | `switch` | — | Import: show what would be created, with the rule parts that would be left out (`[DryRun] Would ...` lines, "would be" in the summary), without changing anything (1.5.0). Ignored with a warning in Export mode |

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

## Testing in a lab

`tests/New-SPMTestSet.ps1` builds a tag category `SPM-T-Tier` (Gold, Silver, Bronze) and five policies in a **lab** vCenter (`-Mode Create`): one tag rule, two rule sets (Gold OR Silver), vSAN FTT 1, vSAN FTT 1 with stripe width 2 and force provisioning, and a policy on the Bronze tag. `-Mode Tags` creates only the category with Gold and Silver: run it on the target first (tags go across before policies), so the Bronze policy meets a tag that doesn't exist there. `-Mode Show` lists the rule sets; `-Mode Remove` deletes the `SPM-T*` policies and the tag category. It only touches `SPM-T*` objects.

Suggested run: Create on the source, Tags on the target, Export each policy with `-PolicyName` and `-FilePath <folder>`, Import on the target, Show on both sides, Remove on both.

Lab findings (vCenter 8.0.3 build 25600417 to 9.1.1 build 25784172, 2026-11-16, clean `pwsh -NoProfile`):

- Four of five policies imported with the same rule sets, the vSAN ones and the two-rule-set policy included.
- The Bronze policy reports the missing tag (*"not found — skipped"*) and then fails with SPBM's *"The storage policy cannot have both CommonRule and AnyOfRuleSets null or empty"* (#6).
- For Export, `-FilePath` is a folder; a file name is ignored and the file lands next to the script (#6).

`tests/New-SPMScaleSet.ps1` builds a large set for scale tests: a tag category `SPM-S-Tier` with 20 tags and 150 policies (`-PolicyCount`) that take turns between tag rules (one to three rule sets), vSAN rules (failures to tolerate, stripe width, object space reservation, force provisioning) and a vSAN rule plus a tag rule in one rule set. `-Mode Tags` creates only the tags (for an import target), `-Mode Remove` cleans up. Scale on VCF 9.1.1 (2026-11-19): 150 policies exported in 4 s and imported in about 80 s, and a re-export after the import was identical. Test cases: `tests/TESTPLAN.md`.

## Credentials

Credentials are encrypted and saved as `<hostname>.cred` next to the script using `Export-Clixml` (DPAPI-protected, tied to the current Windows user). `*.cred` files are git-ignored.

## Author

Paul van Dieen — [hollebollevsan.nl](https://www.hollebollevsan.nl)
