# StoragePolicyManager test plan

How to test `Invoke-StoragePolicyManager.ps1` in a lab: which lab, which
test set, which cases, and what came out. Labs are named by code only;
addresses live outside this repository.

## Labs

| Code | What | Role in the tests |
|---|---|---|
| **L8** | vCenter 8.0 U3, no hosts | Export source for vSphere 8 |
| **LA** | VCF 9.1.1 management vCenter (vSAN) | Main target; scale tests |
| **LB** | VCF 9.0.2 management vCenter | Second source/target |

## Test sets

| Script | What it builds | Cleanup |
|---|---|---|
| `New-SPMTestSet.ps1` | Functional set: tag category `SPM-T-Tier` (Gold, Silver, Bronze) and five policies (one tag rule, Gold OR Silver, vSAN FTT 1, vSAN FTT 1 + stripe 2 + force provisioning, a policy on the Bronze tag). `-Mode Tags` on the target creates Gold and Silver only | `-Mode Remove` |
| `New-SPMScaleSet.ps1` | Scale set: tag category `SPM-S-Tier` with 20 tags and `-PolicyCount` policies (default 150): tag rules (one to three rule sets), vSAN rules (FTT 0-2, stripe 1-4, object space reservation, force provisioning), and a vSAN + tag rule in one rule set. `-Mode Tags` creates the tags only | `-Mode Remove` |

## Test cases

### A. Functional (New-SPMTestSet.ps1)

| # | Lab | Steps | Expected | Result |
|---|---|---|---|---|
| F1 | L8 to LA | Create on L8, Tags on LA, Export each policy, Import on LA, Show on both, Remove on both | Four of five policies with the same rule sets; Bronze reports the missing tag | 2026-11-16: four of five **pass**; Bronze fails with SPBM's *"cannot have both CommonRule and AnyOfRuleSets null or empty"* (#6, open) |

### B. Scale (New-SPMScaleSet.ps1)

| # | Steps | Expected | Result (LA, 2026-11-19, 1.3.0) |
|---|---|---|---|
| S1 | Create 150 policies | 50 tag, 50 vSAN, 50 mixed | **pass**, 37 s |
| S2 | Export `-PolicyName 'SPM-S-*' -FilePath <folder>` | 150 files | **pass**, 4 s |
| S3 | Remove, then `-Mode Tags`, then Import the 150 files | 150 created, 0 failed; a line per policy in the console | **pass**, about 80 s (0.5 s per policy) |
| S4 | Export again, compare with S2 (without the export time) | 0 differences | **pass** |
| S5 | Remove | Policies and tag category gone | **pass** |

**1.4.0 (W38, #9), 2026-11-27, the home vCenter 8 (192.168.2.65), which was off:** start without `-vCenterServer`, certificate question, error trap. `-Mode Export`; answers from stdin.

| # | Steps | Result |
|---|---|---|
| W1 | No `-vCenterServer`; pick the vCenter by number, `y` at the certificate question, `q` at the first picker | **partly**: list `[1] 192.168.2.65`, `[2] vc-mgmt-a.site-a.vcf.lab`; `1`; "Failed to connect ... connected host has failed to respond (192.168.2.65:443)", exit 1. The certificate question is **not tested** here: no saved credential for lab B, lab A down, the home vCenter off. Same code as VITagManager 1.5.0 (V1-V6 there) |

## Open points

- Import into a second vCenter at scale (LA to LB) not yet run: needs a
  saved credential for LB.
- Scale not run on L8 (vCenter 8, no vSAN capabilities without hosts).
