#Requires -Version 7.4

<#
.SYNOPSIS
    Builds or removes a large throwaway set of storage policies in a LAB
    vCenter, to test Invoke-StoragePolicyManager.ps1 at scale.

.DESCRIPTION
    Lab use only. Everything it creates is named SPM-S*, and Remove only
    deletes SPM-S-* policies and the SPM-S-Tier tag category.

    Create  A tag category SPM-S-Tier (datastores) with 20 tags
            (SPM-S-Tag-01 .. 20) and -PolicyCount policies (default 150),
            SPM-S-001 and up, in three kinds that take turns:
              tag      one to three rule sets, each one tag rule (OR)
              vSAN     failures to tolerate 0-2, stripe width 1-4, object
                       space reservation 0/25/50/100, force provisioning
                       on every other one
              mixed    a vSAN rule (failures to tolerate 1) and a tag rule
                       in one rule set
            Running Create again adds what is missing.
    Tags    Only the tag category and its 20 tags: run it on an import
            target first (tags go across before policies).
    Remove  Deletes every SPM-S-* policy and the SPM-S-Tier tag category
            (when no object carries its tags).

.PARAMETER vCenterServer
    FQDN or IP of the LAB vCenter Server.

.PARAMETER Mode
    Create, Tags or Remove.

.PARAMETER PolicyCount
    Number of policies (Create). Default 150.

.PARAMETER CredentialPath
    Saved credential file. Defaults to <vCenterServer>.cred in the repo root
    (where Invoke-StoragePolicyManager.ps1 keeps it).

.PARAMETER SkipCertificateValidation
    Skip TLS certificate validation (self-signed lab certificates).

.EXAMPLE
    .\New-SPMScaleSet.ps1 -vCenterServer vc-lab.lab.local -Mode Create

.EXAMPLE
    .\New-SPMScaleSet.ps1 -vCenterServer vc-lab.lab.local -Mode Remove

.NOTES
    Script   : New-SPMScaleSet.ps1
    Version  : 0.1.0
    Author   : Paul van Dieen
    Blog     : https://www.hollebollevsan.nl
    Requires : PowerShell 7.4+, VCF.PowerCLI 9.0+ or VMware.PowerCLI 13+
    Tested   : Lab, 2026-11-19: VCF 9.1.1 (Holodeck)
#>

<# CHANGELOG
    v0.1.0  2026-11-19  Paul van Dieen
        - Initial version: Create and Remove of a large policy set for scale
          tests of Invoke-StoragePolicyManager.ps1 (#14)
#>

param(
    [Parameter(Mandatory = $true)][string]$vCenterServer,
    [Parameter(Mandatory = $true)][ValidateSet('Create', 'Tags', 'Remove')][string]$Mode,
    [ValidateRange(1, 2000)][int]$PolicyCount = 150,
    [string]$CredentialPath = '',
    [switch]$SkipCertificateValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptVersion = '0.1.0'
$scriptAuthor  = 'Paul van Dieen'
$scriptBlogUrl = 'https://www.hollebollevsan.nl'
$repoRoot      = Split-Path $PSScriptRoot -Parent
$prefix        = 'SPM-S-'
$catName       = 'SPM-S-Tier'
$tagCount      = 20

Write-Host ('=' * 62) -ForegroundColor DarkCyan
Write-Host ('  New-SPMScaleSet.ps1' + (' ' * 26) + "v$scriptVersion") -ForegroundColor Cyan
Write-Host "  Author : $scriptAuthor"  -ForegroundColor Cyan
Write-Host "  Blog   : $scriptBlogUrl" -ForegroundColor DarkGray
Write-Host '  LAB USE ONLY - touches SPM-S* objects only' -ForegroundColor Yellow
Write-Host ('=' * 62) -ForegroundColor DarkCyan

function Say { param([string]$Text, [string]$Color = 'Cyan') Write-Host "  $Text" -ForegroundColor $Color }

$safeVc = $vCenterServer -replace '[^\w\-.]', '_'
if (-not $CredentialPath) { $CredentialPath = Join-Path $repoRoot "$safeVc.cred" }
$credential = if (Test-Path $CredentialPath) { Import-Clixml -Path $CredentialPath } else { Get-Credential -Message "Credentials for $vCenterServer" }
$null = Set-PowerCLIConfiguration -InvalidCertificateAction $(if ($SkipCertificateValidation) { 'Ignore' } else { 'Warn' }) `
    -DefaultVIServerMode Multiple -Confirm:$false -Scope Session -WarningAction SilentlyContinue
$vi = Connect-VIServer -Server $vCenterServer -Credential $credential -WarningAction SilentlyContinue
Say "Connected to $vCenterServer ($($vi.Version), build $($vi.Build))" Green

try {
    if ($Mode -in 'Create', 'Tags') {
        $started = Get-Date
        # --- Tags -----------------------------------------------------------------
        $cat = Get-TagCategory -Server $vi -Name $catName -ErrorAction SilentlyContinue
        if (-not $cat) { $cat = New-TagCategory -Server $vi -Name $catName -Cardinality Single -EntityType Datastore -Description 'StoragePolicyManager scale test - safe to delete' }
        $tagByName = @{}
        foreach ($t in @(Get-Tag -Server $vi -Category $cat -ErrorAction SilentlyContinue)) { $tagByName[$t.Name] = $t }
        for ($n = 1; $n -le $tagCount; $n++) {
            $tn = 'SPM-S-Tag-{0:D2}' -f $n
            if (-not $tagByName.ContainsKey($tn)) { $tagByName[$tn] = New-Tag -Server $vi -Name $tn -Category $cat }
        }
        Say "Tag category $catName with $tagCount tags"
        if ($Mode -eq 'Tags') { return }

        # --- Capabilities, read once ----------------------------------------------
        $caps = @{}
        foreach ($c in @(Get-SpbmCapability -Server $vi)) { $caps[$c.Name] = $c }
        foreach ($need in 'VSAN.hostFailuresToTolerate', 'VSAN.stripeWidth', 'VSAN.proportionalCapacity', 'VSAN.forceProvisioning') {
            if (-not $caps.ContainsKey($need)) { throw "Capability $need not found on this vCenter (no vSAN?)." }
        }
        function Cap { param([string]$Name, $Value) New-SpbmRule -Server $vi -Capability $caps[$Name] -Value $Value }
        function TagRule { param([int]$No) New-SpbmRule -Server $vi -AnyOfTags $tagByName[('SPM-S-Tag-{0:D2}' -f $No)] }

        # --- Policies ---------------------------------------------------------------
        $existing = @{}
        foreach ($p in @(Get-SpbmStoragePolicy -Server $vi | Where-Object { $_.Name -like "$prefix*" })) { $existing[$p.Name] = 1 }
        $made = 0; $had = 0
        $kinds = @{ tag = 0; vsan = 0; mixed = 0 }
        for ($i = 1; $i -le $PolicyCount; $i++) {
            $name = '{0}{1:D3}' -f $prefix, $i
            $kind = @('tag', 'vsan', 'mixed')[($i - 1) % 3]
            $kinds[$kind]++
            if ($existing.ContainsKey($name)) { $had++; continue }
            $sets = switch ($kind) {
                'tag' {
                    $setCount = 1 + ($i % 3)
                    for ($s = 0; $s -lt $setCount; $s++) { New-SpbmRuleSet -AllOfRules (TagRule (1 + (($i + $s * 7) % $tagCount))) }
                }
                'vsan' {
                    $rules = @(
                        (Cap 'VSAN.hostFailuresToTolerate' ($i % 3)),
                        (Cap 'VSAN.stripeWidth' (1 + ($i % 4))),
                        (Cap 'VSAN.proportionalCapacity' (@(0, 25, 50, 100)[$i % 4]))
                    )
                    if ($i % 2 -eq 0) { $rules += Cap 'VSAN.forceProvisioning' $true }
                    New-SpbmRuleSet -AllOfRules $rules
                }
                'mixed' {
                    New-SpbmRuleSet -AllOfRules @((Cap 'VSAN.hostFailuresToTolerate' 1), (TagRule (1 + ($i % $tagCount))))
                }
            }
            $null = New-SpbmStoragePolicy -Server $vi -Name $name -Description "StoragePolicyManager scale test ($kind)" -AnyOfRuleSets @($sets)
            $made++
            if ($i % 25 -eq 0) {
                Write-Progress -Activity 'Creating storage policies' -Status "$i of $PolicyCount" -PercentComplete ([int](100 * $i / $PolicyCount))
                Say "  $i of $PolicyCount"
            }
        }
        Write-Progress -Activity 'Creating storage policies' -Completed
        Say ("Done in {0:mm\:ss}: {1} polic(ies) created, {2} already there ({3} tag, {4} vSAN, {5} mixed)." -f ((Get-Date) - $started), $made, $had, $kinds.tag, $kinds.vsan, $kinds.mixed) Green
    }
    else {
        $pols = @(Get-SpbmStoragePolicy -Server $vi | Where-Object { $_.Name -like "$prefix*" })
        $no = 0
        foreach ($p in $pols) {
            Remove-SpbmStoragePolicy -Server $vi -StoragePolicy $p -Confirm:$false
            $no++
            if ($no % 25 -eq 0) { Say "  $no of $($pols.Count) removed" }
        }
        Say "Removed $($pols.Count) polic(ies)." Yellow
        $cat = Get-TagCategory -Server $vi -Name $catName -ErrorAction SilentlyContinue
        if ($cat) {
            $assigned = @(Get-TagAssignment -Server $vi -Category $cat -ErrorAction SilentlyContinue)
            if ($assigned.Count) { Say "Kept tag category ${catName}: $($assigned.Count) assignment(s)" Red }
            else { Remove-TagCategory -Server $vi -Category $cat -Confirm:$false; Say "Removed tag category $catName and its tags" Yellow }
        }
    }
}
finally {
    Disconnect-VIServer -Server $vi -Confirm:$false
}
