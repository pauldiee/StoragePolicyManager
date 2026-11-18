#Requires -Version 5.1

<#
.SYNOPSIS
    Builds, shows or removes a throwaway set of storage policies in a LAB
    vCenter for Invoke-StoragePolicyManager.ps1.

.DESCRIPTION
    Lab use only. Everything it creates is named SPM-T*, and Remove only
    deletes SPM-T* policies and the SPM-T-Tier tag category.

    Create builds a tag category SPM-T-Tier (datastores) with the tags
    SPM-T-Gold, SPM-T-Silver and SPM-T-Bronze, and five policies:
        SPM-T-TagGold                          one tag rule (Gold)
        SPM-T-GoldOrSilver                     two rule sets: Gold OR Silver
        SPM-T-vSAN-FTT1                        vSAN, failures to tolerate 1
        SPM-T-vSAN-Stripe2-Force               vSAN, FTT 1, stripe width 2,
                                               force provisioning
        SPM-T Bronze (tag missing on target)   a tag rule on SPM-T-Bronze

    Tags creates only the tag category with SPM-T-Gold and SPM-T-Silver: run
    it on the import target first (tags go across before policies), so the
    Bronze policy meets a tag that doesn't exist there.

    Show lists every SPM-T* policy with its rule sets.

.PARAMETER vCenterServer
    FQDN or IP of the LAB vCenter Server.

.PARAMETER Mode
    Create, Tags, Show or Remove.

.PARAMETER CredentialPath
    Saved credential file. Defaults to <vCenterServer>.cred in the repo root
    (the same file Invoke-StoragePolicyManager.ps1 uses).

.PARAMETER SkipCertificateValidation
    Skip TLS certificate validation (self-signed lab certificates).

.NOTES
    Script   : New-SPMTestSet.ps1
    Version  : 1.0.1
    Author   : Paul van Dieen
    Blog     : https://www.hollebollevsan.nl
    Requires : VCF.PowerCLI 9.0+ or VMware.PowerCLI 13+
    Tested   : vCenter 8.0.3 (build 25600417) and 9.1.1 (lab), 2026-11-16
#>

<# CHANGELOG
    v1.0.1  2026-11-18  Paul van Dieen
        - Get-Help shows the full help again: the changelog moved out of the
          help block into its own comment block (the unknown .CHANGELOG
          keyword made Get-Help ignore the whole help block)
        - Blank line between #Requires and the help block, for the same reason

    v1.0.0  2026-11-16  Paul van Dieen
        - Initial release: Create / Tags / Show / Remove of a policy test set
#>

param(
    [Parameter(Mandatory = $true)][string]$vCenterServer,
    [Parameter(Mandatory = $true)][ValidateSet('Create', 'Tags', 'Show', 'Remove')][string]$Mode,
    [string]$CredentialPath = '',
    [switch]$SkipCertificateValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptVersion = '1.0.1'
$scriptAuthor  = 'Paul van Dieen'
$scriptBlogUrl = 'https://www.hollebollevsan.nl'
$repoRoot      = Split-Path $PSScriptRoot -Parent
$catName       = 'SPM-T-Tier'

Write-Host ('=' * 62) -ForegroundColor DarkCyan
Write-Host ('  New-SPMTestSet.ps1' + (' ' * 27) + "v$scriptVersion") -ForegroundColor Cyan
Write-Host "  $scriptAuthor - $scriptBlogUrl" -ForegroundColor Gray
Write-Host '  LAB USE ONLY - touches SPM-T* objects only' -ForegroundColor Yellow
Write-Host ('=' * 62) -ForegroundColor DarkCyan

$safeVc = $vCenterServer -replace '[^\w\-.]', '_'
if (-not $CredentialPath) { $CredentialPath = Join-Path $repoRoot "$safeVc.cred" }
if (-not (Test-Path $CredentialPath)) { throw "No credential file at $CredentialPath." }
$credential = Import-Clixml -Path $CredentialPath

$null = Set-PowerCLIConfiguration -InvalidCertificateAction $(if ($SkipCertificateValidation) { 'Ignore' } else { 'Warn' }) `
    -DefaultVIServerMode Multiple -Confirm:$false -Scope Session -WarningAction SilentlyContinue
$vi = Connect-VIServer -Server $vCenterServer -Credential $credential -WarningAction SilentlyContinue
Write-Host "  Connected to $vCenterServer ($($vi.Version), build $($vi.Build))" -ForegroundColor Green

function Confirm-Tags {
    param([string[]]$Names)
    $cat = Get-TagCategory -Server $vi -Name $catName -ErrorAction SilentlyContinue
    if (-not $cat) { $cat = New-TagCategory -Server $vi -Name $catName -Cardinality Single -EntityType Datastore; Write-Host "  Created tag category $catName" -ForegroundColor Green }
    foreach ($n in $Names) {
        if (-not (Get-Tag -Server $vi -Category $cat -Name $n -ErrorAction SilentlyContinue)) {
            $null = New-Tag -Server $vi -Name $n -Category $cat; Write-Host "  Created tag $n" -ForegroundColor Green
        }
    }
    return $cat
}
function Get-T { param([string]$Name) Get-Tag -Server $vi -Category $catName -Name $Name }
function Cap { param([string]$Name, $Value) New-SpbmRule -Server $vi -Capability (Get-SpbmCapability -Server $vi -Name $Name) -Value $Value }

try {
    switch ($Mode) {
        'Tags' { $null = Confirm-Tags 'SPM-T-Gold', 'SPM-T-Silver' }
        'Create' {
            $null = Confirm-Tags 'SPM-T-Gold', 'SPM-T-Silver', 'SPM-T-Bronze'
            $plan = [ordered]@{
                'SPM-T-TagGold'                        = { @(New-SpbmRuleSet -AllOfRules (New-SpbmRule -Server $vi -AnyOfTags (Get-T 'SPM-T-Gold'))) }
                'SPM-T-GoldOrSilver'                   = { @((New-SpbmRuleSet -AllOfRules (New-SpbmRule -Server $vi -AnyOfTags (Get-T 'SPM-T-Gold'))),
                                                             (New-SpbmRuleSet -AllOfRules (New-SpbmRule -Server $vi -AnyOfTags (Get-T 'SPM-T-Silver')))) }
                'SPM-T-vSAN-FTT1'                      = { @(New-SpbmRuleSet -AllOfRules (Cap 'VSAN.hostFailuresToTolerate' 1)) }
                'SPM-T-vSAN-Stripe2-Force'             = { @(New-SpbmRuleSet -AllOfRules @((Cap 'VSAN.hostFailuresToTolerate' 1), (Cap 'VSAN.stripeWidth' 2), (Cap 'VSAN.forceProvisioning' $true))) }
                'SPM-T Bronze (tag missing on target)' = { @(New-SpbmRuleSet -AllOfRules (New-SpbmRule -Server $vi -AnyOfTags (Get-T 'SPM-T-Bronze'))) }
            }
            foreach ($name in $plan.Keys) {
                if (Get-SpbmStoragePolicy -Server $vi -Name $name -ErrorAction SilentlyContinue) { Write-Host "  $name exists" -ForegroundColor DarkGray; continue }
                $sets = @(& $plan[$name])
                $null = New-SpbmStoragePolicy -Server $vi -Name $name -Description 'StoragePolicyManager test set' -AnyOfRuleSets $sets
                Write-Host "  Created policy $name ($($sets.Count) rule set(s))" -ForegroundColor Green
            }
        }
        'Show' {
            $rows = foreach ($p in @(Get-SpbmStoragePolicy -Server $vi | Where-Object { $_.Name -like 'SPM-T*' } | Sort-Object Name)) {
                $sets = foreach ($rs in @($p.AnyOfRuleSets)) {
                    '(' + ((@($rs.AllOfRules) | ForEach-Object {
                        if ($_.Capability) { "$($_.Capability.Name)=$($_.Value)" } else { 'tag:' + (@($_.AnyOfTags) | ForEach-Object { $_.Name }) -join '/' }
                    }) -join ' AND ') + ')'
                }
                [pscustomobject]@{ Policy = $p.Name; RuleSets = (@($sets) -join ' OR ') }
            }
            if ($rows) { $rows | Format-Table -AutoSize | Out-String -Width 240 | Write-Host } else { Write-Host '  No SPM-T* policies.' }
            $tags = @(Get-Tag -Server $vi -Category $catName -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
            Write-Host ("  Tags in ${catName}: " + $(if ($tags) { ($tags | Sort-Object) -join ', ' } else { 'none' }))
        }
        'Remove' {
            foreach ($p in @(Get-SpbmStoragePolicy -Server $vi | Where-Object { $_.Name -like 'SPM-T*' })) {
                Remove-SpbmStoragePolicy -Server $vi -StoragePolicy $p -Confirm:$false
                Write-Host "  Removed policy $($p.Name)" -ForegroundColor Yellow
            }
            $cat = Get-TagCategory -Server $vi -Name $catName -ErrorAction SilentlyContinue
            if ($cat) {
                $assigned = @(Get-TagAssignment -Server $vi -Category $cat -ErrorAction SilentlyContinue)
                if ($assigned.Count) { Write-Host "  Kept tag category ${catName}: $($assigned.Count) assignment(s)" -ForegroundColor Red }
                else { Remove-TagCategory -Server $vi -Category $cat -Confirm:$false; Write-Host "  Removed tag category $catName and its tags" -ForegroundColor Yellow }
            }
        }
    }
}
finally {
    Disconnect-VIServer -Server $vi -Confirm:$false
}
