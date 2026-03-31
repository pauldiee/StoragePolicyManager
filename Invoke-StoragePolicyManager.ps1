<#
.SYNOPSIS
    Export or import vCenter VM Storage Policies (SPBM).

.DESCRIPTION
    Connects to a vCenter Server and either exports storage policies to JSON
    files, or imports policies from previously exported JSON files.

    Supports capability-based rules (vSAN, vVols, etc.) and tag-based rules.
    On import, capabilities and tags are resolved against the target vCenter;
    any that cannot be found are reported and skipped.

    Credentials are encrypted and saved as <hostname>.cred next to the script
    using Export-Clixml (DPAPI-protected, tied to the current Windows user).

.PARAMETER vCenterServer
    FQDN or IP of the vCenter Server.

.PARAMETER Mode
    'Export' to save policies to JSON files, or 'Import' to create policies
    from previously exported JSON files.

.PARAMETER PolicyName
    Name of the storage policy to export. Optional — if omitted the script
    lists all policies and prompts for interactive selection.

.PARAMETER FilePath
    Path to a JSON file or directory.
    - Export: output directory. Defaults to the script directory.
    - Import: specific JSON file, or a directory to scan for JSON files.
      If omitted, the script directory is scanned.

.PARAMETER NewPolicyName
    Rename the policy on import. Only applies when importing a single file.

.PARAMETER CredentialPath
    Path to the encrypted credential file. Defaults to <vCenterServer>.cred
    next to the script.

.PARAMETER SkipCertificateValidation
    Skip TLS certificate validation. For lab use with self-signed certificates.

.PARAMETER ResetCredentials
    Force a new credential prompt even if a saved credential file exists.

.EXAMPLE
    # Interactive export — lists all policies, prompts for selection
    .\Invoke-StoragePolicyManager.ps1 -vCenterServer vc01.vcf.lab -Mode Export

.EXAMPLE
    # Non-interactive export — export a specific policy directly
    .\Invoke-StoragePolicyManager.ps1 -vCenterServer vc01.vcf.lab -Mode Export -PolicyName "vSAN Gold"

.EXAMPLE
    # Interactive import — scans script directory for JSON files
    .\Invoke-StoragePolicyManager.ps1 -vCenterServer vc02.vcf.lab -Mode Import

.EXAMPLE
    # Import a specific file
    .\Invoke-StoragePolicyManager.ps1 -vCenterServer vc02.vcf.lab -Mode Import -FilePath .\vSAN_Gold.json

.EXAMPLE
    # Lab environment with self-signed certificate
    .\Invoke-StoragePolicyManager.ps1 -vCenterServer vc01.vcf.lab -Mode Export -SkipCertificateValidation

.NOTES
    Author   : Paul van Dieen
    Blog     : https://www.hollebollevsan.nl
    Version  : 1.0.0
    Requires : VCF.PowerCLI 9.0+ (recommended) or VMware.PowerCLI 13+
    Tested   : vSphere 9

.CHANGELOG
    v1.0.0  2026-03-31  Paul van Dieen
        - Initial release
        - Export mode: interactive picker lists all storage policies with rule
          set and rule counts; each policy exported to its own JSON file
        - Import mode: interactive picker scans directory for JSON files, shows
          policy name and rule counts; imports capability-based and tag-based
          rules; missing capabilities/tags reported and skipped
        - Credentials cached via Export-Clixml (DPAPI, Windows-only)
        - All cmdlets scoped with -Server to avoid cross-vCenter operations
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$vCenterServer,
    [Parameter(Mandatory)][ValidateSet('Export','Import')][string]$Mode,
    [string]$PolicyName,
    [string]$FilePath,
    [string]$NewPolicyName,
    [string]$CredentialPath,
    [switch]$SkipCertificateValidation,
    [switch]$ResetCredentials
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptVersion = '1.0.0'
$scriptAuthor  = 'Paul van Dieen'
$scriptBlogUrl = 'https://www.hollebollevsan.nl'
$scriptDir     = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

# --- Console Banner -----------------------------------------------------------
Write-Host ('=' * 62) -ForegroundColor DarkCyan
Write-Host ("  Invoke-StoragePolicyManager.ps1" + (' ' * 14) + "v$scriptVersion") -ForegroundColor Cyan
Write-Host "  Author : $scriptAuthor"  -ForegroundColor Cyan
Write-Host "  Blog   : $scriptBlogUrl" -ForegroundColor DarkGray
Write-Host ('=' * 62) -ForegroundColor DarkCyan

# --- PowerCLI Module Check ----------------------------------------------------
$vcfModule    = Get-Module -Name VCF.PowerCLI              -ListAvailable
$legacyModule = Get-Module -Name VMware.VimAutomation.Core -ListAvailable

if (-not $vcfModule -and -not $legacyModule) {
    Write-Host "  [ERROR] No compatible PowerCLI module found." -ForegroundColor Red
    Write-Host "          Install-Module -Name VCF.PowerCLI -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

if ($vcfModule) {
    if (-not (Get-Module -Name VCF.PowerCLI)) {
        Write-Host "  [INFO] Loading VCF.PowerCLI..." -ForegroundColor Cyan
        Import-Module VCF.PowerCLI -ErrorAction Stop
    }
} else {
    if (-not (Get-Module -Name VMware.VimAutomation.Core)) {
        Write-Host "  [INFO] Loading VMware.PowerCLI..." -ForegroundColor Cyan
        Import-Module VMware.VimAutomation.Core -ErrorAction Stop
    }
}

# --- Credential Management ----------------------------------------------------
$safeVc = $vCenterServer -replace '[^\w\-.]', '_'
if (-not $CredentialPath) { $CredentialPath = Join-Path $scriptDir "$safeVc.cred" }

if ($ResetCredentials -or -not (Test-Path $CredentialPath)) {
    if ($ResetCredentials) {
        Write-Host "  [WARN] -ResetCredentials specified — prompting for new credentials." -ForegroundColor Yellow
    } else {
        Write-Host "  [INFO] No saved credentials found — prompting." -ForegroundColor Cyan
    }
    $credUser   = Read-Host "  Username for $vCenterServer"
    $credPass   = Read-Host "  Password" -AsSecureString
    $credential = [System.Management.Automation.PSCredential]::new($credUser, $credPass)
    $credential | Export-Clixml -Path $CredentialPath
    Write-Host "  [OK]   Credentials saved to $CredentialPath." -ForegroundColor Green
} else {
    Write-Host "  [INFO] Loading saved credentials from $CredentialPath." -ForegroundColor Cyan
    $credential = Import-Clixml -Path $CredentialPath
}

# --- Connect to vCenter -------------------------------------------------------
Write-Host "  [INFO] Connecting to $vCenterServer..." -ForegroundColor Cyan
try {
    $null = Set-PowerCLIConfiguration `
        -InvalidCertificateAction $(if ($SkipCertificateValidation) { 'Ignore' } else { 'Warn' }) `
        -Confirm:$false -Scope Session -WarningAction SilentlyContinue
    $viConn = Connect-VIServer -Server $vCenterServer -Credential $credential -ErrorAction Stop -WarningAction SilentlyContinue
    Write-Host "  [OK]   Connected to $vCenterServer (version $($viConn.Version), build $($viConn.Build))." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Failed to connect to $vCenterServer`: $_" -ForegroundColor Red
    exit 1
}

# --- Helper: Serialize a single rule ------------------------------------------
function Export-PolicyRule {
    param($Rule)
    # Capability-based rule
    if ($Rule.PSObject.Properties['Capability'] -and $Rule.Capability) {
        $val     = $Rule.Value
        $valType = $val.GetType().Name
        # Range value
        if ($val.PSObject.Properties['Minimum'] -and $val.PSObject.Properties['Maximum']) {
            return [PSCustomObject]@{
                Type           = 'Capability'
                CapabilityId   = [string]$Rule.Capability.Id
                CapabilityName = $Rule.Capability.Name
                ValueType      = 'Range'
                ValueMin       = [string]$val.Minimum
                ValueMax       = [string]$val.Maximum
            }
        }
        return [PSCustomObject]@{
            Type           = 'Capability'
            CapabilityId   = [string]$Rule.Capability.Id
            CapabilityName = $Rule.Capability.Name
            ValueType      = $valType
            Value          = [string]$val
        }
    }
    # Tag-based rule
    if ($Rule.PSObject.Properties['AnyOfTags'] -and $Rule.AnyOfTags) {
        $tags = @($Rule.AnyOfTags | ForEach-Object {
            [PSCustomObject]@{
                Category = if ($_.PSObject.Properties['Category'] -and $_.Category) { $_.Category.Name } else { '' }
                Name     = $_.Name
            }
        })
        return [PSCustomObject]@{
            Type = 'Tag'
            Tags = $tags
        }
    }
    # Unknown
    return [PSCustomObject]@{ Type = 'Unknown' }
}

# =============================================================================
# EXPORT
# =============================================================================
if ($Mode -eq 'Export') {
    try {
        $outDir = if ($FilePath -and (Test-Path $FilePath -PathType Container)) { $FilePath } else { $scriptDir }

        Write-Host "  [INFO] Fetching storage policies from $vCenterServer..." -ForegroundColor Cyan
        $allPolicies = @(Get-SpbmStoragePolicy -Server $viConn -ErrorAction Stop | Sort-Object Name)

        if ($allPolicies.Count -eq 0) {
            Write-Host "  [WARN] No storage policies found." -ForegroundColor Yellow
        } else {
            $policiesToExport = [System.Collections.Generic.List[object]]::new()

            if ($PolicyName) {
                # Non-interactive: specific policy supplied
                $match = $allPolicies | Where-Object { $_.Name -eq $PolicyName }
                if (-not $match) {
                    Write-Host "  [ERROR] Policy '$PolicyName' not found." -ForegroundColor Red
                    exit 1
                }
                $policiesToExport.Add($match)
            } else {
                # Interactive picker
                Write-Host ""
                Write-Host "  Storage policies on $vCenterServer`:" -ForegroundColor Cyan
                Write-Host ""
                for ($i = 0; $i -lt $allPolicies.Count; $i++) {
                    $rsCnt   = if ($allPolicies[$i].PSObject.Properties['AnyOfRuleSets'] -and $allPolicies[$i].AnyOfRuleSets) { @($allPolicies[$i].AnyOfRuleSets).Count } else { 0 }
                    $ruleCnt = 0
                    if ($rsCnt -gt 0) {
                        foreach ($rs in @($allPolicies[$i].AnyOfRuleSets)) {
                            if ($rs.PSObject.Properties['AllOfRules'] -and $rs.AllOfRules) { $ruleCnt += @($rs.AllOfRules).Count }
                        }
                    }
                    $line = "   [{0,2}]  {1,-45} ({2} rule set(s), {3} rule(s))" -f ($i + 1), $allPolicies[$i].Name, $rsCnt, $ruleCnt
                    Write-Host $line -ForegroundColor White
                }
                Write-Host ""
                $selection = Read-Host "  Enter number(s) to export (comma-separated, or 'all')"
                $selection = $selection.Trim()

                if ($selection -ieq 'all') {
                    foreach ($p in $allPolicies) { $policiesToExport.Add($p) }
                } else {
                    foreach ($token in ($selection -split ',')) {
                        $token = $token.Trim()
                        $idx   = 0
                        if ([int]::TryParse($token, [ref]$idx) -and $idx -ge 1 -and $idx -le $allPolicies.Count) {
                            $policiesToExport.Add($allPolicies[$idx - 1])
                        } else {
                            Write-Host "  [WARN] '$token' is not a valid selection — skipped." -ForegroundColor Yellow
                        }
                    }
                }
            }

            if ($policiesToExport.Count -eq 0) {
                Write-Host "  [WARN] No policies selected." -ForegroundColor Yellow
            } else {
                Write-Host ""
                foreach ($policy in $policiesToExport) {
                    try {
                        # Serialize rule sets
                        $exportRuleSets = [System.Collections.Generic.List[object]]::new()
                        if ($policy.PSObject.Properties['AnyOfRuleSets'] -and $policy.AnyOfRuleSets) {
                            foreach ($rs in @($policy.AnyOfRuleSets)) {
                                $exportRules = [System.Collections.Generic.List[object]]::new()
                                if ($rs.PSObject.Properties['AllOfRules'] -and $rs.AllOfRules) {
                                    foreach ($rule in @($rs.AllOfRules)) {
                                        $exportRules.Add((Export-PolicyRule $rule))
                                    }
                                }
                                $exportRuleSets.Add([PSCustomObject]@{ Rules = $exportRules.ToArray() })
                            }
                        }

                        # Serialize common rules
                        $exportCommon = [System.Collections.Generic.List[object]]::new()
                        if ($policy.PSObject.Properties['CommonRule'] -and $policy.CommonRule) {
                            foreach ($rule in @($policy.CommonRule)) {
                                $exportCommon.Add((Export-PolicyRule $rule))
                            }
                        }

                        $export = [PSCustomObject]@{
                            ExportedFrom   = $vCenterServer
                            ExportedAt     = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                            ScriptVersion  = $scriptVersion
                            PolicyName     = $policy.Name
                            Description    = if ($policy.PSObject.Properties['Description'] -and $policy.Description) { $policy.Description } else { '' }
                            RuleSetCount   = $exportRuleSets.Count
                            RuleSets       = $exportRuleSets.ToArray()
                            CommonRules    = $exportCommon.ToArray()
                        }

                        $safeName = $policy.Name -replace '[^\w\-]', '_'
                        $outFile  = Join-Path $outDir "$safeName.json"
                        $export | ConvertTo-Json -Depth 10 | Out-File -FilePath $outFile -Encoding UTF8
                        Write-Host "  [OK]   '$($policy.Name)' — $($exportRuleSets.Count) rule set(s) — exported to: $outFile" -ForegroundColor Green
                    } catch {
                        Write-Host "  [ERROR] Failed to export '$($policy.Name)': $_" -ForegroundColor Red
                    }
                }
            }
        }
    } catch {
        Write-Host "  [ERROR] Export failed: $_" -ForegroundColor Red
    }
}

# =============================================================================
# IMPORT
# =============================================================================
if ($Mode -eq 'Import') {
    try {
        # Resolve search directory / file list
        $searchDir = if ($FilePath -and (Test-Path $FilePath -PathType Container)) {
            $FilePath
        } elseif ($FilePath -and [System.IO.Path]::GetExtension($FilePath) -ne '') {
            $null   # specific file
        } else {
            $scriptDir
        }

        $filesToImport = [System.Collections.Generic.List[string]]::new()

        if ($null -eq $searchDir) {
            $filesToImport.Add($FilePath)
        } else {
            $jsonFiles = @(Get-ChildItem -Path $searchDir -Filter '*.json' -File | Sort-Object Name)
            if ($jsonFiles.Count -eq 0) {
                Write-Host "  [WARN] No JSON files found in $searchDir." -ForegroundColor Yellow
            } elseif ($jsonFiles.Count -eq 1) {
                Write-Host "  [INFO] Found 1 file: $($jsonFiles[0].Name)" -ForegroundColor Cyan
                $filesToImport.Add($jsonFiles[0].FullName)
            } else {
                Write-Host ""
                Write-Host "  JSON files in $searchDir`:" -ForegroundColor Cyan
                Write-Host ""
                for ($i = 0; $i -lt $jsonFiles.Count; $i++) {
                    try {
                        $peek    = Get-Content $jsonFiles[$i].FullName -Raw | ConvertFrom-Json
                        $polName = if ($peek.PSObject.Properties['PolicyName'])   { $peek.PolicyName }   else { '?' }
                        $rsCnt   = if ($peek.PSObject.Properties['RuleSetCount']) { $peek.RuleSetCount } else { '?' }
                        $line = "   [{0,2}]  {1,-35} policy: {2,-35} ({3} rule set(s))" -f ($i + 1), $jsonFiles[$i].Name, $polName, $rsCnt
                    } catch {
                        $line = "   [{0,2}]  {1}" -f ($i + 1), $jsonFiles[$i].Name
                    }
                    Write-Host $line -ForegroundColor White
                }
                Write-Host ""
                $selection = Read-Host "  Enter number(s) to import (comma-separated, or 'all')"
                $selection = $selection.Trim()

                if ($selection -ieq 'all') {
                    foreach ($f in $jsonFiles) { $filesToImport.Add($f.FullName) }
                } else {
                    foreach ($token in ($selection -split ',')) {
                        $token = $token.Trim()
                        $idx   = 0
                        if ([int]::TryParse($token, [ref]$idx) -and $idx -ge 1 -and $idx -le $jsonFiles.Count) {
                            $filesToImport.Add($jsonFiles[$idx - 1].FullName)
                        } else {
                            Write-Host "  [WARN] '$token' is not a valid selection — skipped." -ForegroundColor Yellow
                        }
                    }
                }
            }
        }

        if ($filesToImport.Count -eq 0) {
            Write-Host "  [WARN] No files selected." -ForegroundColor Yellow
        } else {
            # Pre-load all capabilities once for matching
            Write-Host "  [INFO] Loading capabilities from $vCenterServer..." -ForegroundColor Cyan
            $allCaps = @(Get-SpbmCapability -Server $viConn -ErrorAction SilentlyContinue)
            Write-Host "  [INFO] $($allCaps.Count) capability/capabilities loaded." -ForegroundColor Cyan
            Write-Host ""

            foreach ($file in $filesToImport) {
                try {
                    Write-Host "  [INFO] Reading $file..." -ForegroundColor Cyan
                    $import = Get-Content -Path $file -Raw -ErrorAction Stop | ConvertFrom-Json

                    $targetName = if ($NewPolicyName -and $filesToImport.Count -eq 1) { $NewPolicyName } else { $import.PolicyName }
                    Write-Host "  [INFO] Importing as '$targetName' ($($import.RuleSetCount) rule set(s), exported from $($import.ExportedFrom) on $($import.ExportedAt))." -ForegroundColor Cyan

                    # Check for existing policy
                    $existing = Get-SpbmStoragePolicy -Server $viConn -Name $targetName -ErrorAction SilentlyContinue
                    if ($existing) {
                        Write-Host "  [WARN] Policy '$targetName' already exists — skipped. Use -NewPolicyName to import under a different name." -ForegroundColor Yellow
                        continue
                    }

                    $ruleSets    = [System.Collections.Generic.List[object]]::new()
                    $commonRules = [System.Collections.Generic.List[object]]::new()

                    # Helper: build a SpbmRule from serialized data
                    $buildRule = {
                        param($ruleData)
                        if ($ruleData.Type -eq 'Capability') {
                            $cap = $allCaps | Where-Object { [string]$_.Id -eq $ruleData.CapabilityId } | Select-Object -First 1
                            if (-not $cap) {
                                Write-Host "         [WARN] Capability '$($ruleData.CapabilityId)' ($($ruleData.CapabilityName)) not found — rule skipped." -ForegroundColor Yellow
                                return $null
                            }
                            if ($ruleData.ValueType -eq 'Range') {
                                return New-SpbmRule -Capability $cap -Value (New-SpbmCapabilityConstraintRange -Minimum $ruleData.ValueMin -Maximum $ruleData.ValueMax) -ErrorAction Stop
                            }
                            $val = switch ($ruleData.ValueType) {
                                'Boolean' { [bool]::Parse($ruleData.Value) }
                                'Int32'   { [int]$ruleData.Value }
                                'Int64'   { [long]$ruleData.Value }
                                'Double'  { [double]$ruleData.Value }
                                default   { $ruleData.Value }
                            }
                            return New-SpbmRule -Capability $cap -Value $val -ErrorAction Stop
                        }
                        if ($ruleData.Type -eq 'Tag') {
                            $tags = [System.Collections.Generic.List[object]]::new()
                            foreach ($t in @($ruleData.Tags)) {
                                $tag = Get-Tag -Server $viConn -Name $t.Name -ErrorAction SilentlyContinue |
                                       Where-Object { $_.Category.Name -eq $t.Category } |
                                       Select-Object -First 1
                                if ($tag) { $tags.Add($tag) }
                                else { Write-Host "         [WARN] Tag '$($t.Category)/$($t.Name)' not found — skipped." -ForegroundColor Yellow }
                            }
                            if ($tags.Count -gt 0) {
                                return New-SpbmRule -AnyOfTags $tags.ToArray() -Server $viConn -ErrorAction Stop
                            }
                            return $null
                        }
                        Write-Host "         [WARN] Unknown rule type '$($ruleData.Type)' — skipped." -ForegroundColor Yellow
                        return $null
                    }

                    # Build rule sets
                    foreach ($rsData in @($import.RuleSets)) {
                        $rules = [System.Collections.Generic.List[object]]::new()
                        foreach ($ruleData in @($rsData.Rules)) {
                            $rule = & $buildRule $ruleData
                            if ($rule) { $rules.Add($rule) }
                        }
                        if ($rules.Count -gt 0) {
                            $ruleSets.Add((New-SpbmRuleSet -AllOfRules $rules.ToArray() -ErrorAction Stop))
                        }
                    }

                    # Build common rules
                    if ($import.PSObject.Properties['CommonRules'] -and $import.CommonRules) {
                        foreach ($ruleData in @($import.CommonRules)) {
                            $rule = & $buildRule $ruleData
                            if ($rule) { $commonRules.Add($rule) }
                        }
                    }

                    # Create the policy
                    $newPolParams = @{
                        Server      = $viConn
                        Name        = $targetName
                        Description = if ($import.PSObject.Properties['Description'] -and $import.Description) { $import.Description } else { '' }
                        ErrorAction = 'Stop'
                    }
                    if ($ruleSets.Count -gt 0)    { $newPolParams['AnyOfRuleSets'] = $ruleSets.ToArray() }
                    if ($commonRules.Count -gt 0)  { $newPolParams['CommonRule']   = $commonRules.ToArray() }

                    $null = New-SpbmStoragePolicy @newPolParams
                    Write-Host "  [OK]   '$targetName' created: $($ruleSets.Count) rule set(s) applied." -ForegroundColor Green
                } catch {
                    Write-Host "  [ERROR] Failed to import '$file': $_" -ForegroundColor Red
                }
            }
        }
    } catch {
        Write-Host "  [ERROR] Import failed: $_" -ForegroundColor Red
    }
}

# --- Disconnect ---------------------------------------------------------------
Disconnect-VIServer -Server $vCenterServer -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "  [INFO] Disconnected from $vCenterServer." -ForegroundColor Cyan
Write-Host ('=' * 62) -ForegroundColor DarkCyan
