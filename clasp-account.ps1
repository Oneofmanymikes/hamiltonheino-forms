<#  Swap the single global clasp credential between campaigns.

    clasp keeps ONE credential at ~/.clasprc.json, so whichever account logged
    in last wins — and a session working on the other campaign silently steals
    it. This saves/restores named copies instead of re-doing the browser login.

    Usage:
      .\clasp-account.ps1 save heino      # after logging in as info@voteheino.ca
      .\clasp-account.ps1 use  heino      # switch back to that saved credential
      .\clasp-account.ps1 who             # which account is active right now
#>
param(
  [Parameter(Mandatory=$true)][ValidateSet('save','use','who','list')][string]$Action,
  [string]$Name
)
$rc = "$env:USERPROFILE\.clasprc.json"

switch ($Action) {
  'who'  { clasp show-authorized-user }
  'list' { Get-ChildItem "$env:USERPROFILE\.clasprc.json.*" |
             ForEach-Object { "{0,-12} {1}" -f ($_.Name -replace '^\.clasprc\.json\.',''), $_.LastWriteTime } }
  'save' {
    if (-not $Name) { throw "Give it a name, e.g. save heino" }
    Copy-Item $rc "$rc.$Name" -Force
    "Saved the ACTIVE credential as '$Name'."
    clasp show-authorized-user
  }
  'use' {
    if (-not $Name) { throw "Give it a name, e.g. use heino" }
    $src = "$rc.$Name"
    if (-not (Test-Path $src)) { throw "No saved credential called '$Name'. Run: .\clasp-account.ps1 list" }
    Copy-Item $src $rc -Force
    "Switched to '$Name'."
    clasp show-authorized-user
  }
}