<#
Softphone User voice policy change
Version 0.1
OCWS
#>

param([string]$userCsv = "C:\Sources\users.csv")

#User CSV loading
$usersList = $null
$usersList = Import-Csv $userCsv
$count = $usersList.count
Write-Host "User count within CSV file=" $count
Write-Host ""

foreach ($user in $usersList)
{

  if ($user.VoicePolicy -eq "")
  {
    Write-Host "CSV File doesn't contain any Voice policy data"
    Write-Host "If clause"
    $SipAddressURI = "sip:" + $user.SipAddress
    $CurrentUserObject = Get-CsUser -identity $SipAddressURI
    $CurrentVoicePolicy = ($CurrentUserObject).VoicePolicy
    if ($CurrentVoicePolicy.friendlyname -eq "VP-Full-Access")
    {
      Grant-CsVoicePolicy -identity $user.upn -PolicyName VP-BTIP-Full-Access
      if ($? -eq $true)
      {
        Write-Host "VP-Full-Access switched to VP-BTIP-Full-Access for user" $user.upn
      }
    }
    elseif ($CurrentVoicePolicy.friendlyname -eq "VP-Phone-Limited")
    {
      Grant-CsVoicePolicy -identity $user.upn -PolicyName VP-BTIP-Phone-Limited
      if ($? -eq $true)
      {
        Write-Host "VP-Phone-Limited switched to VP-BTIP-Phone-Limited for user" $user.upn
      }
    }
    elseif ($CurrentVoicePolicy.friendlyname -eq "VP-National-Limited")
    {
      Grant-CsVoicePolicy -identity $user.upn -PolicyName VP-BTIP-National-Limited
      if ($? -eq $true)
      {
        Write-Host "VP-National-Limited switched to VP-BTIP-National-Limited for user" $user.upn
      }
    }
    elseif ($CurrentVoicePolicy.friendlyname -eq "VP-National-CallCharge")
    {
      Grant-CsVoicePolicy -identity $user.upn -PolicyName VP-BTIP-National-CallCharge
      if ($? -eq $true)
      {
        Write-Host "VP-National-CallCharge switched to VP-BTIP-National-CallCharge for user" $user.upn
      }
    }
    elseif ($CurrentVoicePolicy.friendlyname -eq "VP-National")
    {
      Grant-CsVoicePolicy -identity $user.upn -PolicyName VP-BTIP-National
      if ($? -eq $true)
      {
        Write-Host "VP-National switched to VP-BTIP-National for user" $user.upn
      }
    }
    elseif ($CurrentVoicePolicy.friendlyname -eq "VP-International")
    {
      Grant-CsVoicePolicy -identity $user.upn -PolicyName VP-BTIP-International
      if ($? -eq $true)
      {
        Write-Host "VP-International switched to VP-BTIP-International for user" $user.upn
      }
    }
  }
  elseif ($user.VoicePolicy -ne "")
  {
    Write-Host "CSV File contains Voice policy data"
    Grant-CsVoicePolicy -identity $user.upn -PolicyName $user.VoicePolicy
    if ($? -eq $true)
    {
      Write-Host -NoNewline "Voice policy "; Write-Host -NoNewline -ForegroundColor Green $user.VoicePolicy; Write-Host -NoNewline " granted to useruser: "; Write-Host -ForegroundColor Cyan $user.upn
    }
  }
  Write-Host "****************************User " -NoNewline; Write-Host -ForegroundColor Cyan $user.upn -NoNewline; Write-Host " set****************************"
  Write-Host ""
}
