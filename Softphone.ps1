<#
Softphone User voice policy change
Version 0.1
OCWS
#>
<#
param (
	[Parameter(Position=0,
				HelpMessage="Path to input CSV file")]
	[alias("CSVPath")]
	[String]$InputCSVPath="C:\Users\NXLX8474\Desktop\users.csv"
)
#>

#User CSV loading
$usersList = $null
$usersList = Import-CSV C:\Sources\scriptsoftphone\SoftphoneUsers.csv
$count = $usersList.count

write-host "User count within CSV file=" $count
write-host ""

ForEach($user in $usersList) 
{

    if($user.VoicePolicy -eq "" )
        {
        write-host "CSV File doesn't contain any Voice policy data"
        write-host "If clause"
        $SipAddressURI="sip:"+ $user.SipAddress
        $CurrentUserObject = get-csuser -identity $SipAddressURI
        $CurrentVoicePolicy = ($CurrentUserObject).VoicePolicy

        if($CurrentVoicePolicy.friendlyname -eq "VP-Full-Access")
            {
            
            Grant-csvoicepolicy -Identity $user.upn -PolicyName VP-BTIP-Full-Access
            if($? -eq $true)
                {
		        write-host "VP-Full-Access switched to VP-BTIP-Full-Access for user" $user.upn
	            }
	        }
        }
        elseif($CurrentVoicePolicy.friendlyname -eq "VP-Phone-Limited")
              {            
              Grant-csvoicepolicy -Identity $user.upn -PolicyName VP-BTIP-Phone-Limited
              if($? -eq $true)
                 {
		         write-host "VP-Phone-Limited switched to VP-BTIP-Phone-Limited for user" $user.upn
	             }
              }
        elseif($CurrentVoicePolicy.friendlyname -eq "VP-National-Limited")
              {
              Grant-csvoicepolicy -Identity $user.upn -PolicyName VP-BTIP-National-Limited
              if($? -eq $true)
                {
		        write-host "VP-National-Limited switched to VP-BTIP-National-Limited for user" $user.upn
	            }
              }
        elseif($CurrentVoicePolicy.friendlyname -eq "VP-National-CallCharge")
              {
              Grant-csvoicepolicy -Identity $user.upn -PolicyName VP-BTIP-National-CallCharge
              if($? -eq $true)
                {
		        write-host "VP-National-CallCharge switched to VP-BTIP-National-CallCharge for user" $user.upn
	            }
              }
        elseif($CurrentVoicePolicy.friendlyname -eq "VP-National")
              {
              Grant-csvoicepolicy -Identity $user.upn -PolicyName VP-BTIP-National
              if($? -eq $true)
                {
		        write-host "VP-National switched to VP-BTIP-National for user" $user.upn
	            }
              }
        elseif($CurrentVoicePolicy.friendlyname -eq "VP-International")
              {
              Grant-csvoicepolicy -Identity $user.upn -PolicyName VP-BTIP-International
              if($? -eq $true)
                {
		        write-host "VP-International switched to VP-BTIP-International for user" $user.upn
	            }
              }

     elseif($user.VoicePolicy -ne "" )
           {
           write-host "CSV File contains Voice policy data"
           Grant-CsVoicePolicy -Identity $user.upn -PolicyName $user.VoicePolicy
	       if($? -eq $true)
             {
		     write-host -NoNewLine "Voice policy "; Write-Host -NoNewline -ForegroundColor Green $user.VoicePolicy; Write-Host -NoNewline " granted to useruser: "; Write-Host -ForegroundColor Cyan $user.upn
	         }
	       }
    write-host "****************************User " -nonewline; Write-Host -foregroundcolor Cyan $user.upn -nonewline; Write-Host  " set****************************"
    write-host ""
}	