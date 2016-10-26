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
$usersList = Import-CSV C:\Sources\SoftphoneUsers.csv
$count = $usersList.count

write-host "User count within CSV file=" $count
write-host ""

ForEach ($user in $usersList) 
{

    if ($user.VoicePolicy -eq $null )
        {
        $CurrentUserObject = get-csuser -identity $user.upn
        $CurrentVoicePolicy = ($CurrentUserObject).VoicePolicy

        if($CurrentVoicePolicy.friendlyname -eq "VP-Full-Access")
            {
            Grant-csvoicepolicy -Identity $user.upn -PolicyName VP-BTIP-Full-Access 
            }
        elseif($CurrentVoicePolicy.friendlyname -eq "VP-Phone-Limited")
            {
            Grant-csvoicepolicy -Identity $user.upn -PolicyName VP-BTIP-Phone-Limited
            }
        elseif($CurrentVoicePolicy.friendlyname -eq "VP-National-Limited")
            {
            Grant-csvoicepolicy -Identity $user.upn -PolicyName VP-BTIP-National-Limited
            }
        elseif($CurrentVoicePolicy.friendlyname -eq "VP-National-CallCharge")
            {
            Grant-csvoicepolicy -Identity $user.upn -PolicyName VP-BTIP-National-CallCharge
            }
        elseif($CurrentVoicePolicy.friendlyname -eq "VP-National")
            {
            Grant-csvoicepolicy -Identity $user.upn -PolicyName VP-BTIP-National 
            }
        elseif($CurrentVoicePolicy.friendlyname -eq "VP-International")
            {
            Grant-csvoicepolicy -Identity $user.upn -PolicyName VP-BTIP-International
            }
        }
	
    elseif ($user.VoicePolicy -ne $null )
        {
        Grant-CsVoicePolicy -Identity $user.upn -PolicyName $user.VoicePolicy
	    if ($? -eq $true)
            {
		    write-host -NoNewLine "Voice policy "; Write-Host -NoNewline -ForegroundColor Green $user.VoicePolicy; Write-Host -NoNewline " granted to useruser: "; Write-Host -ForegroundColor Cyan $user.upn
	        }
	    }

	write-host "****************************User " -nonewline; Write-Host -foregroundcolor Cyan $user.upn -nonewline; Write-Host  " set****************************"
	write-host ""
}	