#Requires -version 2.0
#
# version : '0.4'
# 2012/9/25 下午 06:01:18
#
# Displayname Rule
# 1. 有 firstname & lastname and mail 含 @umc.com and 中文姓名 : 姓名(英文)
# 2. 有 firstname & lastname and mail 含 @umc.com and 英文姓名 : 名 姓
# 3. 有 firstname & lastname and 無 mail and 中文姓名 : 姓名
# 4. 有 firstname & lastname and 無 mail and 英文姓名 : 名 姓
# 5. 無 firstname : 工號
#
# Download QAD http://www.quest.com/powershell/activeroles-server.aspx
#
#
# Exchange	UMCT136
# Lync			UTTRA027
# 
# Execute on UMCT136
#
# Set-ExecutionPolicy "RemoteSigned"
# Need to “Unblock” the script
# For Exchange 2010:
# Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
# For Exchange 2007:
# Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
#
# powershell -command "& {.\AccountSync.ps1}"

set-psdebug -strict

$ExchangeS = 'UMCT136'
$LyncS = 'UTTRA027'

$startT = get-date

if (Test-Path "$pwd\UDF.ps1") {
   . .\UDF.ps1
   Write-Output "UDF loaded"
} else {
	 Write-Output "Missing file UDF.ps1"
	 exit -250
}

if ( (Get-PSSnapin -Name quest.activeroles.admanagement -ErrorAction SilentlyContinue) -eq $null ) 
{ 
   add-pssnapin quest.activeroles.admanagement
   #Set-QADProgressPolicy -ShowProgress $false | out-null
   Write-Output "QAD loaded"
} 

if ($Env:COMPUTERNAME -eq $ExchangeS) {
   Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
}

#if ($Env:COMPUTERNAME -eq $LyncS) {
	 #Import the Lync modules
   #Load-Module -Name "C:\Program Files\Common Files\Microsoft Lync Server 2010\Modules\Lync\Lync.psd1" -Para "-PassThru"
   Import-Module Lync -PassThru
   Write-Output "Lync loaded"
   $env:ADPS_LoadDefaultDrive = 0
   #Load-Module -Name ActiveDirectory -Para "-Cmdlet Get-ADUser,Set-ADUser -PassThru"
   Import-Module -Name ActiveDirectory -Cmdlet Get-ADUser,Set-ADUser -PassThru
   Write-Output "ActiveDirectory loaded"
#}


$pwd = Get-Location

#set-location $pwd
$logd = "\\$ExchangeS\c$\SOURCE\test\AccountSync.log"
$llog = "$pwd\AccountSync-LY.log"

if (Test-Path "$pwd\Script4logging.ps1") {
   . .\Script4logging.ps1
   $LogFile = "$pwd\AccountSync-LY.log"
   Write-Log "=== Job Begin ==="
} else {
	 Write-Output "Missing file Script4logging.ps1"
	 exit -220
}


$NewAdd = Import-CSV ((get-childitem "\\$ExchangeS\c$\SOURCE\test" NewAdd.*.csv | sort lastwritetime -desc)[0]).versioninfo.filename
#$outfile = $("\\$ExchangeS\c$\SOURCE\test\ASync.AD.{0}.csv" -f (get-date -format "yyyyMMdd-HHmmss"))
#New-Variable name
#$result = @()
#$count = 0

if ($NewAdd.count -ne $null) {

foreach ($nu in $NewAdd) {
# Write-Log "Returns information about all the user accounts in Active Directory Domain Services (AD DS)"
#$Lfilter = "(&(!enabled=True)(!name=IUSR_*)(!name=IWAM_*)(!name=SystemMailbox*)(!name=*$))"
#$xuser = Get-CsAdUser -Filter {Enabled -ne $True -and Name -ne 'guest' -and Name -notlike 'SystemMailbox*' -and Name -notlike '*$' -and Name -ne 'krbtgt' -and Name -notlike 'DiscoverySearchMailbox*'}
#Write-Log "All user information collected"

#foreach ($xu in $xuser) { Write-Log $xu.name }
   Write-Log ("Process " + $nu.name)
   Get-CsAdUser -Identity $nu.name | Enable-CsUser -RegistrarPool "pools.uttra.local" -SipAddressType EmailAddress
   Write-Log ("User {0} enabled" -f $nu.name)
   #$xuser | Set-ADUser -Replace @{"msRTCSIP-OriginatorSid"=$_.msExchMasterAccountSid.Value.ToString()}
   $sidstr = (Get-AdUser -Identity $nu.name -Properties msExchMasterAccountSid | %{$_.msExchMasterAccountSid})
   Write-Log ("get sid [{0}]" -f $sidstr.tostring())
   #Set-QADUser -UseDefaultExcludedProperties $true -ObjectAttributes @{'msRTCSIP-OriginatorSID'=$_.msExchMasterAccountSid.Value.ToString()}
   Set-ADUser -Identity $nu.name -Replace @{"msRTCSIP-OriginatorSid"=$sidstr.tostring()}
   $sidstr = (Get-AdUser -Identity $nu.name -Properties "msRTCSIP-OriginatorSid" | %{$_."msRTCSIP-OriginatorSid"})
   Write-Log ("User {0} SIDs copied [{1}]" -f $nu.name,$sidstr.tostring())
   #Get-CsAdUser -OU "OU=Lync,DC=uttra,DC=local" -Identify "$a.Name" | Enable-CsUser -RegistrarPool "pools.uttra.local" -SipAddressType EmailAddress
   #Get-CsAdUser -Identify "$a.Name" | Set-ADUser -Replace @{"msRTCSIP-OriginatorSid"=$_.msExchMasterAccountSid.Value.ToString()}
}
}

#Get-CsAdUser | Export-CSV $outfile -encoding "unicode"

$endT = get-date
$ts = $endT - $startT
Write-Output $("Process time [{0:00}:{1:00}:{2:00}]" -f $ts.Hours,$ts.Minutes,$ts.Seconds)
Write-Log $("Process time [{0:00}:{1:00}:{2:00}]" -f $ts.Hours,$ts.Minutes,$ts.Seconds)
Write-Log "--- Job End ---"
