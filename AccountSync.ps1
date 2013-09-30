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
# Exchange	UMCT136
# Lync			UTTRA027
#
# Execute on UTTRA027 D:\SOURCE\Job
#
# Step A
# Download QAD http://www.quest.com/powershell/activeroles-server.aspx & install
#
# Enable Powershell to run script:
# Set-ExecutionPolicy "RemoteSigned"
#
# Create Credential
# powershell -command "& {.\Set-LyncCredential.ps1 lpoc.txt}"
# powershell -command "& {.\Set-LyncCredential.ps1 oaadmin.txt}"
#
# For Exchange 2010:
# Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
# For Exchange 2007:
# Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
# For Lync:
# Import-Module Lync
# For QAD:
# Add-pssnapin quest.activeroles.admanagement
#
# ## Query for SID not the same
# Get-AdUser -ldapfilter "(&(!name=IUSR_*)(!name=IWAM_*)(!name=SystemMailbox*)(!name=*$)(!name=DiscoverySearchMailbox*)(!name=FederatedEmail*)(!name=krbtgt))" -ResultsetSize $null -Properties msExchMasterAccountSid,"msRTCSIP-OriginatorSid" | where {$_."msRTCSIP-OriginatorSid" -ne $_.msExchMasterAccountSid } | select name
#
# ## Query for not enabled
# Get-CsAdUser -Filter {Enabled -ne $True} | Where {$_.Name -NotLike "IUSR_*" -and $_.Name -NotLike "IWAM_*" -and $_.Name -NotLike "SystemMailbox*" -and $_.Name -NotLike "*$" -and $_.Name -NotLike "DiscoverySearchMailbox*" -and $_.Name -NotLike "FederatedEmail*" -and $_.Name -NotLike "krbtgt*" -and $_.mail -contains "@"} | Enable-CsUser -RegistrarPool "pools.uttra.local" -SipAddressType EmailAddress
# 
# $likeArray = @("IUSR_.*","IWAM_.*","SystemMailbox.*",".*\$","DiscoverySearchMailbox.*","FederatedEmail.*","krbtgt.*")
# [regex] $a_regex = '(?i)^(' + (($likeArray |foreach {$_}) –join "|") + ')$'
# Get-CsAdUser -Filter {Enabled -ne $True} | Where {$_Name -notmatch $a_regex -and $_.mail -contains "@"}
#
# Run this script:
# powershell -command "& {.\AccountSync-Remoting.ps1}"

set-psdebug -strict

$ExchangeS = 'UMCT162'
$LyncS = 'UMCT148'

$startT = get-date

if ( (Get-PSSnapin -Name quest.activeroles.admanagement -ErrorAction SilentlyContinue) -eq $null )
{
   add-pssnapin quest.activeroles.admanagement
   #Set-QADProgressPolicy -ShowProgress $false | out-null
   Write-Output "QAD loaded"
}

#if ($Env:COMPUTERNAME -eq $ExchangeS) {
#   Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
#}

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

if (Test-Path "$pwd\Script4logging.ps1") {
   . .\Script4logging.ps1
   $LogFile = "$pwd\AccountSync-Remoting.log"
   Write-Log "=== Job Begin ==="
} else {
	 Write-Output "Missing file Script4logging.ps1"
	 exit -220
}

$secfile = "$pwd\lpoc.txt"
$ACC = 'lpoc\administrator'

if (Test-Path "$pwd\Get-LyncCredential.ps1") {
   if (Test-Path $secfile) {
   	 . .\Get-LyncCredential.ps1 $ACC $secfile
   } else {
   	 Write-Log "Missing file $secfile"
   	 exit -200
   }
} else {
	 Write-Log "Missing file Get-LyncCredential.ps"
	 exit -210
}

$ExCred = $credential

$secfile = "$pwd\oaadmin.txt"
$ACC = 'umc\oaadmin'

if (Test-Path "$pwd\Get-LyncCredential.ps1") {
   if (Test-Path $secfile) {
   	 . .\Get-LyncCredential.ps1 $ACC $secfile
   } else {
   	 Write-Log "Missing file $secfile"
   	 exit -200
   }
} else {
	 Write-Log "Missing file Get-LyncCredential.ps"
	 exit -210
}

Write-Log "Prepare to remoting to $ExchangeS"
$session = New-PSSession -Configurationname Microsoft.Exchange –ConnectionUri "http://$ExchangeS/powershell" -Credential $ExCred -ErrorAction Stop

Import-PSSession $session
Write-Log "Session created"

##################################################################################

##################################################################################

$outfile = $("$pwd\ASync.{0}.csv" -f (get-date -format "yyyyMMdd-HHmmss"))
New-Variable name
$result = @()
$count = 0
$keeplog = 30

$plist = Get-QADPSSnapinSettings -DefaultOutputPropertiesForUserObject
Set-QADPSSnapinSettings -DefaultOutputPropertiesForUserObject @("name","mail","firstname","lastname")

$startT = get-date

$Lfilter = "(&(!name=IUSR_*)(!name=IWAM_*)(!name=SystemMailbox*)(!name=*$)(!name=DiscoverySearchMailbox*)(!name=FederatedEmail*)(!name=krbtgt))"

$conn = Connect-QADService umc.com -Credential $Credential
Write-Log "Collect all AD users"
$acc = Get-QADUser -SearchRoot "CN=users,DC=umc,DC=com" -sizelimit 0 -LDAPFilter $Lfilter -Connection $conn | Sort-Object Name
Set-QADPSSnapinSettings -DefaultOutputPropertiesForUserObject $plist

Write-Log "Displayname Composition"
foreach ($a in $acc) {
   $count++
   if ($count % 2000 -eq 0) { Write-Log ("{0,5} accounts displayname processed" -f $count) }
   $tObj = New-Object PSObject

   # 若 firstname, lastname and mail 皆有值
   if ($a.mail -like '*@umc.com' -and $a.firstname -ne $null -and $a.lastname -ne $null) {
      Write-Output $("[{0}] [{1}] [{2}]" -f $a.firstname, $a.lastname, $a.mail)

      # 若 firstname/lastname 組合與 mail 英文名稱相同
      if ($("{0} {1}" -f $a.firstname, $a.lastname) -eq `
      	     (Get-Culture).textinfo.totitlecase($($a.mail.replace("@umc.com","")).replace("_", " ")) -or $a.mail -eq '@umc.com') {
   	     # 是英文就分開，姓在後。是中文就合起來，姓在前
   	     if ($a.firstname -match '^[A-Z]') {
      	    $name = "{0} {1}" -f $a.firstname, $a.lastname
      	 } else {
      	    $name = "{1}{0}" -f $a.firstname, $a.lastname
      	 }
   	     $tObj | Add-Member NoteProperty Alias $a.Name
   	     $tObj | Add-Member NoteProperty Name $a.Name
   	     $tObj | Add-Member NoteProperty LMAccount $("UMC\" + $a.Name)
   	     $tObj | Add-Member NoteProperty DisplayName $name
   	     $tObj | Add-Member NoteProperty UPN $($a.Name + "@uttra.local")
      } else {
   	     # 無中文姓名
   	     if ($a.firstname -match '^[A-Z]') {
   	     	  $name = (Get-Culture).textinfo.totitlecase($($a.mail.replace("@umc.com","")).replace("_", " "))
   	     } else {
   	     	  $name = "{1}{0}({2})" -f $a.firstname, $a.lastname, (Get-Culture).textinfo.totitlecase($($a.mail.replace("@umc.com","")).replace("_", " "))
   	     }
   	     $tObj | Add-Member NoteProperty Alias $a.Name
   	     $tObj | Add-Member NoteProperty Name $a.Name
   	     $tObj | Add-Member NoteProperty LMAccount $("UMC\" + $a.Name)
   	     $tObj | Add-Member NoteProperty DisplayName $name
   	     $tObj | Add-Member NoteProperty UPN $($a.Name + "@uttra.local")
      }
   } else {
   	     $tObj | Add-Member NoteProperty Alias $a.Name
   	     $tObj | Add-Member NoteProperty Name $a.Name
   	     $tObj | Add-Member NoteProperty LMAccount $("UMC\" + $a.Name)
   	     $tObj | Add-Member NoteProperty DisplayName $a.Name
   	     $tObj | Add-Member NoteProperty UPN $($a.Name + "@uttra.local")
   }
   $result += $tObj
}

$result | Export-CSV -Path "$outfile" -encoding unicode -notype -ErrorAction SilentlyContinue

Write-Log "Export $count accounts to csv $outfile"

$endT = get-date
$ts = $endT - $startT
Write-Log $("Process time [{0:00}:{1:00}:{2:00}]" -f $ts.Hours,$ts.Minutes,$ts.Seconds)

$R = @{}
foreach ($a in $result) {
	 $h = @{"Alias"=$a.Alias;"LMAccount"=$a.LMAccount;"DisplayName"=$a.DisplayName;"UPN"=$a.UPN}
	 $R.add($a.Name, $h)
}

$endT = get-date
$ts = $endT - $startT
Write-Log $("Make Hash time [{0:00}:{1:00}:{2:00}]" -f $ts.Hours,$ts.Minutes,$ts.Seconds)

#$Lfilter = "<LDAP://ou=Lync,dc=uttra,dc=local>;" +
#           "(&(objectCategory=person)(objectClass=user)(msRTCSIP-UserEnabled=TRUE)(msExchMasterAccountSid=*)(userAccountControl:1.2.840.113556.1.4.803:=2));" +
#           "ADsPath,cn,msRTCSIP-PrimaryUserAddress,msExchMasterAccountSid,msRTCSIP-OriginatorSid;subtree"

$count = 0
$newacc = 0

##$mailbox = @{}
##$allmailbox = Get-Mailbox -ResultSize Unlimited
##foreach ($m in $allmailbox) {
##	 $u = @{"Name"=$m.samaccountname;"Displayname"=$m.displayname}
##	 $mailbox.add($m.Name, $u)
##}

######
# Add missing account
######
foreach ($aUser in $result) {
	 $count++
	 if ($count % 2000 -eq 0) { Write-Log ("{0,5} accounts existence checked" -f $count)}
	 # check mailbox existence
	 if ([bool](Get-Mailbox -Identity $aUser.UPN -ErrorAction SilentlyContinue)) {
	 ##if ($mailbox.ContainsKey($aUser.name)) {
	    # check display name correctness
	    $Mbox = Get-Mailbox -Identity $aUser.UPN
	    if ($aUser.DisplayName -ne $Mbox.DisplayName) {
	    ##if ($aUser.DisplayName -ne $mailbox.item($aUser.name).item("Displayname")) {
	    	 Write-Log ("{0} DisplayName updated. AD({1}) Lync({2})" -f $aUser.Name, $aUser.DisplayName)
	    	 ##Write-Log ("{0} DisplayName updated. AD({1}) Lync({2})" -f $aUser.Name, $mailbox.item($aUser.name).item("Displayname"))
	    	 Set-Mailbox -Identity $aUser.UPN -DisplayName $aUser.DisplayName
	    }
	 } else {
	 	  # add new mailbox
	 	  $NName = $aUser.Name
	 	  Write-Log ("Create New mailbox " + $NName )
	 	  $nbox = New-Mailbox -Alias $aUser.Alias -Name $NName  -UserPrincipalName $aUser.UPN -LinkedDomainController "UMCZ03.umc.com" -LinkedMasterAccount $aUser.LMAccount -DisplayName $aUser.DisplayName -OrganizationalUnit Lync -LinkedCredential $credential
	 	  if ([bool]$nbox) {
	 	  	 Write-Log ("`tNew mailbox " + $NName  + " created")
	 	  } else {	 
	 	  #if (!([bool](Get-Mailbox -Identity $aUser.UPN -ErrorAction SilentlyContinue))) {
	 	  	 Write-Log ("Fail to create mailbox " + $aUser.UPN)
	 	  }
      #Write-Log ("`tProcess " + $aUser.name)
      # 對 Lync Server 建立帳號
      Get-CsAdUser -Identity $NName  | Enable-CsUser -RegistrarPool "poc.umc.com" -SipAddressType EmailAddress
      # 對 Lync Server 上的帳號進行設定
      Write-Log ("`tUser {0} enabled" -f $NName )
      
      # 對 Exchange Server 上的 AD DS 帳號進行設定
      $sidstr = (Get-ADUser -Identity $NName  -Properties msExchMasterAccountSid | %{$_.msExchMasterAccountSid})
      Write-Log ("`tget sid [{0}]" -f $sidstr.tostring())
      Set-ADUser -Identity $NName  -Replace @{"msRTCSIP-OriginatorSid"=$sidstr.tostring()}
      $sidstr = (Get-AdUser -Identity $NName  -Properties "msRTCSIP-OriginatorSid" | %{$_."msRTCSIP-OriginatorSid"})
      Write-Log ("`tUser {0} SIDs [{1}] copied" -f $NName ,$sidstr.tostring())
      
      # 對 Exchange Server 上的帳號進行設定
      Set-Mailbox -Identity $aUser.UPN -DisplayName $aUser.DisplayName
      
      # 對 Lync Server 上的帳號進行最後設定
      Get-csaduser -Identity $NName | Set-csuser -SipAddress ('sip:' + $aUser.name + '@umc.com')
      $newacc++
	 }
}
Write-Log "$newacc new accounts created"


$count = 0
######
# Remove extra account
######
$connL = Connect-QADService uttra.local -Credential $ExCred
Set-QADPSSnapinSettings -DefaultOutputPropertiesForUserObject @("name","mail","firstname","lastname")
$accs = Get-QADUser -SizeLimit 0 -Connection $connL -LDAPFilter $Lfilter
Set-QADPSSnapinSettings -DefaultOutputPropertiesForUserObject $plist
foreach ($a in $accs) {
	 #$j = $result | where {$_.name -eq $a.name}
	 if (! $R.ContainsKey($a.name)) {
	 	  $count++
	 	  Remove-Mailbox -Identity $a.name -Confirm:$False
	 	  Write-Log ("Extra account " + $a.name + " removed")
	 }
}
Write-Log "$count accounts removed"

$oldfiles = (Get-childitem "$pwd" ASync.*.csv | where {$_.lastwritetime -lt (Get-Date).AddDays(-$keeplog)} | %{$_.versioninfo.filename})
if ($oldfiles.count -ne $null) {
   foreach ($ff in $oldfiles) {
      Remove-Item $ff
      Write-Log ("{0} removed" -f $ff)
   }
}

Remove-PsSession $session
Write-Log "Session removed"

$endT = get-date
$ts = $endT - $startT
Write-Output $("Process time [{0:00}:{1:00}:{2:00}]" -f $ts.Hours,$ts.Minutes,$ts.Seconds)
Write-Log $("Process time [{0:00}:{1:00}:{2:00}]" -f $ts.Hours,$ts.Minutes,$ts.Seconds)
Write-Log "--- Job End ---"
