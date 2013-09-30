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

$keeplog = 30

add-PSSnapin quest.activeroles.admanagement
#Set-QADProgressPolicy -ShowProgress $false | out-null

if ($Env:COMPUTERNAME -eq $ExchangeS) {
   Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
}

if ($Env:COMPUTERNAME -eq $LyncS) {
	 #Import the Lync modules
   Import-Module "C:\Program Files\Common Files\Microsoft Lync Server 2010\Modules\Lync\Lync.psd1"
}

$pwd = Get-Location

#set-location $pwd

if (Test-Path "$pwd\Script4logging.ps1") {
   . .\Script4logging.ps1
   $LogFile = "$pwd\AccountSync-EX.log"
   Write-Log "=== Job Begin ==="
} else {
	 Write-Output "Missing file Script4logging.ps1"
	 exit -220
}

$secfile = "$pwd\uttra.txt"
$ACC = 'uttra\administrator'

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

$LyncCred = $credential

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

$outfile = $("$pwd\ASync.{0}.csv" -f (get-date -format "yyyyMMdd-HHmmss"))
New-Variable name
$result = @()
$count = 0

$startT = get-date

#$acc = Get-QADUser -SearchRoot "CN=00002940,CN=users,DC=umc,DC=com" -sizelimit 0 `
#         | Where {$_.Name -notlike '*\$' -and $_.Name -notlike "IUSR_*" -and $_.Name -notlike "IWAM_*" `
#         	         -and $_.Name -notlike "SystemMailbox*"} | Sort-Object Name

#Get-QADUser -SearchRoot "CN=users,DC=umc,DC=com" -sizelimit 0 -LDAPFilter "(&(cn='00002*')(!name=IUSR_*)(!name=IWAM_*)(!name=SystemMailbox*)(!name=*$))" `
#         | Where {$_.Name -notlike '*\$' -and $_.Name -notlike "IUSR_*" -and $_.Name -notlike "IWAM_*" `
#         	         -and $_.Name -notlike "SystemMailbox*"} | Sort-Object Name | Foreach-Object {

$Lfilter = "(&(!name=IUSR_*)(!name=IWAM_*)(!name=SystemMailbox*)(!name=*$)(!name=DiscoverySearchMailbox*)(!name=FederatedEmail*))"
#$acc = Get-QADUser -SearchRoot "CN=users,DC=umc,DC=com" -sizelimit 0 -LDAPFilter $Lfilter | Sort-Object Name #| Foreach-Object {
$conn = Connect-QADService umc.com -Credential $Credential
Write-Log "Collect all AD users"
$acc = Get-QADUser -SearchRoot "CN=users,DC=umc,DC=com" -sizelimit 0 -LDAPFilter $Lfilter -Connection $conn | Sort-Object Name

Write-Log "Displayname Composition"
foreach ($a in $acc) {
   $count++
   if ($count % 2000 -eq 0) { Write-Log "$count accounts displayname processed" }
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

#$R


#$session = New-PSSession -ConnectionUri "https://$LyncS.uttra.local/OcsPowershell" -Credential $LyncCred
#Import-PsSession $session

#$Lfilter = "<LDAP://ou=Lync,dc=uttra,dc=local>;" +
#           "(&(objectCategory=person)(objectClass=user)(msRTCSIP-UserEnabled=TRUE)(msExchMasterAccountSid=*)(userAccountControl:1.2.840.113556.1.4.803:=2));" +
#           "ADsPath,cn,msRTCSIP-PrimaryUserAddress,msExchMasterAccountSid,msRTCSIP-OriginatorSid;subtree"

$acca = @('Name')
$count = 0

######
# Add missing account
######
foreach ($a in $result) {
	 $count++
	 if ($count % 2000 -eq 0) { Write-Log "$count accounts existence checked" }
	 # check mailbox existence
	 if ([bool](Get-Mailbox -Identity $a.UPN -ErrorAction SilentlyContinue)) {
	    # check display name correctness
	    $Mbox = Get-Mailbox -Identity $a.UPN
	    if ($a.DisplayName -ne $Mbox.DisplayName) {
	    	 Write-Log ("{0} DisplayName updated. AD({1}) Lync({2})" -f $a.Name, $a.DisplayName, $Mbox.DisplayName)
	    	 Set-Mailbox -Identity $a.UPN -DisplayName $a.DisplayName
	    }
	 } else {
	 	  # add new mailbox
	 	  Write-Log ("Create New mailbox " + $a.Name)
	 	  New-Mailbox -Alias $a.Alias -Name $a.Name -UserPrincipalName $a.UPN -LinkedDomainController "UMCZ03.umc.com" -LinkedMasterAccount $a.LMAccount -OrganizationalUnit Lync -LinkedCredential $credential -DisplayName $a.DisplayName
	 	  $acca += $a.Name
	 	  Write-Log ("New mailbox " + $a.Name + " created")
	 	  if (!([bool](Get-Mailbox -Identity $a.UPN -ErrorAction SilentlyContinue))) {
	 	  	 Write-Log ("Fail to create mailbox " + $a.UPN)
	 	  }
     #Get-CsAdUser -OU "OU=Lync,DC=uttra,DC=local" -Identify "$a.Name" | Enable-CsUser -RegistrarPool "pools.uttra.local" -SipAddressType EmailAddress
     #Get-CsAdUser -Identify "$a.Name" | Set-ADUser -Replace @{"msRTCSIP-OriginatorSid"=$_.msExchMasterAccountSid.Value.ToString()}
	 	 #Write-Log ("New mailbox " + $a.Name + " synced")
	 }
}
$acca

$Naccf = "$pwd\NewAdd.{0}.csv" -f (get-date -format "yyyyMMdd-HHmmss")
$acca | Out-File -filepath $Naccf -encoding unicode -ErrorAction SilentlyContinue
Write-Log "Accounts existence check complete and save to $Naccf"

######
# Complete new account configuration
######
$Command = "schtasks.exe /run /s $LyncS /TN 'AccountSync-LY'"
Write-Log ("Trigger task on $LyncS")
Invoke-Expression $Command
Clear-Variable Command -ErrorAction SilentlyContinue
Write-Log ("Task on $LyncS triggered")


######
# Remove extra account
######
$connL = Connect-QADService uttra.local -Credential $LyncCred
$accs = Get-QADUser -SizeLimit 50000 -Connection $connL -LDAPFilter $Lfilter 
foreach ($a in $accs) {
	 #$j = $result | where {$_.name -eq $a.name}
	 if (($a.name.trim() -ne "") -and (! $R.ContainsKey($a.name))) {
	 	  Remove-Mailbox -Identity $a.name -Confirm:$False
	 	  Write-Log ("Extra account " + $a.name + " removed")
	 }
}

#Remove-PsSession $session

$oldfiles = (Get-childitem "c:\SOURCE\test" ASync.*.csv | where {$_.lastwritetime -lt (Get-Date).AddDays(-$keeplog)} | %{$_.versioninfo.filename})
foreach ($ff in $oldfiles) {
   Remove-Item $ff
   Write-Log ("{0} removed" -f $ff)
}

$oldfiles = (Get-childitem "c:\SOURCE\test" NewAdd.*.csv | where {$_.lastwritetime -lt (Get-Date).AddDays(-$keeplog)} | %{$_.versioninfo.filename})
foreach ($ff in $oldfiles) {
   Remove-Item $ff
   Write-Log ("{0} removed" -f $ff)
}

$endT = get-date
$ts = $endT - $startT
Write-Output $("Process time [{0:00}:{1:00}:{2:00}]" -f $ts.Hours,$ts.Minutes,$ts.Seconds)
Write-Log $("Process time [{0:00}:{1:00}:{2:00}]" -f $ts.Hours,$ts.Minutes,$ts.Seconds)
Write-Log "--- Job End ---"
