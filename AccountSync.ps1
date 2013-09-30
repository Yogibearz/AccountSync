#Requires -version 2.0
#
# version : '0.4' 
# 2013/2/22 上午 09:07:54
#
# Displayname Rule
# 1. 有 firstname & lastname and mail 含 @umc.com and 中文姓名 : 姓名(英文)
# 2. 有 firstname & lastname and mail 含 @umc.com and 英文姓名 : 名 姓
# 3. 有 firstname & lastname and 無 mail and 中文姓名 : 姓名
# 4. 有 firstname & lastname and 無 mail and 英文姓名 : 名 姓
# 5. 無 firstname : 工號
#
# Powershell -command "& {.\Set-LyncCredential.ps1 jcsecfile.txt}"
# Powershell -command "& {.\AccountSync.0.4.ps1}"

set-psdebug -strict

$lyncdomain = "@lpoc.com"

add-PSSnapin quest.activeroles.admanagement
#Set-QADProgressPolicy -ShowProgress $false | out-null
$pwd = Get-Location

#set-location $pwd

. .\Script4logging.ps1
$LogFile = "$pwd\AccountSync.log"
Write-Log "=== Job Begin ==="

$secfile = "$pwd\jcsecfile.txt"
$ACC = 'umc\00003058'

if (Test-Path $secfile) {
	 . .\Get-LyncCredential.ps1 $ACC $secfile
} else {
	 Write-Log "$secfile missing"
	 exit -200
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

$Lfilter = "(&(!name=IUSR_*)(!name=IWAM_*)(!name=SystemMailbox*)(!name=*$))"
#$acc = Get-QADUser -SearchRoot "CN=users,DC=umc,DC=com" -sizelimit 0 -LDAPFilter $Lfilter | Sort-Object Name #| Foreach-Object {
$acc = Get-QADUser -SearchRoot "CN=users,DC=umc,DC=com" -sizelimit 0 -LDAPFilter $Lfilter -Credential $credential | Sort-Object Name

foreach ($a in $acc) {
   $count++
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
   	     $tObj | Add-Member NoteProperty UPN $($a.Name + $lyncdomain)
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
   	     $tObj | Add-Member NoteProperty UPN $($a.Name + $lyncdomain)
      }
   } else {
   	     $tObj | Add-Member NoteProperty Alias $a.Name
   	     $tObj | Add-Member NoteProperty Name $a.Name
   	     $tObj | Add-Member NoteProperty LMAccount $("UMC\" + $a.Name)
   	     $tObj | Add-Member NoteProperty DisplayName $a.Name
   	     $tObj | Add-Member NoteProperty UPN $($a.Name + $lyncdomain)
   }
   $result += $tObj
} 

$result | Export-CSV -Path "$outfile" -encoding unicode -notype -ErrorAction SilentlyContinue 

Write-Log "Export $count accounts to csv $outfile"

######
# Add missing account
######
#foreach ($a in $result) {
#	 # check mailbox existence
#	 if ([bool](Get-Mailbox -Identity $a.UPN -ErrorAction SilentlyContinue)) {
#	    # check display name correctness
#	    $Mbox = Get-Mailbox -Identity $a.UPN
#	    if ($a.DisplayName -ne $Mbox.DisplayName) {
#	    	 #Set-Mailbox -Identity $a.UPN -DisplayName $a.DisplayName
#	    	 Write-Log $a.UPN + " DisplayName not match"
#	    }
#	 } else {
#	 	  # add new mailbox
#	 	  #New-Mailbox -Alias $a.Alias -Name $a.Name -UserPrincipalName $a.UPN -LinkedDomainController "UMCZ03.umc.com" -LinkedMasterAccount $a.LMAccount -OrganizationalUnit Lync -LinkedCredential $credential -DisplayName $a.DisplayName
#	 	  Write-Log "New mailbox " + $a.UPN + " created"
#	 }
#}

######
# Remove extra account
######
#$accs = Get-CsAdUser
#foreach ($a in $accs) {
#	 $j = $result | where {$_.name -eq $a.name}
#	 if ($j.name - $a.name) {
#	 	  #Remove-Mailbox -Identity $j.name
#	 	  Write-Log "Extra account " + $a.name + " removed"
#	 }
#}


$endT = get-date
$ts = $endT - $startT
Write-Output $("Process time [{0:00}:{1:00}:{2:00}]" -f $ts.Hours,$ts.Minutes,$ts.Seconds)
Write-Log $("Process time [{0:00}:{1:00}:{2:00}]" -f $ts.Hours,$ts.Minutes,$ts.Seconds)
Write-Log "--- Job End ---"
