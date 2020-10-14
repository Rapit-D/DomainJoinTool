$ScriptExecPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ConfigPath = $ScriptExecPath + "\Config.ini"
#check whether config exist
If (!(Test-Path $ConfigPath))
{
$ws = New-Object -ComObject WScript.Shell
$ws.popup("No Config File.",0)| Out-Null
break
}
Else {

#load winapi
$ini = Add-Type -MemberDefinition @"
[DllImport("Kernel32")]
public static extern int GetPrivateProfileString (
string Section ,  
string Key , 
string Def , 
StringBuilder RetVal ,  
int size , 
string FilePath ); 
"@ -Passthru -Name MyPrivateProfileString -UsingNamespace System.Text
#
$Section="Configuration"
$FilePath=$ConfigPath 
$RetVal=New-Object System.Text.StringBuilder(32)

#read domain info from config file

#read value of Domain 
$Key="Domain"
$null=$ini::GetPrivateProfileString($Section,$Key,"",$RetVal,64,$FilePath)
$Domain = $RetVal.ToString()

#read value of DNS 
$Key="DomainDNS"
$null=$ini::GetPrivateProfileString($Section,$Key,"",$RetVal,64,$FilePath)
$DomainDNS = $RetVal.ToString()

#read value of DomainController 
$Key="DomainController"
$null=$ini::GetPrivateProfileString($Section,$Key,"",$RetVal,64,$FilePath)
$DomainController = $RetVal.ToString()

#read value of DomainAdmin 
$Key="DomainAdmin"
$null=$ini::GetPrivateProfileString($Section,$Key,"",$RetVal,64,$FilePath)
$DomainAdmin = $RetVal.ToString()

#read value of DomainAdminPwd 
$Key="DomainAdminPwd"
$null=$ini::GetPrivateProfileString($Section,$Key,"",$RetVal,64,$FilePath)
$DomainAdminPwd = $RetVal.ToString()
$Bytes  = [System.Convert]::FromBase64String($DomainAdminPwd)
$DeCodePwd = [System.Text.Encoding]::UTF8.GetString($Bytes)

#read value of OU 
$Key="OU"
$null=$ini::GetPrivateProfileString($Section,$Key,"",$RetVal,64,$FilePath)
$OU = $RetVal.ToString()
}
Function Restart_Computer_Confirm {
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm2 = New-Object System.Windows.Forms.Form 
$objForm2.Text = "Computer Restart Confirming"
$objForm2.AutoSize = $True
$objForm2.Size = New-Object System.Drawing.Size(280,100)
$FormFont = New-Object System.Drawing.Font("宋体",14,[System.Drawing.FontStyle]::Regular)
$objForm2.Font = $FormFont
$objForm2.StartPosition = "CenterScreen"

$OKButtonAction2 = {
Log "Confirm to Restart."
Restart-Computer -Force
$objForm2.Close()
}

$CancelButtonAction2 = {
Log "Cancel Restarting."
$objForm2.Close()
}
$objLabel2 = New-Object System.Windows.Forms.Label
$objLabel2.Location = New-Object System.Drawing.Size(5,20)
$objLabel2.Size = New-Object System.Drawing.Size(280,60)
$objLabel2.Text = "All changes will be effective after restart, Please confirm..."
$objForm2.Controls.Add($objLabel2) 

$OKButton2 = New-Object System.Windows.Forms.Button
$OKButton2.Location = New-Object System.Drawing.Size(5,80)
$OKButton2.Size = New-Object System.Drawing.Size(155,35)
$OKButton2.Text = "Confirm"
$OKButton2.Add_Click($OKButtonAction2)
$objForm2.Controls.Add($OKButton2)

$CancelButton2 = New-Object System.Windows.Forms.Button
$CancelButton2.Location = New-Object System.Drawing.Size(165,80)
$CancelButton2.Size = New-Object System.Drawing.Size(155,35)
$CancelButton2.Text = "Cancel"
$CancelButton2.Add_Click($CancelButtonAction2)
$objForm2.Controls.Add($CancelButton2)

$objForm2.KeyPreview = $True
$objForm2.Add_KeyDown({if ($_.KeyCode -eq "Escape") {
Log "User cancel restarting."
$objForm2.Close()}})

$objForm2.Add_Shown({$objForm2.Activate()})
[void] $objForm2.ShowDialog()
}
#rename computer
Function RenameComputerName {
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm3 = New-Object System.Windows.Forms.Form 
$objForm3.Text = "Rename Computer"
$objForm3.AutoSize = $True
$objForm3.Size = New-Object System.Drawing.Size(280,100)
$FormFont = New-Object System.Drawing.Font("宋体",14,[System.Drawing.FontStyle]::Regular)
$objForm3.Font = $FormFont
$objForm3.StartPosition = "CenterScreen"

$OKButtonAction3 = {
$global:NewComputerName = $objTextBox3.Text
$ComputerInfo = Get-WmiObject -Class Win32_ComputerSystem
$global:Result = $ComputerInfo.Rename($NewComputerName)
	If ($Result.Return -eq 0)
	{
	$info = "Rename successfully! New ComputerNmae is $($NewComputerName).Program running, please wait..."
	WSnobreak $info 5
    $objForm3.Close()
			}
	Else {
	$info = "Fail to Rename，error code:$Result.Return."
	$objLabel3_4.Text = $info
	WS $info
	$objForm3.Close()
	}
}

$objLabel3_1 = New-Object System.Windows.Forms.Label
$objLabel3_1.Location = New-Object System.Drawing.Size(5,20)
$objLabel3_1.Size = New-Object System.Drawing.Size(280,20)
$objLabel3_1.Text = "Input New ComputerName:"
$objForm3.Controls.Add($objLabel3_1)

$objLabel3_2 = New-Object System.Windows.Forms.Label
$objLabel3_2.Location = New-Object System.Drawing.Size(5,40)
$objLabel3_2.Size = New-Object System.Drawing.Size(280,20)
$objLabel3_2.Text = ""
$objLabel3_2Font = New-Object System.Drawing.Font("宋体",10,[System.Drawing.FontStyle]::Regular)
$objLabel3_2.Font = $objLabel3_2Font
$objForm3.Controls.Add($objLabel3_2)

$objLabel3_3 = New-Object System.Windows.Forms.Label
$objLabel3_3.Location = New-Object System.Drawing.Size(5,90)
$objLabel3_3.Size = New-Object System.Drawing.Size(40,20)
$objLabel3_3.Size = New-Object System.Drawing.Size(40,20)
$objLabel3_3.Text = "Tips:"
$objLabel3_3.Font = $objLabel3_2Font
$objForm3.Controls.Add($objLabel3_3)

$objLabel3_4 = New-Object System.Windows.Forms.Label
$objLabel3_4.Location = New-Object System.Drawing.Size(45,90)
$objLabel3_4.Size = New-Object System.Drawing.Size(220,20)
$objLabel3_4.Text = ""
$objLabel3_4.Font = $objLabel3_2Font
$objForm3.Controls.Add($objLabel3_4)

$objTextBox3 = New-Object System.Windows.Forms.TextBox 
$objTextBox3.Location = New-Object System.Drawing.Size(5,60) 
$objTextBox3.Size = New-Object System.Drawing.Size(260,10)
$objTextBox3.Text = ""
$objForm3.Controls.Add($objTextBox3)

$OKButton3 = New-Object System.Windows.Forms.Button
$OKButton3.Location = New-Object System.Drawing.Size(5,110)
$OKButton3.Size = New-Object System.Drawing.Size(260,35)
$OKButton3.Text = "Confirm"
$OKButton3.Add_Click($OKButtonAction3)
$objForm3.Controls.Add($OKButton3)

$objForm3.KeyPreview = $True

$objForm3.Add_Shown({$objForm3.Activate()})
[void] $objForm3.ShowDialog()
}
#join domain and move to specific OU
Function JoinDomainAndMoveToOU($Restart) {
Log "Begin to join Domain."
cmd /c netsh advfirewall set publicprofile state off | Out-Null
LogAndWrite "firewall--public network,closed."
cmd /c netsh advfirewall set privateprofile state off | Out-Null
LogAndWrite "firewall-family or work network,closed."
Write-Warning "Computer will restart after joining Domian,please close running programs."
$ComputerInfo = Get-WmiObject -Class Win32_ComputerSystem
$JoinDomainResult = $ComputerInfo.JoinDomainOrWorkgroup($Domain,$DeCodePwd,$DomainAdmin,$OU,1027)
    If ($JoinDomainResult.Return -eq 0)
    {
    LogAndWrite "Computer join $($Domain)domain successfully，well be effective after restarting."
        If ($Restart -eq "Restart"){Restart_Computer_Confirm}
    }
    Else 
    {
    WS "Fail to Join to specific domain，contact Administator,error code：$($JoinDomainResult.Return)"
    }
}


Set-Location $ScriptExecPath

#store log
Remove-Item -Path ($($ScriptExecPath) + "\Log\" + "*.log") -Force | Out-Null
$LogFormat = "\" + (Get-Date).ToString("yyyyMMddhhmmss") + "_Prepare.log"
$Log = $($ScriptExecPath) + "\Log" + $LogFormat
New-Item -path ($($ScriptExecPath) + "\Log") -name $LogFormat -Itemtype File | Out-Null


#Log format
Function Log($LogContent){
$LogTime = (Get-Date).GetDateTimeFormats()[34] + " :"
"$($LogTime)$LogContent"| Out-File $Log -Append 
}

#Log and Write
Function LogAndWrite($LogContent){
$LogTime = (Get-Date).GetDateTimeFormats()[34] + " :"
"$($LogTime)$LogContent"| Out-File $Log -Append
Write $LogContent
}

#windows pop up 
Function WS($info) {
Log $info
$ws = New-Object -ComObject WScript.Shell
$ws.popup($info,0)| Out-Null
break
}

Function WSnobreak($info,$Wait) {
Log $info
$ws = New-Object -ComObject WScript.Shell
$ws.popup($info,$Wait)| Out-Null
}

#network check 
write "Checking the network......................"
If (Test-Connection $DomainController -Quiet){
Log "Successfully Connected to $($DomainController)."
}
else {
WS "Fail to connect to domain,contact your administrator to check the network!"
}
write "network check completed."

#port check
write "Checking DC net ports......................"
$Ports = @(53,88,135,389,445,636,3268)
foreach ($Port in $Ports) {
$Response = .\PortQry.exe -n $DomainController -e $Port | Select-String "TCP port"
If ($Response -match "NOT LISTENING"){WS "$($DomainController)的$($Port)Ports connesction fail."}
Else {Log "$($domain)的$($Port)Ports connection OK."}
}
write "DC net ports check completed."

#service check
write "Checking local server......................"
$Services = @("TCP/IP NetBIOS Helper","Workstation")
$Services | % {
If ((Get-Service $_).Status -eq "Running") {Log "$($_)Local server runing."}
Else {WS "$($_)Local server not running，please check status."}
}
write "Local server check completed."


#windows version check
write "Checking Operation System version......................"
$OSedition = (Get-WmiObject -Class Win32_OperatingSystem).Caption
If ($OSedition -match "家庭版") {WS "Operation System version is$OSedition，Cannot join domain，Please contact your Administrator."}
Else {Log "Operation System version is $($OSedition)."}
write "Operation System version check completed."


#share content check
write "Checking local share path......................"
$Share = (Get-WmiObject -Class Win32_Share) | % {$_.name}
$ShareItems = @("IPC$","ADMIN$")
Foreach ($item in $ShareItems){
$ShareItems_Reg = "^" + [regex]::Escape($item)
If ($Share -match $ShareItems_Reg){Log "$($item) shered."}
Else {WS "$($item)not shared，please contact your Administrator to add or change registry  as 1,then restart computer:HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\lanmanserver\parameters Name：AutoShareWks，Type：DWORD"}
}
write "Local share path check completed."


#grand full access permission to disk except Driver C
write "Authorizing partitions except C:\......................"
$Disks = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType='3'"
$DisksExpC = $Disks | where {$_.DeviceID -ne "C:"}
$DisksExpC | foreach {
Cacls $_.DeviceID /E /G BUILTIN\Users:f
Log "Complete authorizing to $($_.DeviceID)"
}
write "Disk authorizing completed."


#check lcoal DNS setting
$DNS = ipconfig /all | Select-String "DNS Servers"
$DomainDNS_Reg = [regex]::Escape($DomainDNS)
If ($DNS -match $DomainDNS_Reg){Log "DNS point to$DomainDNS,Configuration OK."}
Else {Log "DNS configuration not right，DNS point address will be changed to $($DomainDNS)."
#change local DNS setting
netsh interface ip delete dnsservers "Local Area Connection" all
netsh interface ip add dnsservers "Local Area Connection" $DomainDNS
Log "DNS point address changed to $($DomainDNS)."
}

LogAndWrite "Pre-conditions chenck for joining domain completed."

RenameComputerName
If ($Result.Return -eq 0)
        {
        LogAndWrite "Joining domain......"
        JoinDomainAndMoveToOU Restart
        }