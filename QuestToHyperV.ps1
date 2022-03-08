$continue = 1

$RestoreName = ""
$CoreName = ""
$DomainName = ''
$ExportLocation = "<RESTORE PATH>"
$Network = "<NETWORK NAME>"

while($continue -ge 1){

if($cred -eq $null){
$cred = $host.ui.PromptForCredential("Backup Testing Export Program", "Please enter your $DomainName Domain Administrator Credentials.`n These credentials will be used to access Quest Rapid Recovery on $CoreName and Hyper-V on $RestoreName.","","")
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.Password)
$UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
}

if($cred.username -notlike "$DomainName\*"){
    $username = $cred.username
    $hostusername = "$DomainName\$username"
}
else{
    $hostusername = $cred.username
    $username = $cred.username.substring($DomainName.length+1)
}

Write-Host "Welcome $username !"

if ((New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
Write-Output "Elevated."
}
else{
Add-Type -AssemblyName PresentationCore,PresentationFramework
$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Error
$MessageBody = "You are not elevated. The Quest Rapid Recovery Core will not be accessable, Please exit and run this in an elevated state. Continue Anyway?"
$MessageTitle = "Oh no."
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
if($Result -ne "Yes"){
    exit
}
}

$ExportHost = "$RestoreName"
$Password = $UnsecurePassword

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select a Computer'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please select a computer:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

$servers = Get-ProtectedServers | Where-Object status -eq "Online" | Sort-Object -Property DisplayName

foreach($server in $servers){

[void] $listBox.Items.Add($server.DisplayName)

}

$form.Controls.Add($listBox)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $SelectedMachine = $listBox.SelectedItem
    $SelectedMachine
}
else{
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select a Computer to Restore'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please select a timestamp:'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

$points = Get-RecoveryPoints -ProtectedServer $SelectedMachine -number 5

$MachineRealName = $points[0].AgentHostName

foreach($point in $points){

[void] $listBox.Items.Add($point.DateTimestamp)

}

$form.Controls.Add($listBox)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $SelectedTimeStamp = $listBox.SelectedItem
    $SelectedTimeStamp
    foreach($point in $points){
        if($point.DateTimestamp -eq $SelectedTimeStamp){
            $Number = $point.Number
            #debug
            #Write-Host "Point is $Number"
        }
    }
}
else{
    exit
}

$date = Get-Date -Format "MM/dd/yyyy"

$log = "$MachineRealName`t$date`tExport to $RestoreName`t$CoreName`t$SelectedTimeStamp"

$log | clip

Write-Host "Information Copied to Clipboard. Select Column B when updating the testing log."

#Write-Host "Are you sure that you would like to export $MachineRealName to $RestoreName for the $SelectedTimeStamp backup?"

Add-Type -AssemblyName PresentationCore,PresentationFramework
$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Information
$MessageBody = "Are you sure that you would like to export $MachineRealName to $ExportHost for the $SelectedTimeStamp backup?"
$MessageTitle = "Export Confimation"
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
if($Result -ne "Yes"){
    $ResultMsg = "Machine not exported."
}
else{
Write-Host "Waiting 5 seconds before beginning export."

Start-Sleep -Seconds 5

Write-Host "Starting Export."

Start-HyperVExport -User $UserName -Password $Password -ProtectedServer $SelectedMachine -hostname $ExportHost -hostport 8010 -hostusername $HostUserName -hostpassword $Password -vmlocation $ExportLocation -rpn $Number -vmname $MachineRealName -usesourceram

Invoke-Command -ComputerName $ExportHost -Credential $cred -ScriptBlock {Add-VMNetworkAdapter -VMName $MachineRealName -SwitchName $Network}

$ResultMsg = "Machine has been exported."
}

Add-Type -AssemblyName PresentationCore,PresentationFramework
$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Information
$MessageBody = "$ResultMsg Would you like to export another machine?"
$MessageTitle = "OK"
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
if($Result -ne "Yes"){
    $continue = 0
}

}
