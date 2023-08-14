# PointOfService 폴더에 쓰기 작업 하려면 관리자 권한이 필요하다.
# PowerShell이 관리자 권한으로 실행중이 아니라면, UAC 창을 띄운다.
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $arguments = "& '" + $MyInvocation.MyCommand.Definition + "'"
    Start-Process powershell -Verb RunAs -ArgumentList $arguments
    Exit
}

############################## 팡션 ##############################

function global:ProcessSelectedItem {
    $is112Exist = Test-Path -Path $folderPath112 -PathType Container
    $is114Exist = Test-Path -Path $folderPath114 -PathType Container
    $nonX86Exist = Test-Path -Path $folderPathNonX86 -PathType Container

    if (-not $is112Exist) {
        [System.Windows.Forms.MessageBox]::Show('ERROR(112)', 'ERROR')
        return
    } elseif (-not $is114Exist) {
        [System.Windows.Forms.MessageBox]::Show('ERROR(112)', 'ERROR')
        return
    } elseif (-not $nonX86Exist) {
        New-Item -Path $folderPathNonX86 -ItemType Directory
    }

    switch ($listBox.SelectedItem) {
        $listBox.Items[0] {
            try {
                Copy-Item 'C:\Program Files (x86)\Microsoft Point Of Service_Bak\SDK\1.12\*' 'C:\Program Files (x86)\Microsoft Point Of Service\SDK' -Force
                Copy-Item 'C:\Program Files (x86)\Microsoft Point Of Service_Bak\SDK\1.12\*' 'C:\Program Files\Microsoft Point Of Service\SDK' -Force    
            } catch {
                $errorMessage = $_.Exception.Message
                [System.Windows.Forms.MessageBox]::Show($errorMessage, 'ERROR')
            }
        }
        $listBox.Items[1] {
            try {
                Copy-Item 'C:\Program Files (x86)\Microsoft Point Of Service_Bak\SDK\1.14\*' 'C:\Program Files (x86)\Microsoft Point Of Service\SDK' -Force    
                Copy-Item 'C:\Program Files (x86)\Microsoft Point Of Service_Bak\SDK\1.14\*' 'C:\Program Files\Microsoft Point Of Service\SDK' -Force    
            } catch {
                $errorMessage = $_.Exception.Message
                [System.Windows.Forms.MessageBox]::Show($errorMessage, 'ERROR')
            }
        }
    }
}

############################## 팡션 ##############################


$versionInfo = Get-ItemProperty -Path "C:\Program Files (x86)\Microsoft Point Of Service\SDK\Microsoft.PointOfService.dll" | Select-Object -ExpandProperty VersionInfo
$fileVersion = $versionInfo.FileVersion

$Global:folderPath112 = 'C:\Program Files (x86)\Microsoft Point Of Service_Bak\SDK\1.12'
$Global:folderPath114 = 'C:\Program Files (x86)\Microsoft Point Of Service_Bak\SDK\1.14'
$Global:folderPathNonX86 = 'C:\Program Files\Microsoft Point Of Service\SDK'

# Write-Host "Microsoft POS Version: $fileVersion"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = ''
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
$label.Text = "Current version: $fileVersion"
$form.Controls.Add($label)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,40)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = '---Select to change version---'
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,60)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 30

[void] $listBox.Items.Add('PointOfService 1.12')
[void] $listBox.Items.Add('PointOfService 1.144444')

$listBox.Add_MouseDoubleClick({
    ProcessSelectedItem
    $form.Dispose()
})


$form.Controls.Add($listBox)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    ProcessSelectedItem
}