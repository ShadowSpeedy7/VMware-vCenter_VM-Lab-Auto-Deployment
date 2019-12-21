Import-Module VMware.VimAutomation.Core
Set-ExecutionPolicy -Scope Process Bypass
Add-Type -AssemblyName System.Windows.Forms

#$vCenterIP = "192.168.1.1"
$vCenterIP = Read-Host "Enter vCenter IP "

#$user = "administrator@something.something"
$user = Read-Host "Enter vCenter Username "

#$password = "blah"
$password = Read-Host "Enter vCenter Password "

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'Only Kick Ass CSV Files Bro|*.csv'
    Title = 'Super Lab Creator Deluxe Pro Gamer Edition - Version 9000.02'
}
$null = $FileBrowser.ShowDialog()
if ('' -eq $FileBrowser.FileName) {
    Write-Host "No File Selected"
    exit
}

Write-Host "[A] = Add VMs to vCenter based on csv list"
Write-Host "[D] = Delete VMs from vCenter based on csv list"
Write-Host "[C] = Cancel"
$Action = Read-Host "What action will this script be performing? [A/D/C] "
if ($Action -ne "A" -and $Action -ne "D") {
    Write-Host "Cancelling Action - Exiting ..."
    Exit
} elseif ($Action -eq "A") {
    $VMpower = Read-Host "Power VMs on after creation? [Y/N] "
}

$ToDoList = Import-Csv -Path $FileBrowser.FileName -ErrorAction Stop
$VMfolder = $FileBrowser.SafeFileName.TrimEnd(".csv")

$ToDoList
Write-Host "***PLEASE CONFIRM THE FOLLOWING INFORMATION***"
if ($Action -eq "A") {
    Write-Host "The VM list above will be added to vCenter $vCenterIP"
    Write-Host "The VMs will be added to folder $VMfolder after creation"
    if ($VMpower -eq "Y") {
        Write-Host "The VMs will be powered ON"
    } else {
        Write-Host "The VMs will remain powered OFF"
    }
} else {
    Write-Host "The VM list above will be powered OFF and deleted from vCenter $vCenterIP"
    Write-Host "The folder $VMfolder will be deleted afterwards"
}
if ((Read-Host "Continue? [Confirm]") -ne "Confirm") {
    Write-Host "Action Cancelled - Exiting ..."
    Exit
}

Connect-VIServer -Server $vCenterIP -User $user -Password $password

if ($Action -eq "A") {
    New-Folder -Name $VMfolder -Location "VM" -Confirm:$false
    for($i = 0; $i -lt $ToDoList.count; $i++) {
        New-VM -VMHost $ToDoList[$i]."VM-Host" -Name $ToDoList[$i]."VM-Name" -Location $VMfolder -Datastore $ToDoList[$i]."Datastore" -DiskStorageFormat "Thin" -Description $ToDoList[$i]."Description" -Template $ToDoList[$i]."Template" -RunAsync -Confirm:$false
    }
} elseif ($Action -eq "D") {
    for($i = 0; $i -lt $ToDoList.count; $i++) {
        Stop-VM -VM $ToDoList[$i]."VM-Name" -Confirm:$false -ErrorAction SilentlyContinue
        Remove-VM -VM $ToDoList[$i]."VM-Name" -DeleteFromDisk -Confirm:$false
    }
    Remove-Folder -Folder $VMfolder -Confirm:$false
}

Disconnect-VIServer -Server $vCenterIP -Force -Confirm:$false
pause