$ErrorActionPreference = "Stop"

add-pssnapin VMware.VimAutomation.Core

# Use this to enter credentials:
# $pw = read-host “Enter Password” –AsSecureString
# ConvertFrom-SecureString $pw | out-file "C:\path\to\password\file\here\textfile.txt"

$pwdSec = Get-Content "C:\path\to\password\file\here\textfile.txt" | ConvertTo-SecureString

$bPswd = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pwdSec)
$pswd = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bPswd)


$vcenter_server = "server"
$vcenter_user = "domain\user"

Function Send-Mail  {

    $ol = New-Object -comObject Outlook.Application
    $mail = $ol.CreateItem(0)
    $Mail.Recipients.Add("username@email.com")
    $Mail.Subject = "Could not connect to vsphere"
    $Mail.HTMLBody =  "Could not connect to vSphere - maybe your creds expired for the CD-Rom report script"
    $Mail.Send()

}

function ConnectViServer ($Server) {

    try {
            
        Connect-VIServer -Server $Server -Protocol https -User $vcenter_user -Password $pswd

        }

    Catch {

        Send-Mail
        
        }

 }

 ConnectViServer $vcenter_server

$VM = Get-VM | Get-CDDrive | ? {$_.ConnectionState.Connected -eq $true} | select Parent
                               

if ($VM) {

    $ol = New-Object -comObject Outlook.Application
    $Mail = $ol.CreateItem(0)
    $Mail.Recipients.Add("username@email.com")
    $Mail.Subject = "VMs with virtual cd-roms connected report"
    $Mail.HTMLBody = $($VM | ConvertTo-HTML | Out-String)
    $Mail.Send()
    
}