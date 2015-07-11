    Param
    (
        [Parameter(Mandatory=$false)]
        $ExportPath = "\\localhost\E\"
    )

if (!(get-module|where path -Match "Exchange")){Write-Host "This script needs to be run from the Exchange Management Shell" -ForegroundColor Red
Exit}
if (!(Get-Command -Name New-MailboxExportRequest -ErrorAction Ignore)){Write-Host "You need to be a member of the `"Mailbox Import Export`" role to run this script." -ForegroundColor Red 
Write-Host "Please run `"New-ManagementRoleAssignment -Role `"Mailbox Import Export`" -User `"<user name or alias>`"`"" -ForegroundColor Red
Exit}

$Mailboxes = Get-Mailbox
if ($Mailboxes -gt 0){
foreach ($Mailbox in $Mailboxes) 
{New-MailboxExportRequest -Mailbox $Mailbox -FilePath "$ExportPath$($Mailbox.Alias).pst"}

do {
Clear-Host
Write-Host "$((Get-MailboxExportRequest|where status -NotMatch "Completed").count) left to export" -foregroundcolor "green"
Write-Host "$((Get-MailboxExportRequest|where status -eq inprogress).count) in progress" -foregroundcolor "yellow"
Get-MailboxExportRequest|where status -eq inprogress|Select-Object Mailbox,Filepath| Format-Table -AutoSize
sleep -Seconds 10
}
while ((Get-MailboxExportRequest).count -gt 0)

}
else
{
Write-host "No mailboxes found" -ForegroundColor Red
}