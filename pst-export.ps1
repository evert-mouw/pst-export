# Exchange 2016 Mailbox Export for specific users
# needs the Mailbox Replication service
# Evert Mouw

Get-MailboxExportRequest | Remove-MailboxExportRequest -Confirm:$false

New-MailboxExportRequest -Mailbox USER1 -FilePath "\\EXCHANGESERVER\pst-export\USER1.pst"
New-MailboxExportRequest -Mailbox USER2 -FilePath "\\EXCHANGESERVER\pst-export\USER2.pst"

while ( Get-MailboxExportRequest | Where {$_.Status -eq "Queued" -or $_.Status -eq "InProgress"} )
{
    Write-Host "Waiting..."
    Start-Sleep -Seconds 30
}
