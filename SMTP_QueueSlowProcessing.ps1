#Slow processing local spammer notifications

#MailMarshal Host
$MailMarshalServer = 'MaiMarshalHost'

#Folder for local spammer
$spamfolder = "\\$MailMarshalServer\c$\spamfolder\"

$lists = Get-ChildItem $spamfolder

if ($lists.Count -gt 0) {

    do
    {
        foreach ($list in $lists) {
            $list
            Start-Sleep -Seconds 8
            $list | Move-Item -Destination '\\$MailMarshalServer\c$\Program Files\MailMarshal\Queues\Incoming\'

        } 
        $lists = Get-ChildItem $spamfolder
        
        Start-ScheduledTask -TaskName "SMTP_QueuesCleaner"

    }
    while ($lists.count -gt 0)


}
