#Slowing down messages in MailMarshal from employers-spammers (excessive notifications)

#MailMarshal Host
$MailMarshalServer = 'MaiMarshalHost'

#MailMarshal Sending folder path
$mails_path = "\\$MailMarshalServer\c$\Program Files\MailMarshal\Queues\Sending"

#Folder for local spammer
$spamfolder = "\\$MailMarshalServer\c$\spamfolder\"

#Mail
$smtpServer = "mail.bgelov.ru"
$smtpFrom = "notif@bgelov.ru"
$smtpTo = "alert-group@bgelov.ru"

#Telegram
$notifChat = -11111111111111

if ((Get-ChildItem $mails_path).Count -gt 1000) {

    $emls = @{}
    $mails = Get-ChildItem "$mails_path\*.mml"


<#Send notif to Telegram
    function send-telegram($chat_id, $text) {

        [switch]$markdown,
        [switch]$nopreview

        $token = "*******************************"
        if($nopreview) { $preview_mode = "True" }
        if($markdown) { $markdown_mode = "Markdown" } else {$markdown_mode = ""}

        $payload = @{
            "chat_id" = $chat_id;
            "text" = $text;
            "parse_mode" = $markdown_mode;
            "disable_web_page_preview" = $preview_mode;
        }

        Invoke-WebRequest `
            -Uri ("https://api.telegram.org/bot{0}/sendMessage" -f $token) `
            -Method Post `
            -ContentType "application/json;charset=UTF-8" `
            -Body (ConvertTo-Json -Compress -InputObject $payload)

    }
#>


    foreach ($mail in $mails) {
        $read = gc $mail 
        $whereFrom = $read -match "^From:\s"
        $parsing = $whereFrom  -replace 'from:\s=.+=' -replace '<' -replace '>'  -replace '.+ '
        $emls.Add($parsing,$mail.fullname)
    }

    $topemls = $emls.Keys | Group-Object | Sort-Object Count -Descending | select count, name -First 1

    if ($topemls.count -gt 600) { 

        $messageSubject = "Queue on MailMarshal from "+$topemls.name
        $body = "Message count: "+$topemls.count
        send-mailmessage -from "$smtpFrom" -to "$smtpTo" -subject "$messageSubject" -Body $body -smtpServer "$smtpserver"

        $text = $messageSubject+". "+$body+"."
        send-telegram -chat_id $notifChat -text $text

        foreach ($eml in $emls.GetEnumerator()) {

            if ($eml.name -eq $topemls.Name) { Move-Item $eml.Value -Destination $spamfolder -Force }

        }

     }

     Start-ScheduledTask -TaskName "SMTP_QueueSlowProcessing"

 }