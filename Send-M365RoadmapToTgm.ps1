$scriptPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$configPath = "$scriptPath\Settings.xml"
[xml]$ScriptConfig = Get-Content $configPath

#config
$Roadmap = 'https://www.microsoft.com/en-us/microsoft-365/RoadmapFeatureRSS'
$tokenTelegram = $ScriptConfig.Settings.TelegramSettings.TelegramBotToken
[string]$chatID = $ScriptConfig.Settings.TelegramSettings.ChatId
$MinutesRange = $ScriptConfig.Settings.ScheduleSettings

#Time settings
$ChecktTime = ((Get-Date).ToUniversalTime()).AddMinutes(-$MinutesRange) 
$CurrentTime = (Get-Date).ToUniversalTime()

function Send-TelegramMessage {

    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$true)]
        [string]$messageText,
        [Parameter(Mandatory=$true)]
        [string]$tokenTelegram,
        [Parameter(Mandatory=$true)]
        [string]$chatID
    )

    $URL_set = "https://api.telegram.org/bot$tokenTelegram/sendMessage"
      
    $body = @{
        text = $messageText
        parse_mode = "markdown"
        chat_id = $chatID
    }

    $messageJson = $body | ConvertTo-Json

    try {
        Invoke-RestMethod $URL_set -Method Post -ContentType 'application/json; charset=utf-8' -Body $messageJson
        Write-Verbose "Message has been sent"
    }
    catch {
        Write-Error "Can't sent message"
    }
    
}
#Collect messages from RSS
$messages = (Invoke-RestMethod -Uri $Roadmap -Headers $headerParams -Method Get)

#Format collected messages
$SortedMessages = $messages | Select-Object link, @{L='Status';E={$($_.category[0])}}, title, description, @{L='Update'; E={$($([dateTime]$_.updated)).ToUniversalTime()}}, `
@{L='PubDate'; E={$($([dateTime]$_.pubDate)).ToUniversalTime()}}


foreach($message in $SortedMessages){
    $PubDate = $message.PubDate
    $Status = $message.Status
    if ($PubDate -ge $ChecktTime -and $PubDate -le $CurrentTime) {
        $MessageTitle = '*' + $($message.title) + '*'
        $MessageBodyWithLink, $availability = $message.description -split '<br>Availability date: '
        if ($MessageBodyWithLink -like '*More info:*') {
            $MessageBody, $MoreInfo = $MessageBodyWithLink -split '<br>More info: '
            $MoreInfoFormatted = "[More Info]($MoreInfo)"
        }else {
            $MessageBody = $MessageBodyWithLink
            $MoreInfo = $null
        }
        $Availability = "*Availability date:* $availability"
        $UpDate = $message.Update
        $PubDate = $message.PubDate
        $CurrentStatus = "*Current status:* $status"
        $Published = "*Published:* $PubDate"
        $Updated = "*Updated:* $UpDate"
        
        if ($null -eq $MoreInfo) {
            [string]$TgmMessage = "$MessageTitle `n$MessageBody `n$Availability `n$CurrentStatus `n$Published `n$Updated"
        }else {
            [string]$TgmMessage = "$MessageTitle `n$MessageBody `n$Availability `n$CurrentStatus `n$Published `n$Updated `n$MoreInfoFormatted"
        }
        Send-TelegramMessage -messageText $TgmMessage -tokenTelegram $tokenTelegram -chatID $chatID
    }
}
