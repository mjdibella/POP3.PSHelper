function Connect-POP3 {
    Param(
        [Parameter(Mandatory=$true)][string]$server,
        [Parameter(Mandatory=$true)][string]$username,
        [Parameter(Mandatory=$true)][string]$password,
        [Parameter(Mandatory=$false)][int]$port = 110,
        [Parameter(Mandatory=$false)][bool]$enableSSL = $false
    )
    $pop3Client = New-Object OpenPop.Pop3.Pop3Client
    $pop3Client.connect( $server, $port, $enableSSL )
    if (!$pop3Client.connected) {
        throw "Unable to create POP3 client. Connection failed with server $server"
    }
    $pop3Client.authenticate( $username, $password )
    return $pop3Client
}

function Get-POP3Message {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.Pop3Client]$pop3Client,
        [Parameter(Mandatory=$false)][int]$messageIndex
    )
    if ($messageIndex) {
        $incomingMessage = $pop3Client.getMessage($messageIndex)
        $incomingMessage
    } else {
        $messageCount = $pop3Client.getMessageCount()
        for ($messageIndex = $messageCount; $messageIndex -gt 0; $messageIndex--) {
            $incomingMessage = $pop3Client.getMessage($messageIndex)
            $incomingMessage
        }
    }
}

function Remove-POP3Message {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.Pop3Client]$pop3Client,
        [Parameter(Mandatory=$true)][int]$messageIndex
    )
    $pop3Client.deleteMessage($messageIndex)
}

function Clear-POP3Mailbox {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.Pop3Client]$pop3Client
    )
    $pop3Client.deleteAllMessages
}

function ConvertTo-MIMEMessage {
    [cmdletbinding()]
    param(
        [Parameter(ValueFromPipeline)][string[]]$messages
    )
    begin {
    }
    process {
        foreach ($message in $messages) {
            $inStream = New-Object IO.FileStream $message,"Open"
            $mimemessage = [OpenPop.Mime.Message]::load( $inStream )
            $inStream.close()
            $mimemessage
        }
    }
    end {
    }
}

function Get-POP3UIDL {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.Pop3Client]$pop3Client,
        [Parameter(Mandatory=$false)][int]$messageIndex
    )
    if ($messageIndex) {
        $uid = $pop3Client.getMessageUID($messageIndex)
        $size = $pop3Client.getMessageSize($messageIndex)
        $results = New-Object PSObject
        $results | Add-Member Noteproperty number $number
        $results | Add-Member Noteproperty uid $uid
        $result | Add-Member Noteproperty size $size
    } else {
        $messageCount = $pop3Client.getMessageCount()
        for ($messageIndex = $messageCount; $messageIndex -gt 0; $messageIndex--) {
            $uid = $pop3Client.getMessageUID($messageIndex)
            $size = $pop3Client.getMessageSize($messageIndex)
            $result = New-Object PSObject
            $result | Add-Member Noteproperty number $messageIndex
            $result | Add-Member Noteproperty uid $uid
            $result | Add-Member Noteproperty size $size
            $result
        }
    }
    $results
}

function Disconnect-POP3 {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.Pop3Client]$pop3Client
    )
    $pop3Client.Disconnect | Out-Null
    $pop3Client.Dispose | Out-Null
    Remove-Variable -Name pop3Client
}

[Reflection.Assembly]::LoadFile($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath("OpenPop.dll"))