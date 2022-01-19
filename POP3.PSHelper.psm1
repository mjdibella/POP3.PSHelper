function Connect-POP3 {
    Param(
        [Parameter(Mandatory=$true)][string]$server,
        [Parameter(Mandatory=$true)][string]$username,
        [Parameter(Mandatory=$true)][string]$password,
        [Parameter(Mandatory=$false)][int]$port = 110,
        [Parameter(Mandatory=$false)][bool]$enableSSL = $false
    )
    $session = New-Object OpenPop.Pop3.pop3client
    $session.connect( $server, $port, $enableSSL )
    if (!$session.connected) {
        throw "Unable to create POP3 client. Connection failed with server $server"
    }
    $session.authenticate( $username, $password )
    return $session
}

function Get-POP3Message {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session,
        [Parameter(Mandatory=$false)][int]$index
    )
    if ($index) {
        $incomingMessage = $session.getMessage($index)
        $incomingMessage | Add-Member Noteproperty Index $index
        $incomingMessage
    } else {
        $messageCount = $session.getMessageCount()
        for ($index = $messageCount; $index -gt 0; $index--) {
            $incomingMessage = $session.getMessage($index)
            $incomingMessage | Add-Member Noteproperty Index $index
            $incomingMessage
        }
    }
}

function Get-POP3MessageHeader {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session,
        [Parameter(Mandatory=$false)][int]$index
    )
    if ($index) {
        $incomingMessageHeaders = $session.getMessageHeaders($index)
        $incomingMessageHeaders | Add-Member Noteproperty Index $index
        $incomingMessageHeaders
    } else {
        $messageCount = $session.getMessageCount()
        for ($index = $messageCount; $index -gt 0; $index--) {
            $incomingMessageHeaders = $session.getMessageHeaders($index)
            $incomingMessageHeaders | Add-Member Noteproperty Index $index
            $incomingMessageHeaders
        }
    }
}

function Get-POP3RawMessage {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session,
        [Parameter(Mandatory=$false)][int]$index
    )
    if ($index) {
        $rawMessage = $session.getMessageAsBytes($index)
        [System.Text.Encoding]::ASCII.GetString($rawMessage)
    } else {
        $messageCount = $session.getMessageCount()
        for ($index = $messageCount; $index -gt 0; $index--) {
            $rawMessage = $session.getMessageAsBytes($index)
            [System.Text.Encoding]::ASCII.GetString($rawMessage)
        }
    }
}

function Remove-POP3Message {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session,
        [Parameter(ValueFromPipeline,parameterSetName="fromPipeline")][PSObject[]]$messageObjects,    
        [Parameter(Mandatory=$false,parameterSetName="byParam")][int]$index
    )
    begin {
    }
    process {
        if ($_) {
            foreach ($messageObject in $messageObjects) {
                $session.deleteMessage($messageObject.index)
            }
        } else {
            if ($index) {
                $session.deleteMessage($index)
            }
        }
    }
    end {
    }
}

function Clear-POP3Mailbox {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session
    )
    $session.deleteAllMessages()
}

function Reset-POP3Mailbox {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session
    )
    $session.reset()
}

function Get-POP3Capabilities {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session
    )
    $session.capabilities()
}

function Get-POP3UIDL {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session,
        [Parameter(Mandatory=$false)][int]$index
    )
    if ($index) {
        $uid = $session.getMessageUID($index)
        $size = $session.getMessageSize($index)
        $results = New-Object PSObject
        $results | Add-Member Noteproperty number $number
        $results | Add-Member Noteproperty uid $uid
        $result | Add-Member Noteproperty size $size
    } else {
        $messageCount = $session.getMessageCount()
        for ($index = $messageCount; $index -gt 0; $index--) {
            $uid = $session.getMessageUID($index)
            $size = $session.getMessageSize($index)
            $result = New-Object PSObject
            $result | Add-Member Noteproperty number $index
            $result | Add-Member Noteproperty uid $uid
            $result | Add-Member Noteproperty size $size
            $result
        }
    }
    $results
}

function Disconnect-POP3 {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session
    )
    $session.Disconnect()
    $session.Dispose()
    Remove-Variable -Name session
}

[Reflection.Assembly]::LoadFile($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath("OpenPop.dll"))