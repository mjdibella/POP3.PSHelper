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
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session,
        [Parameter(ValueFromPipeline)][PSObject[]]$messageObjects,    
        [Parameter(Mandatory=$false)][int]$index
    )
    begin {
    }
    process {
        if ($_) {
            foreach ($messageObject in $messageObjects) {
                $incomingMessage = $session.getMessage($messageObject.index)
                $incomingMessage | Add-Member Noteproperty Index $messageObject.index
                $incomingMessage
            }
        } else {
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
    }
}

function Get-POP3MessageHeader {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session,
        [Parameter(ValueFromPipeline)][PSObject[]]$messageObjects,    
        [Parameter(Mandatory=$false)][int]$index
    )
    begin {
    }
    process {
        if ($_) {
            foreach ($messageObject in $messageObjects) {
                $incomingMessageHeaders = $session.getMessageHeaders($messageObject.index)
                $incomingMessageHeaders | Add-Member Noteproperty Index $messageObject.index
                $incomingMessageHeaders
            }
        } else {
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
    }
}

function Get-POP3RawMessage {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session,
        [Parameter(ValueFromPipeline)][PSObject[]]$messageObjects,    
        [Parameter(Mandatory=$false)][int]$index
    )
    begin {
    }
    process {
        if ($_) {
            foreach ($messageObject in $messageObjects) {
                $rawMessage = $session.getMessageAsBytes($messageObject.index)
                [System.Text.Encoding]::ASCII.GetString($rawMessage)
            }
        } else {
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
    }
}

function Remove-POP3Message {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session,
        [Parameter(ValueFromPipeline)][PSObject[]]$messageObjects,    
        [Parameter(Mandatory=$false)][int]$index
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
    param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session,
        [Parameter(ValueFromPipeline)][PSObject[]]$messageObjects,    
        [Parameter(Mandatory=$false)][int]$index
    )
    begin {
    }
    process {
        if ($_) {
            foreach ($messageObject in $messageObjects) {
                $uid = $session.getMessageUID($messageObject.index)
                $size = $session.getMessageSize($messageObject.index)
                $result = New-Object PSObject
                $result | Add-Member Noteproperty index $messageObject.index
                $result | Add-Member Noteproperty uid $uid
                $result | Add-Member Noteproperty size $size
                $result
            }
        } else {
            if ($index) {
                $uid = $session.getMessageUID($index)
                $size = $session.getMessageSize($index)
                $result = New-Object PSObject
                $result | Add-Member Noteproperty index $index
                $result | Add-Member Noteproperty uid $uid
                $result | Add-Member Noteproperty size $size
                $result
            } else {
                $messageCount = $session.getMessageCount()
                for ($index = $messageCount; $index -gt 0; $index--) {
                    $uid = $session.getMessageUID($index)
                    $size = $session.getMessageSize($index)
                    $result = New-Object PSObject
                    $result | Add-Member Noteproperty index $index
                    $result | Add-Member Noteproperty uid $uid
                    $result | Add-Member Noteproperty size $size
                    $result
                }
            }
        }
    }
}

function Disconnect-POP3 {
    Param(
        [Parameter(Mandatory=$true)][OpenPop.Pop3.pop3client]$session
    )
    $session.Disconnect()
    $session.Dispose()
    Remove-Variable -Name session
}

$moduleName = $executionContext.sessionState.module
foreach ($modulePath in $($env:PSModulePath).Split(";")) {  
    $loaded = $False
    $dllPath = "$($modulePath.TrimEnd("\"))\$moduleName\OpenPop.dll"
    if ([System.IO.File]::Exists($dllPath)) {
        [Reflection.Assembly]::LoadFile($dllPath)
        $loaded = $True
        break
    }
}
if (-not $loaded) {
    Write-warning "Required file missing from module path."
}
Write-host "Cmdlets added:`n$(Get-Command | where {$_.ModuleName -eq $moduleName})`n"