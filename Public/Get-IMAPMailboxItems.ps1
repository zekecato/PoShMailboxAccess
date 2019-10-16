<#
.Synopsis
   Fetches mailbox messages given a connected IMAP mailbox and message filters
.INPUTS
   none
.OUTPUTS
   IMAP message objects
#>

function Get-IMAPMailboxItems {
    #Looks in the inbox and returns all of the message items for processing
    param(
        [Parameter(
            Mandatory=$true
        )]
        [ImapX.Folder]
        $Mailbox,

        [Parameter()]
        [String]
        $Filter = 'ALL',
        
        [Parameter()]
        [IMAPMessageFetchMode]
        $MessageFetchMode = [IMAPMessageFetchMode]::Full


    )
    try{
    #Get messages from the mailbox. Gets ALL messages. Downloads FULL message data, maximum returned messages is all messages existing in the inbox.
    $Messages = $Mailbox.Search($Filter,$MessageFetchMode.ToString(),$Mailbox.Exists)
    }catch{
    Write-Warning "Unable to retrieve messages from specified mailbox"
    throw $_
    }
    return $Messages
}
