<#
.Synopsis
   Connect to IMAP mailbox using ImapX library. Returns mailbox object for use in other operations.
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   none
.OUTPUTS
   ImapX folder object for Inbox
.NOTES
   General notes
.FUNCTIONALITY
   Logs into mailbox via IMAP and creates object that represents the mailfolder you use
#>

Function Connect-IMAPMailbox {
    param(
        
        [Parameter(
            
        )]
        [ValidateNotNullorEmpty()]
        [string]
        $Server = $MailboxConfig.Server,

        [Parameter(
            Mandatory = $true    
        )]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter(
            
        )]
        [string]
        $MailFolder = 'Inbox',

        [Parameter(
            
        )]
        [IMAPMessageFetchMode]
        $MessageFetchMode = [IMAPMessageFetchMode]::Tiny,

        [Parameter(

        )]
        [Switch]
        $SSL=$true,
        
        [int]
        $Port = $MailboxConfig.Port

    )

    #Convert credential into type accepted by IMAP library
    $NetCred = $Credential.GetNetworkCredential()
    #connect to mailbox.
    $client = New-Object ImapX.ImapClient
    $client.Behavior.MessageFetchMode = $MessageFetchMode.ToString()
    $client.Host = $Server
    $client.Port = $Port
    $client.UseSsl = [bool]$SSL
    $client.Connect() | out-Null
    
    try{
    [void] $client.Login($NetCred.UserName, $NetCred.Password)
    }catch{
        Write-Warning "Logging into Mailbox failed"
        throw $_
    }

    #Assign Inbox folder object to variable that can be accessed by other functions in the script
    $Inbox = $client.folders| Where-Object { $_.path -eq 'Inbox' }

    return $Inbox
}
