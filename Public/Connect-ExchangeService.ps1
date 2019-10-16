#This function makes the connection to Exchange and returns the service and inbox objects for use in other functions.
Function Connect-ExchangeService {
    param (
        [Parameter(
            Position = 0,
            HelpMessage="Enter the email address of the mailbox to connect to."
        )]
        [ValidatePattern(".+?@.+?\..+")]
        [string]$Email,

#        [Parameter()]
#        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]
#        $Folder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,

        [Parameter(ParameterSetName = 'Credential')]
        [ValidatePattern(".+?@.+?\..+")]
        [string]$Impersonate,

        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'Credential'
        )]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [Microsoft.Exchange.WebServices.Data.ExchangeVersion]
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
    )

    $ExchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

    if ($PSCmdlet.ParameterSetName -eq 'Credential'){
        Write-Verbose "Using specified credentials"
        $ExchangeService.Credentials = $Credential.GetNetworkCredential()
        if(!$Email){
            $Email = $Credential.UserName
        }
    }else{
        Write-Verbose "Using default credentials for user running command"
        $ExchangeService.UseDefaultCredentials = $true
    }

    try{
    Write-Verbose "Connecting to exchange mailbox"
    $ExchangeService.AutodiscoverUrl($Email, {$true})
    }catch{
    Write-Warning "Autodiscover failed. Check username and password."
    throw $_
    }

    if($Impersonate){
        $ExchangeService.ImpersonatedUserId = [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$Impersonate)

    }
    
    $Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ExchangeService,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)

    $ReturnItems = @{
        Service = $ExchangeService
        Inbox = $Inbox
    }

    return $ReturnItems
}
