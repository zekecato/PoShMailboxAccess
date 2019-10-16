
enum IMAPMessageFetchMode {
    
    Tiny

    Minimal

    Basic

    Full

}

# load the assembly for interacting with mailbox via Imap
$IMAPDLLPath = "$PSScriptRoot\ImapX.dll"
try{
    [void] [Reflection.Assembly]::LoadFile($IMAPDLLPath)
}catch{
    Write-Warning "Problem loading ImapX.dll. Library can be found at https://www.nuget.org/packages/ImapX/"
    throw $_
}

#load the assembly for interacting with Exchange mailboxes
$ExchangeDLLPath = "$PSScriptRoot\Microsoft.Exchange.WebServices.dll"
try{
    [void] [Reflection.Assembly]::LoadFile($ExchangeDLLPath)
}catch{
    Write-Warning "Problem loading Microsoft.Exchange.WebServices.dll. Get it at https://www.microsoft.com/en-us/download/details.aspx?id=35371"
    throw $_
}

#load the default server and port config for IMAP
$Config = Import-PowerShellDataFile "$PSScriptRoot\MailboxConfig.psd1"
$Script:MailboxConfig = $Config.MailboxConfig

#Get public and private function definition files.
    $Public  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )
    $Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )

#Dot source the files
    Foreach($import in @($Public + $Private))
    {
        Try
        {
            . $import.fullname
        }
        Catch
        {
            Write-Error -Message "Failed to import function $($import.fullname): $_"
        }
    }

# Here I might...
    # Read in or create an initial config file and variable
    # Export Public functions ($Public.BaseName) for WIP modules
    # Set variables visible to the module and its functions only

Export-ModuleMember -Function $Public.Basename
