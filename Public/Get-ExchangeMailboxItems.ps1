#this function retrieves all items in the specified mailbox folder for processing.
function Get-ExchangeMailboxItems{
    param(
        [Parameter(
            Mandatory=$true
        )]
        [Microsoft.Exchange.WebServices.Data.Folder]
        $Mailbox,

        [Parameter(
            HelpMessage = "Enter the number of results to retrieve from the server"
        )]
        [int]
        $Results,


        [Parameter()]
        [Microsoft.Exchange.WebServices.Data.LogicalOperator]
        $LogicalOperator = [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And ,

        [Parameter(

        )]
        [String[]]
        $Subject,

        [Parameter(

        )]
        [String]
        $Body,

        [Parameter(

        )]
        [String]
        $From,

        [Parameter(

        )]
        [switch]
        $Attachments,

        [Parameter(

        )]
        [switch]
        $Unread,

        [Microsoft.Exchange.WebServices.Data.BodyType]
        $BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

    )
#If number of results are not specified, will grab every message
if (!$Results){
    $Results = [int]::MaxValue
}

#load message properties to be retrieved.
$PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$PropertySet.RequestedBodyType = $BodyType
$ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($Results)
$ItemView.PropertySet = $PropertySet

#region Load Search Filters
$FilterGroup = @()

#Gets emails with subject containing specified string
if ($Subject){
    foreach($string in $Subject){
        Write-Verbose "Adding subject filter to mail request"
        $SubjectFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring
        $SubjectFilter.PropertyDefinition = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject
        $SubjectFilter.Value = $string
        $FilterGroup += $SubjectFilter
    }
}

if ($Body){
Write-Verbose "Adding body filter to mail request"
$BodyFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring
$BodyFilter.PropertyDefinition = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Body
$BodyFilter.Value = $Body
$FilterGroup += $BodyFilter
}

if ($From){
Write-Verbose "Adding from filter to mail request"
$FromFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo
$FromFilter.PropertyDefinition = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::From
$FromFilter.Value = $From
$FilterGroup += $FromFilter
}

#Gets unread emails
if ($Unread.IsPresent){
Write-Verbose "Adding unread filter to mail request"
$UnreadFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo
$UnreadFilter.PropertyDefinition = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead
$UnreadFilter.Value = $false
$FilterGroup += $UnreadFilter
}

#Gets emails with attachments
if ($Attachments.IsPresent){
Write-Verbose "Adding attachment filter to mail request"
$AttachmentFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo
$AttachmentFilter.PropertyDefinition = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments
$AttachmentFilter.Value = $true
$FilterGroup += $AttachmentFilter
}

#Collect all search filters together and retrieve items
Write-Verbose "Getting mail messages"
if ($FilterGroup){
$SearchFilterCollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection($LogicalOperator,$FilterGroup)

    $Messages = $Mailbox.FindItems($SearchFilterCollection, $ItemView)
}else{
    $Messages = $Mailbox.FindItems($ItemView)
}
#endregion Load Search Filters

    return $Messages

}