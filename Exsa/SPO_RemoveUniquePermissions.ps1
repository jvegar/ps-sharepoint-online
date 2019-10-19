#Load SharePoint Online Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
 
#To call a non-generic method Load
Function Invoke-LoadMethod() {
    param(
            [Microsoft.SharePoint.Client.ClientObject]$Object = $(throw "Please provide a Client Object"),
            [string]$PropertyName
        ) 
   $ctx = $Object.Context
   $load = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load") 
   $type = $Object.GetType()
   $clientLoad = $load.MakeGenericMethod($type)
   
   $Parameter = [System.Linq.Expressions.Expression]::Parameter(($type), $type.Name)
   $Expression = [System.Linq.Expressions.Expression]::Lambda([System.Linq.Expressions.Expression]::Convert([System.Linq.Expressions.Expression]::PropertyOrField($Parameter,$PropertyName),[System.Object] ), $($Parameter))
   $ExpressionArray = [System.Array]::CreateInstance($Expression.GetType(), 1)
   $ExpressionArray.SetValue($Expression, 0)
   $clientLoad.Invoke($ctx,@($Object,$ExpressionArray))
}
    
##Variables for Processing
$SiteUrl = "https://exsanet.sharepoint.com/sites/psad"
$ListName = "Documentos Publicados" 

#Get Credentials to connect
$userID = "Sharepointadmin@exsa.net"
$pwd = ConvertTo-SecureString 'Sh4repoint' -AsPlainText -Force
$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
   
#Set up the context
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl) 
$Context.Credentials = $Credentials
    
#Get the List
$List = $Context.web.Lists.GetByTitle($ListName)
 
$Query = New-Object Microsoft.SharePoint.Client.CamlQuery
$Query.ViewXml = "<View Scope='RecursiveAll'><RowLimit>2000</RowLimit></View>"
 
#Batch process list items - to mitigate list threashold issue on larger lists
Do {  
    #Get items from the list in batches
    $ListItems = $List.GetItems($Query)
    $Context.Load($ListItems)
    $Context.ExecuteQuery()
           
    $Query.ListItemCollectionPosition = $ListItems.ListItemCollectionPosition
  
    #Loop through each List item
    ForEach($ListItem in $ListItems)
    {
        Invoke-LoadMethod -Object $ListItem -PropertyName "HasUniqueRoleAssignments"
        $Context.ExecuteQuery()
        if ($ListItem.HasUniqueRoleAssignments -eq $true)
        {
            #Reset Permission Inheritance
            $ListItem.ResetRoleInheritance()
            Write-host  -ForegroundColor Yellow "Inheritence Restored on Item:" $ListItem.ID
        }
    }
    $Context.ExecuteQuery()
} While ($null -ne $Query.ListItemCollectionPosition)
  
Write-host "Broken Permissions are Deleted on All Items!" -ForegroundColor Green



