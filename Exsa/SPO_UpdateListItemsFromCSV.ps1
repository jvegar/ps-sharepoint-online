Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime"

$siteURL = "https://exsanet.sharepoint.com/sites/psad"
$userID = "Sharepointadmin@exsa.net"
$pwd = ConvertTo-SecureString 'Sh4repoint' -AsPlainText -Force
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)  
$ctx.credentials = $creds 
try{  
    $lists = $ctx.web.Lists  
    $list = $lists.GetByTitle("Documentos Vigentes")  
    $listItems = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())  
    $ctx.load($listItems)  
      
    $ctx.executeQuery()  
    foreach($listItem in $listItems)  
    {  
        Write-Host "ID - " $listItem["ID"] "Title - " $listItem["Title"]  
    }  
}  
catch{  
    write-host "$($_.Exception.Message)" -foregroundcolor red  
}  