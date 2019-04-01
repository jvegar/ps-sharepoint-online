#Adding Sharepoint Online Client libraries 
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
#Setting global variables
$siteURL = "https://exsanet.sharepoint.com/sites/psad"
$userID = "Sharepointadmin@exsa.net"
$pwd = ConvertTo-SecureString 'Sh4repoint' -AsPlainText -Force
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)  
$ctx.credentials = $creds
$CSVMetadatos = Import-Csv -Path "EstructuraMetadatosFixedDelimiter.csv" -Delimiter '|' -Encoding Default

#Function that return list items by title
function loadListItems {
    param ([string]$listTitle)
    $lists = $ctx.Web.Lists
    $list = $lists.GetByTitle($listTitle)
    $listItems = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
    $ctx.load($listItems)      
    $ctx.executeQuery()
    return $listItems
}

#Function that get Id column from "Lista: Gerencias" list using Title and Pais columns
function getGerenciaId {
    param ($title, $pais, $items)
    foreach($item in $items){
        if(($item["Title"] -match $title) -and ($item["Pais"] -eq $pais)){
            return $item["ID"]
        }
    }
    return $null
}
#Function that get Id column from "Lista: Areas" list using Title and Pais columns
function getAreaId {
    param ($title, $pais, $gerenciaId, $items)
    foreach($item in $items){
        if(($item["Title"] -match $title) -and ($item["Pais"] -eq $pais) -and (([Microsoft.SharePoint.Client.FieldLookupValue]$item.FieldValues["GerenciaFuncional"]).LookupId -eq $gerenciaId)){
            return $item["ID"]
        }
    }
    return $null
}
try{  
    #Load Gerencias and Areas List Items
    $itemsGerencias = loadListItems("Lista: Gerencias")
    $itemsAreas = loadListItems("Lista: Areas")
    #Write-Host $itemsGerencias.Count
    #Write-Host $itemsAreas.Count
    $lists = $ctx.Web.Lists
    $list = $lists.GetByTitle("Documentos Vigentes")    
    $i = 0
    foreach($csvDoc in $CSVMetadatos){
        Write-Host "Proccessing " $csvDoc.CodigoNuevo " with Position:" ($i+1) "from " $CSVMetadatos.Count "elements."
        if($csvDoc.Gerencia -eq "OPERACIONES Y LOGÍSTICA DE CLIENTES"){
            $listItem = $list.GetItemById($csvDoc.IdDoc)  
            <# $listItem["CodAntiguo"] = $csvDoc.CodigoAntiguo
            $listItem["CodigoDoc"]  = $csvDoc.CodigoNuevo
            $listItem["EstadoDoc"]  = "Aprobado"
            $listItem["NroEdicion"]  = $csvDoc.NroEdicion
            $listItem["NroRevision"]  = $csvDoc.NroRevision
            $listItem["Pais"]  = "Perú"
            $listItem["Title"]  = $csvDoc.Titulo
            $listItem["TipoDoc"]  = $csvDoc.TipoDocumento #>
            $lookupGerencia = New-Object Microsoft.SharePoint.Client.FieldLookupValue
            $lookupGerencia.LookupId  = getGerenciaId $csvDoc.Gerencia $csvDoc.Pais $itemsGerencias             
            $listItem["GerenciaFuncional"]  = $lookupGerencia.LookupId
            $lookupArea = New-Object Microsoft.SharePoint.Client.FieldLookupValue
            $lookupArea.LookupId  = getAreaId $csvDoc.Area $csvDoc.Pais $lookupGerencia.LookupId $itemsAreas 
            $listItem["Area"] = $lookupArea.LookupId
            $listItem.Update()  
            $ctx.load($listItem)      
            $ctx.executeQuery()
        }        
        $i = $i + 1
    }
}  
catch{  
    write-host "$($_.Exception.Message)" -foregroundcolor red  
}  