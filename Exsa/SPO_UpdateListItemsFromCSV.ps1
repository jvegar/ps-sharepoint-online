#Adding Sharepoint Online Client libraries 
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
#Setting global variables
$siteURL = "https://exsanet.sharepoint.com/sites/psad"
$userID = "[username]"
$pwd = ConvertTo-SecureString '[password]t' -AsPlainText -Force
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)  
$ctx.credentials = $creds
$CSVMetadatos = Import-Csv -Path "EstructuraMetadatosTiposDocs.csv" -Delimiter ';' -Encoding Default
#$CSVMetadatos = Import-Csv -Path "D:\Installers\EliminarIDsVigentes.csv" -Delimiter ';' -Encoding Default
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
function getTipoDocId {
    param ($title, $pais, $areaId, $gerenciaId, $items)
    foreach($item in $items){
        if(($item["Title"] -match $title) -and ($item["Pais"] -eq $pais) -and (([Microsoft.SharePoint.Client.FieldLookupValue]$item.FieldValues["Area"]).LookupId -eq $areaId) -and (([Microsoft.SharePoint.Client.FieldLookupValue]$item.FieldValues["GerenciaFuncional"]).LookupId -eq $gerenciaId)){
            return $item["ID"]
        }
    }
    return $null
}
function getUserId {
    param ($userLogin)
    $UserID = "i:0#.f|membership|" + $userLogin
    $SPOUser = $ctx.Web.EnsureUser($UserID)    
    $ctx.Load($SPOUser)
    $ctx.ExecuteQuery()
    return $SPOUser.Id
}

try{  
    #Load Gerencias and Areas List Items
    #$itemsGerencias = loadListItems("Lista: Gerencias")
    #$itemsAreas = loadListItems("Lista: Areas")
    #$itemsTipoDocs = loadListItems("Lista: Tipos Documentos")
    #Write-Host $itemsGerencias.Count
    #Write-Host $itemsAreas.Count

    #$Query = New-Object Microsoft.SharePoint.Client.CamlQuery
    #$Query.ViewXml =$CAMLQuery

    $lists = $ctx.Web.Lists
    $list = $lists.GetByTitle("Lista: Tipos Documentos")
    
    #$listItems = $list.GetItems($Query)
    #$ctx.Load($listItems)
    #$ctx.ExecuteQuery()

    $i = 0
    foreach($csvDoc in $CSVMetadatos){
        try{ 
        Write-Host "Proccessing " $csvDoc.IdDoc " with Position:" ($i+1) "from " $CSVMetadatos.Count "elements."
        #if($csvDoc.Area -eq "rtenorio@exsa.net"){
            $listItem = $list.GetItemById($csvDoc.IdDoc)
            <#$listItem["FechaAprobacion"]=[datetime]::ParseExact($csvDoc.FechaAprobacion,"dd/MM/yyyy",$null)
            $listItem["FechaEdicion"]=[datetime]::ParseExact($csvDoc.FechaEdicion,"dd/MM/yyyy",$null)
            $listItem["CodAntiguo"] = $csvDoc.CodigoAntiguo
            $listItem["CodigoDoc"]  = $csvDoc.CodigoNuevo
            $listItem["EstadoDoc"]  = "Aprobado"
            $listItem["NroEdicion"]  = $csvDoc.NroEdicion
            $listItem["NroRevision"]  = $csvDoc.NroRevision
            $listItem["Pais"]  = "Perú"
            $listItem["Title"]  = $csvDoc.Titulo            
            $lookupGerencia = New-Object Microsoft.SharePoint.Client.FieldLookupValue
            $lookupGerencia.LookupId  = getGerenciaId $csvDoc.Gerencia.Trim() $csvDoc.Pais $itemsGerencias             
            $listItem["GerenciaFuncional"]  = $lookupGerencia.LookupId
            $lookupArea = New-Object Microsoft.SharePoint.Client.FieldLookupValue
            $lookupArea.LookupId  = getAreaId $csvDoc.Area.Trim() $csvDoc.Pais $lookupGerencia.LookupId $itemsAreas 
            $listItem["Area"] = $lookupArea.LookupId
            $lookupTipoDoc = New-Object Microsoft.SharePoint.Client.FieldLookupValue
            $lookupTipoDoc.LookupId = getTipoDocId $csvDoc.TipoDocumento $csvDoc.Pais $lookupArea.LookupId $lookupGerencia.LookupId $itemsTipoDocs
            $listItem["TipoDoc0"]  = $lookupTipoDoc.LookupId
            #$SPOUserValue = New-Object Microsoft.SharePoint.Client.FieldUserValue
            #$SPOUserValue.LookupId = getUserId $csvDoc.'Revisor '
            #$listItem["Revisor"] = $SPOUserValue.LookupId
            #$listItem["Comentario"] = ""
            #>
            $listItem["UltimoCorrelativo"]=$csvDoc.UltimoCorrelativo
            $ctx.load($listItem)
            #$listItem["FechaRevision"] = $null
            $listItem.Update()            
            #$listItem.DeleteObject()    
            $ctx.executeQuery()
        #}
            $i=$i+1  
            
        }
        catch{  
            write-host "$($_.Exception.Message)" -foregroundcolor red  
        }  
    }
}  
catch{  
    write-host "$($_.Exception.Message)" -foregroundcolor red  
}  

#for($i=0;$i -lt $items.Count; $i++){
#    Copy-Item $items[$i] -Destination "D:\Installers\Exsa"
#}
    
#gci "D:\Installers\Docs" -Recurse -Filter *.pptx | % {copy-item $_.FullName -Destination  "D:\Installers\Exsa" -Force -Container }

#dir > list.txt
#Copy-Item "C:\Wabash\Logfiles\mar1604.log.txt" -Destination "C:\Presentation"
