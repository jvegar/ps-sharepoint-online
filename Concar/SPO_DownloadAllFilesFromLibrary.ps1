#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
 
Function Download-AllFilesFromLibrary()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.Folder] $SourceFolder,
        [Parameter(Mandatory=$true)] [string] $TargetFolder
    )
    

        $missingFiles = Import-Csv -Path "E:\Repos\Powershell.SharePointOnline\Concar\Data\concesiones.csv" -Delimiter ";" -Encoding Default
        foreach($mf in $missingFiles){
            $rutaRelativa = "/" + $mf.'Ruta de acceso'
            $rutaCompleta = "/" + $mf.'Ruta de acceso'+"/"+$mf.Nombre
            $FolderName =  $rutaRelativa -replace "/","\"
            $LocalFolder = $TargetFolder + $FolderName
            #$LocalFolder = $LocalFolder.Replace("F:\\sitios\cd-cs-canchaque\03Construcc\","G:\\")
            If (!(Test-Path -Path $LocalFolder)) {
                    New-Item -ItemType Directory -Path $LocalFolder | Out-Null
            }
            if($mf.'Tipo de elemento' -eq "Item"){

                Try 
                {
                    $TargetFile = $LocalFolder+"\"+$mf.Nombre
                    #$TargetFile = $TargetFile.Replace("F:\\sitios\cd-cs-canchaque\03Construcc\","G:\\")
                    #Download the file
                    If (!(Test-Path -Path $TargetFile)) {
                    $FileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Ctx,$rutaCompleta)
                    #$WriteStream = [System.IO.File]::Open($TargetFile,[System.IO.FileMode]::Create)
                    $WriteStream = [System.IO.File]::Create($TargetFile)
                    $FileInfo.Stream.CopyTo($WriteStream)
                    $WriteStream.Close()
                    write-host -f Red "Downloaded File:"$TargetFile
                    }
                    else{
                    write-host -f Green "File already downloaded:"$TargetFile
                    }
                }Catch 
                {
                    write-host -f Red "Error Downloading Files from Library!" $_.Exception.Message
                }
            }
        }        
       

}
 
#Set parameter values
$SiteURL="http://intranetinfra.granaymontero.com.pe/sitios/cd"
$LibraryName="Control de gestión - Concesiones"
$TargetFolder="E:\CG"
 
#Setup Credentials to connect
$Credentials = New-Object System.Net.NetworkCredential("serviciosti", "s1st3m42020$")
 
#Setup the context
$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
$Ctx.Credentials = $Credentials
$Ctx.RequestTimeout = -1
      
#Get the Library
$List = $Ctx.Web.Lists.GetByTitle($LibraryName)
$Ctx.Load($List)
$Ctx.Load($List.RootFolder)
$Ctx.ExecuteQuery()
 
 
#Call the function: sharepoint online download multiple files powershell
Download-AllFilesFromLibrary -SiteURL $SiteURL -SourceFolder $List.RootFolder -TargetFolder $TargetFolder


#Read more: https://www.sharepointdiary.com/2017/03/sharepoint-online-download-all-files-from-document-library-using-powershell.html#ixzz69tdgxNlq