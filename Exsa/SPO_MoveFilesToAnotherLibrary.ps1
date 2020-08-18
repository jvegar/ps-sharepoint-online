#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
 #Get Credentials to connect
$userID = "[username]"
$pwd = ConvertTo-SecureString '[password]' -AsPlainText -Force
$Cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  

#Function to Copy a File
function Copy-SPOFile([String]$SiteURL, [String]$SourceFileURL, [String]$TargetFileURL)
{
    Try{
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $Ctx.Credentials = $Cred
      
        #Copy the File
        $MoveCopyOpt = New-Object Microsoft.SharePoint.Client.MoveCopyOptions
        $Overwrite = $True
        [Microsoft.SharePoint.Client.MoveCopyUtil]::CopyFile($Ctx, $SourceFileURL, $TargetFileURL, $Overwrite, $MoveCopyOpt)
        $Ctx.ExecuteQuery()
  
        Write-host -f Green "File Copied Successfully!"
    }
    Catch {
    write-host -f Red "Error Copying the File!" $_.Exception.Message
    }
}
  
#Set Config Parameters
$SiteURL="https://exsanet.sharepoint.com/sites/psad"
#$SourceFileURL="https://exsanet.sharepoint.com/sites/psad/Shared Documents/Discloser Asia.doc"
#$TargetFileURL="https://crescenttech.sharepoint.com/Shared Documents/Discloser Asia.doc"
  

  
#Call the function to Copy the File
Copy-SPOFile $SiteURL https://exsanet.sharepoint.com/sites/psad/libDocVigentes/ALM-P-002.docx https://exsanet.sharepoint.com/sites/psad/libDocPublicados/ALM-P-002.docx




#Read more: http://www.sharepointdiary.com/2017/02/sharepoint-online-copy-file-between-document-libraries-using-powershell.html#ixzz5uhy7pDSg
