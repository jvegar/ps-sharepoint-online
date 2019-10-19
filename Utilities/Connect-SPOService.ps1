$adminUPN="jose.vega@contosoperu01.onmicrosoft.com"
$orgName="contosoperu01"
$userCredential = Get-Credential -UserName $adminUPN -Message "Type the password."
Connect-SPOService -Url https://$orgName-admin.sharepoint.com -Credential $userCredential