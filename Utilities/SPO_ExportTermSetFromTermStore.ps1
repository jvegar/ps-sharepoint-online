#Adding Sharepoint Online Client libraries 
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll"

#Setting global variables
$siteURL = "https://contosoperu01.sharepoint.com/sites/Labs/"
$userID = "jose.vega@contosoperu01.onmicrosoft.com"
$pwd = ConvertTo-SecureString 'P2ssw0rd01' -AsPlainText -Force
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)  
$ctx.credentials = $creds


$GroupName = Read-Host -Prompt "Enter the metadata group name" 
$TermSetName = Read-Host -Prompt "Enter the termset name"

#Recursive function to get terms

function GetTerms([Microsoft.SharePoint.Client.Taxonomy.Term] $Term,[String]$ParentTerm,[int] $Level)
{
  $Terms = $Term.Terms;
  $Context.Load($Terms)
  $Context.ExecuteQuery();
  if($ParentTerm)
  {
   $ParentTerm = $ParentTerm + "," + $Term.Name;
  }
  else
  {
   $ParentTerm = $Term.Name;
  }

  Foreach ($SubTerm in $Terms)
  {
     $Level = $Level + 1;
     #up to 7 terms levels are written
     $NumofCommas =  7 - $Level;
     $commas ="";

     For ($i=0; $i -lt $NumofCommas; $i++)  
     {
        $commas = $commas + ",";
     }

    $file.Writeline("," + "," + "," + "," + $Term.Description + "," + $ParentTerm + "," + $SubTerm.Name + $commas );
     GetTerms -Term $SubTerm -ParentTerm $ParentTerm -Level $Level;
  } 
}


$MMS = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)

$ctx.Load($MMS)

$ctx.ExecuteQuery()
 
#Get Term Stores

$TermStores = $MMS.TermStores

$ctx.Load($TermStores)

$ctx.ExecuteQuery()
 
$TermStore = $TermStores[0]

$ctx.Load($TermStore)
$ctx.ExecuteQuery()

#Get Groups
$Group = $TermStore.Groups.GetByName($GroupName)
$ctx.Load($Group)
$ctx.ExecuteQuery()


#Bind to Term Set
$TermSet = $Group.TermSets.GetByName($TermSetName)
$ctx.Load($TermSet)
$ctx.ExecuteQuery() 
 
#Create the file and add headings
$OutputFile = "Output File Path1.csv"
$file = New-Object System.IO.StreamWriter($OutputFile)

$file.Writeline("Term Set Name,Term Set Description,LCID,Available for Tagging,Term Description,Level 1 Term,Level 2 Term,Level 3 Term,Level 4 Term,Level 5 Term,Level 6 Term,Level 7 Term");
 
$Terms = $TermSet.Terms
$ctx.Load($Terms);
$ctx.ExecuteQuery();
$lineNum = 1;
Foreach ($Term in $Terms)
{  
  if($lineNum -eq 1)
  {
   ##output term properties on first line only
   $file.Writeline($TermSet.Name + "," + $TermSet.Description + "," + $TermStore.DefaultLanguage + "," + $TermSet.IsAvailableForTagging + "," + $Term.Description + "," + $Term.Name + "," + "," + "," + "," + "," + "," );
  }
  else
  {
    $file.Writeline("," + "," + "," + "," + $Term.Description + "," + $Term.Name + "," + "," + "," + "," + "," + "," );
  }
  $lineNum = $lineNum + 1;
  $TermTreeLevel  = 1; 
  GetTerms -Term $Term -Level $TermTreeLevel -ParentTerm "";
}

 $file.Flush();
$file.Close(); 