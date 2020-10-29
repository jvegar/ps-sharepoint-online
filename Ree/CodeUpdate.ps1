$SCPrefixDict = @{ 
"https://normamto.ree.es" = "E-"; 
"https://normareamto.ree.es" = "I-"; 
"https://normareintelmto.ree.es" = "T-"; 
"https://normareincanmto.ree.es" = "C-"; 
"https://normacorpmto.ree.es" = "G-"; 
}

$SCLists = @("Normativa","MovimientosNormativa","Pyramid")

$ListsInternalNames = @("<nombre interno 1>","<nombre interno 2>","<nombre interno3>")

foreach($url in $SCPrefixDict.Keys){
    Write-Host "Accediendo al sitio: " $url;
    $web = Get-SPWeb $url;
    foreach($listName in $SCLists){
        Write-Host "Accediendo a la lista: " $listName;
        $list = $web.Lists[$listName];
        foreach($item in $list.Items){
            foreach($internalName in $ListsInternalNames){
                if($item[$internalName].ToString().StartsWith("T")){
                    $item[$internalName]=$item[$internalName].ToString().Replace("T","C");
                    $item[$internalName]=$SCPrefixDict[$url]+"-"+$item[$internalName];
                    $item.Update();
                }
            }
        }
    }      
}