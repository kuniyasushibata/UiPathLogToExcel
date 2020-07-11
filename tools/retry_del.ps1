#ファイル削除のretry版
function retryDel([string]$path,[int]$numOfRetries){
    $isContinue=$true
    $retryCount=0;
    
    do{
        try{
            del $path -ErrorAction:Stop
            $isContinue=$false
        }catch{
            if ($retryCount -eq $numOfRetries){
                $isContinue=$false
                Throw $_.Exception
            }
            #Write-Host $Error[0]
            Start-Sleep -Seconds 1
            $retryCount+=1
        }
    }while($isContinue)
}
