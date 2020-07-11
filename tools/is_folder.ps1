#フォルダ/ファイルに判定
function isFolder([string]$path){
    Get-Item $path | ForEach-Object {
        if($_.PSIsContainer)
        {
            return $true
        }
        else
        {
            return $false
        }
    }
}