Param([string]$Recurse="false")
#log to excelのps1のパス
$psFile = $PSScriptRoot + "`\" + "log_to_excel.ps1"

#excelのテンプレートパス
$excelTemplateFile = $PSScriptRoot + "`\" + "template.xlsx"

#その他ps1ファイルのロード
$ps = $PSScriptRoot +"\" + "is_folder.ps1"
. $ps

#処理対象を追加するリスト
$list = New-Object 'System.Collections.Generic.List[string[]]'

#追加したファイル数
$numOfFiles = 0

#ファイルが処理対象(拡張子、実施済み)が調べ、リストに追加
function checkFile([string]$file){
    #Write-Host $file
    $ext = [System.IO.Path]::GetExtension($file)
    if (($ext -eq ".log") -Or ($ext -eq ".LOG")){
        $f = [System.IO.Path]::GetDirectoryName($file) + "`\" + [System.IO.Path]::GetFileNameWithoutExtension($file) + ".xlsx"
        #Write-Host $f
        if (Test-Path $f){
        #if ($false){
            Write-Host "■対象外(処理済み)■"$file
        }else{
            #Write-Host list added $file
            $script:list.add($file)
            $script:numOfFiles++
        }
    }else{
        Write-Host "■対象外(拡張子)■"$file
    }
}

#フォルダ指定された場合、フォルダ内のファイル/フォルダを調べ処理する
#サブフォルダ含むを指定された場合は、Recurseオプション(幅優先になる、
#今回はこちらの方が見やすい)を使用、ここで自身で再帰すると深さ優先になる
function checkFolder([string]$folder){
    if($Recurse -eq "true"){
        $items = Get-ChildItem -Recurse $folder
    }else{
        $items = Get-ChildItem $folder
    }
    foreach($item in $items){
        if (isFolder $item.FullName){
            #checkFolder $item.FullName
        }else{
            checkFile $item.FullName
        }
    }
}

#引数チェック
$args | ForEach-Object{
    if ($_ -ne ""){
        if (isFolder $_ ){
            checkFolder $_
        }else{
            checkFile $_
        }
    }
}

#処理時間測定コマンドレットを挟んでおく
Measure-Command {
    #処理ファイル数
    $i = 1
    
    #logファイル毎の処理
    foreach($a in $list){
        $status = $i.ToString() + "ファイル目処理中" + "/" + $numOfFiles.ToString() + "ファイル"
        [int]$p = 100 * $i / $numOfFiles
        Write-Progress -activity "状況" -status $status -percentComplete $p -Id 0
        $logFile = $a
        $excelFile = [System.IO.Path]::GetDirectoryName($logFile) + "`\" + [System.IO.Path]::GetFileNameWithoutExtension($logFile) + ".xlsx"
        #Write-Host $logFile $excelFile
        &$psFile $logFile $excelFile $excelTemplateFile
        $i++
    }
    Write-Host 処理時間
}
