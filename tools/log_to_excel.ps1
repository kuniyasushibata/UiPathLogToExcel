Param([string]$logFile,[string]$excelFile,[string]$excelTemplateFile)

#その他ps1ファイルのロード
$loadPss = "get_omit_path.ps1","retry_del.ps1"
$loadPss | ForEach-Object{
    $loadPs = $PSScriptRoot +"\" + $_
    . $loadPs
}

#絶対パス表示の省略表示取得
[string]$omitLogFile = getOmitPath $logFile 90

$isUNC=$false
$tmpFile

#UNC(\\)確認
if ($logFile.Substring(0,2) -eq "`\`\"){
#if ($logFile.Substring(0,2) -eq "C:"){
    $isUNC=$true
}

$outFile
if ($isUNC -eq $true){
    #toolsフォルダにテンポラリファイル作成、念のためファイル名先頭にPIDを付与
    $tmpFile=$PSScriptRoot + "`\" + $PID.ToString() + "_" + [System.IO.Path]::GetFileName($excelFile)
    #Write-Host $tmpFile
    $outFile=$tmpFile
}else{
    $outFile=$excelFile
}
try{
    #テンプレートからコピーしておく
    copy $excelTemplateFile $outFile

    #ファイルサイズ取得
    Write-Progress -activity $omitLogFile -status "ファイルサイズ取得中" -percentComplete 0  -Id 1 -ParentId 0
    [int64]$numOfTotalBytes=(Get-Item $logFile).Length
}catch{
    Write-Host $logFile で例外発生、処理を中止します。
    Write-Host $error[0]
    exit
}

$isException=$false
try{
    #エクセル操作
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($outFile)
    $sheet = $excel.Worksheets.Item("log")

    #エクセルRange操作
    $startCol="A"
    $endCol="C"
    $writeUnit=2000
    $startRow=2
    $endRow=$startRow + $writeUnit - 1
    $low = 0
    $range

    function getRange(){
        $script:range = $startCol + $startRow.ToString() + ":" + $endCol + $endRow.ToString()
        #Write-Host "get range" $script:range
        $script:table = $sheet.Cells.Range($range).Value2
    }

    function setRange(){
        #Write-Host "set range" $script:range
        $sheet.Cells.Range($range) = $script:table
        $script:startRow = $script:endRow + 1
        $script:endRow=$script:startRow + $script:writeUnit - 1
    }

    #プログレス用
    [int64]$numOfBytes=3 #BOM分
    [int]$prevP=0
    function updateProgress(){
        [int]$p = 100 * $numOfBytes / $numOfTotalBytes
        if ((($p - $script:prevP) -ge 1) -Or ($numOfBytes -eq $numOfTotalBytes)){
            $status = $p.ToString("0") + "%" + "(" + $numOfBytes.ToString("#,0") + "Byte完了/" + $numOfTotalBytes.ToString("#,0") + "Byte中)"
            Write-Progress -activity $omitLogFile -status $status -percentComplete $p  -Id 1 -ParentId 0
            $script:prevP = $p
        }
    }

    updateProgress
    getRange

    #logファイルを読み込み、1行ずつ処理
    Write-Progress -activity $omitLogFile -status "ファイルオープン中" -percentComplete 0  -Id 1 -ParentId 0
    Get-Content $logFile | ForEach-Object {
        $s = -1
        $e = 0
        $json = ""
        #スペースで区切り1つ目がTime,2つ目がLevel
        for($j=1; $j -lt 3; $j++) {
            $e = $_.IndexOf(" ", $s + 1)
            $table[($low + 1), $j] = $_.Substring($s + 1, $e - ($s + 1))
            $s = $e
        }
        #3つ目がjson
        $json = $_.Substring($s + 1, $_.Length - ($s + 1))
        #json解析すると遅いので12文字目から","level":"までを取りだす
        $table[($low + 1), 3] = $json.Substring(12, $json.IndexOf("`",`"level`":`"") - 12)

        $low++
        #$writeUnit毎Excel書き込み
        if ($low % $writeUnit -eq 0){
            setRange
            getRange
            $low = 0
            updateProgress
        }    
        $numOfBytes += [System.Text.Encoding]::GetEncoding("utf-8").GetByteCount($_) + 2 #CrLf分
    }

    if ($low % $writeUnit -ne 0){
        setRange
    }
    if ($numOfBytes -gt $numOfTotalBytes){
        $numOfBytes = $numOfTotalBytes
    }
    updateProgress

    $book.Save()
}catch{
    $isException=$true
    Write-Host ■処理中止■ $logFile / xlsxで例外発生
    Write-Host $error[0]
}finally{
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    if ($isException -eq $true){
        try{
            del $outFile -ErrorAction:Stop
        }catch{
            Write-Host $error[0]
        }
    }else{
        if ($isUNC -eq $true){
            try{
                Write-Progress -activity $omitLogFile -status "ファイルコピー中" -percentComplete 0  -Id 1 -ParentId 0
                copy $tmpFile $excelFile
                Write-Host ■完了■ $excelFile
            }catch{
                Write-Host ■処理中止■ $logFile / xlsxで例外発生
                Write-Host $error[0]
            }finally{
                try{
                    retryDel $tmpFile 5
                }catch{
                    Write-Host $error[0]
                }
            }
        }else{
            Write-Host ■完了■ $excelFile
        }
    }
}
