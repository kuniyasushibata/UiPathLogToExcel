Param([string]$logFile,[string]$excelFile,[string]$excelTemplateFile)

#���̑�ps1�t�@�C���̃��[�h
$loadPss = "get_omit_path.ps1","retry_del.ps1"
$loadPss | ForEach-Object{
    $loadPs = $PSScriptRoot +"\" + $_
    . $loadPs
}

#��΃p�X�\���̏ȗ��\���擾
[string]$omitLogFile = getOmitPath $logFile 90

$isUNC=$false
$tmpFile

#UNC(\\)�m�F
if ($logFile.Substring(0,2) -eq "`\`\"){
#if ($logFile.Substring(0,2) -eq "C:"){
    $isUNC=$true
}

$outFile
if ($isUNC -eq $true){
    #tools�t�H���_�Ƀe���|�����t�@�C���쐬�A�O�̂��߃t�@�C�����擪��PID��t�^
    $tmpFile=$PSScriptRoot + "`\" + $PID.ToString() + "_" + [System.IO.Path]::GetFileName($excelFile)
    #Write-Host $tmpFile
    $outFile=$tmpFile
}else{
    $outFile=$excelFile
}
try{
    #�e���v���[�g����R�s�[���Ă���
    copy $excelTemplateFile $outFile

    #�t�@�C���T�C�Y�擾
    Write-Progress -activity $omitLogFile -status "�t�@�C���T�C�Y�擾��" -percentComplete 0  -Id 1 -ParentId 0
    [int64]$numOfTotalBytes=(Get-Item $logFile).Length
}catch{
    Write-Host $logFile �ŗ�O�����A�����𒆎~���܂��B
    Write-Host $error[0]
    exit
}

$isException=$false
try{
    #�G�N�Z������
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($outFile)
    $sheet = $excel.Worksheets.Item("log")

    #�G�N�Z��Range����
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

    #�v���O���X�p
    [int64]$numOfBytes=3 #BOM��
    [int]$prevP=0
    function updateProgress(){
        [int]$p = 100 * $numOfBytes / $numOfTotalBytes
        if ((($p - $script:prevP) -ge 1) -Or ($numOfBytes -eq $numOfTotalBytes)){
            $status = $p.ToString("0") + "%" + "(" + $numOfBytes.ToString("#,0") + "Byte����/" + $numOfTotalBytes.ToString("#,0") + "Byte��)"
            Write-Progress -activity $omitLogFile -status $status -percentComplete $p  -Id 1 -ParentId 0
            $script:prevP = $p
        }
    }

    updateProgress
    getRange

    #log�t�@�C����ǂݍ��݁A1�s������
    Write-Progress -activity $omitLogFile -status "�t�@�C���I�[�v����" -percentComplete 0  -Id 1 -ParentId 0
    Get-Content $logFile | ForEach-Object {
        $s = -1
        $e = 0
        $json = ""
        #�X�y�[�X�ŋ�؂�1�ڂ�Time,2�ڂ�Level
        for($j=1; $j -lt 3; $j++) {
            $e = $_.IndexOf(" ", $s + 1)
            $table[($low + 1), $j] = $_.Substring($s + 1, $e - ($s + 1))
            $s = $e
        }
        #3�ڂ�json
        $json = $_.Substring($s + 1, $_.Length - ($s + 1))
        #json��͂���ƒx���̂�12�����ڂ���","level":"�܂ł���肾��
        $table[($low + 1), 3] = $json.Substring(12, $json.IndexOf("`",`"level`":`"") - 12)

        $low++
        #$writeUnit��Excel��������
        if ($low % $writeUnit -eq 0){
            setRange
            getRange
            $low = 0
            updateProgress
        }    
        $numOfBytes += [System.Text.Encoding]::GetEncoding("utf-8").GetByteCount($_) + 2 #CrLf��
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
    Write-Host ���������~�� $logFile / xlsx�ŗ�O����
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
                Write-Progress -activity $omitLogFile -status "�t�@�C���R�s�[��" -percentComplete 0  -Id 1 -ParentId 0
                copy $tmpFile $excelFile
                Write-Host �������� $excelFile
            }catch{
                Write-Host ���������~�� $logFile / xlsx�ŗ�O����
                Write-Host $error[0]
            }finally{
                try{
                    retryDel $tmpFile 5
                }catch{
                    Write-Host $error[0]
                }
            }
        }else{
            Write-Host �������� $excelFile
        }
    }
}
