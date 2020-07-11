Param([string]$Recurse="false")
#log to excel��ps1�̃p�X
$psFile = $PSScriptRoot + "`\" + "log_to_excel.ps1"

#excel�̃e���v���[�g�p�X
$excelTemplateFile = $PSScriptRoot + "`\" + "template.xlsx"

#���̑�ps1�t�@�C���̃��[�h
$ps = $PSScriptRoot +"\" + "is_folder.ps1"
. $ps

#�����Ώۂ�ǉ����郊�X�g
$list = New-Object 'System.Collections.Generic.List[string[]]'

#�ǉ������t�@�C����
$numOfFiles = 0

#�t�@�C���������Ώ�(�g���q�A���{�ς�)�����ׁA���X�g�ɒǉ�
function checkFile([string]$file){
    #Write-Host $file
    $ext = [System.IO.Path]::GetExtension($file)
    if (($ext -eq ".log") -Or ($ext -eq ".LOG")){
        $f = [System.IO.Path]::GetDirectoryName($file) + "`\" + [System.IO.Path]::GetFileNameWithoutExtension($file) + ".xlsx"
        #Write-Host $f
        if (Test-Path $f){
        #if ($false){
            Write-Host "���ΏۊO(�����ς�)��"$file
        }else{
            #Write-Host list added $file
            $script:list.add($file)
            $script:numOfFiles++
        }
    }else{
        Write-Host "���ΏۊO(�g���q)��"$file
    }
}

#�t�H���_�w�肳�ꂽ�ꍇ�A�t�H���_���̃t�@�C��/�t�H���_�𒲂׏�������
#�T�u�t�H���_�܂ނ��w�肳�ꂽ�ꍇ�́ARecurse�I�v�V����(���D��ɂȂ�A
#����͂�����̕������₷��)���g�p�A�����Ŏ��g�ōċA����Ɛ[���D��ɂȂ�
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

#�����`�F�b�N
$args | ForEach-Object{
    if ($_ -ne ""){
        if (isFolder $_ ){
            checkFolder $_
        }else{
            checkFile $_
        }
    }
}

#�������ԑ���R�}���h���b�g������ł���
Measure-Command {
    #�����t�@�C����
    $i = 1
    
    #log�t�@�C�����̏���
    foreach($a in $list){
        $status = $i.ToString() + "�t�@�C���ڏ�����" + "/" + $numOfFiles.ToString() + "�t�@�C��"
        [int]$p = 100 * $i / $numOfFiles
        Write-Progress -activity "��" -status $status -percentComplete $p -Id 0
        $logFile = $a
        $excelFile = [System.IO.Path]::GetDirectoryName($logFile) + "`\" + [System.IO.Path]::GetFileNameWithoutExtension($logFile) + ".xlsx"
        #Write-Host $logFile $excelFile
        &$psFile $logFile $excelFile $excelTemplateFile
        $i++
    }
    Write-Host ��������
}
