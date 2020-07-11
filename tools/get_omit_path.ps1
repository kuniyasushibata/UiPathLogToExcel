#プログレスの表示幅に限りがあるため、\\...\aaa\bbb.xlsxのような省略形式を返す
function getOmitPath([string]$p, [int]$n){
    #if ([System.Text.Encoding]::GetEncoding("utf-8").GetByteCount($p) -le $n){
    $width = Get-StringWidth($p)
    if ($width -le $n){
        return $p
    }

    $startSize = 5 + 3 #先頭5文字 "..."
    $splitted = $p.Split("`\")
    $count = $startSize
    $index = 0
    $result=""
    for($i = $splitted.Length - 1; $i -ge 0; $i--){
        #$count += [System.Text.Encoding]::GetEncoding("utf-8").GetByteCount($splitted[$i]) + 1
        $count += Get-StringWidth($splitted[$i]) + 1
        #Write-Host $i $splitted[$i] $count
        if ($count -gt $n){
            #Write-Host $i
            $index = $i
            break
        }
    }
    $result = $p.Substring(0, $startSize) + "..."
    for($i = $index; $i -lt $splitted.Length; $i++){
        $result += "`\" + $splitted[$i]
    }
    return $result
}

Add-Type -AssemblyName "Microsoft.VisualBasic"

Function Get-StringWidth([String]$String) {
    $wide  = [Microsoft.VisualBasic.Strings]::StrConv($String,[Microsoft.VisualBasic.VbStrConv]::Wide)
    $width = 0
    for ($i=0; $i -lt $String.Length; $i++) {
        $width++
        if ($String[$i] -eq $wide[$i]) { $width++ }
    }
    return $width
}
