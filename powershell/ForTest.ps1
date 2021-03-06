#
#2017-10-24  By ral
# 更新说明：
# 1、运行前需将 Ionic.Zip.dll  复制到 C:\Windows\System32 目录下
# 2、免去输入步骤，默认 密码文件、压缩文件 与当前脚本是在同一路径下

#提示
$ws = New-Object -ComObject WScript.Shell 

#默认当前路径
$gl = Get-Location 
$pwdPath = $gl.Path.ToString()

#载入dll
[System.Reflection.Assembly]::LoadFrom($env:windir+"\system32\Ionic.Zip.dll") >$null
#判断是否存在
if(-not (Test-Path ($pwdPath + "\密码.txt"))){
    $n1=$ws.popup("找不到文件：密码.txt ！",0,"提示",0 + 64)
}
else{
    #默认当前位置
    $souPath = $gl.Path.ToString()
    $savePath = $souPath + "\解压"
    if(-not (Test-Path $savePath)){
        mkdir $savePath >$null
    }
   
    # 读取带密码文件内容
    $lines = Get-Content -Path ($pwdPath + "\密码.txt")
    #按行读取
    foreach ($line in $lines){
        #分隔文件名和密码
        $items = [regex]::split($line, '[\t ]+')
        if ($items.Length -gt 1)
        {     
            #拼接zip文件全路径
            $SourceFile = $souPath + "\" + $items[0] + ".zip" 
            #进行解压
            $zip = [Ionic.Zip.ZipFile]::Read($SourceFile)
            $zip.Password = $items[1]
            $zip.ExtractExistingFile= [Ionic.Zip.ExtractExistingFileAction]::OverwriteSilently #覆盖文件
            $zip.ExtractAll($savePath)
            $zip = $null
        }
    }
    #完成，是否合并
    $n2=$ws.popup("解压完成！是否需要合并成为一个CSV文件？",0,"提示",1 + 64)
    if ($n2 -eq 1 ){
       Get-Content -Path ($savePath + "\*.csv") | Out-File ($souPath + "\合并后.csv") -Encoding default >$null
       $n3=$ws.popup("合并完成！",0,"提示",0 + 64)
    }
}
$gl = $null
$ws = $null
