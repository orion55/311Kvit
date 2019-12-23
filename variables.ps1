[string]$curDir = Split-Path -Path $myInvocation.MyCommand.Path -Parent
[boolean]$debug = $true

#рабочий временный каталог
[string]$tmp = "$curDir\temp"
[string]$arj32 ="$curDir\util\arj32.exe"

#каталог с извещениями
[string]$noticePath = "$tmp\OUT"

#настройка почты
[string]$311mailAddrFiz = "tmn-goe@tmn.apkbank.ru"
#[string]$311mailAddrFiz = "tmn-f311@tmn.apkbank.apk"
$311mailAddrJur = "tmn-goe <tmn-goe@tmn.apkbank.ru>", "lma <lma@tmn.apkbank.ru>"
#[string[]]$311mailAddrJur = "<tmn-lov@tmn.apkbank.ru>", "<tmn_oit@tmn.apkbank.apk>"

[string]$mailServer = "191.168.6.50"
[string]$mailFrom = "atm_support@tmn.apkbank.apk"

#входящие - настройки
[string]$311MaskFiz = "^NN.+\.arj$"
[string]$311MaskJur = "^ON.+\.arj$"
[string]$311MaskJur2 = "^S.+\.arj$"

[string]$311Archive = "$tmp\311p\Arhive"
[string]$311DirKvit = "$tmp\311pMsk"
[string]$311JurArchive = "$tmp\311jur\Arhive"

$curDate = Get-Date -Format "ddMMyyyy"
[string]$logPath = "$curDir\log"
#имя лог-файла
[string]$logName = $logPath + '\' + $curDate + "_kvit.log"

[string]$spki = "C:\Program Files\MDPREI\spki\spki1utl.exe"
[string]$vdkeys = "d:\SKAD\Floppy\foiv"
[string]$profile = "r2880_2"
[string]$logSpki = $curDir + "\log\" + $curDate + "_spki_tr.log"