#Программа проверки квитанций по форме 311 для юр. лиц
#от 21.12.2019

#текущий путь
$curDir = Split-Path -Path $myInvocation.MyCommand.Path -Parent

[string]$lib = "$curDir\lib"
. $curDir/variables.ps1
. $lib/PSMultiLog.ps1
. $lib/libs.ps1

Set-Location $curDir
Clear-Host

[boolean]$flag311Fiz = $false
[boolean]$flag311Jur = $false
[string]$cur311Archive = ''
[string]$kvit311Archive = ''
[string]$kvit311ArchiveMsk = ''
function init311Fiz {
	$curDate = Get-Date -Format "ddMMyyyy"

	[string]$global:cur311Archive = $311Archive + '\' + $curDate
	if (!(Test-Path -Path $cur311Archive )) {
		New-Item -ItemType directory $cur311Archive -Force | out-null
	}

	[string]$global:kvit311Archive = $cur311Archive + '\' + 'KVIT'
	if (!(Test-Path -Path $kvit311Archive )) {
		New-Item -ItemType directory $kvit311Archive -Force | out-null
	}

	[string]$global:kvit311ArchiveMsk = $311DirKvit + '\' + $curDate
	if (!(Test-Path -Path $kvit311ArchiveMsk )) {
		New-Item -ItemType directory $kvit311ArchiveMsk -Force | out-null
	}

	$global:flag311Fiz = $true
}

Start-HostLog -LogLevel Information
Start-FileLog -LogLevel Information -FilePath $logName -Append

testDir(@($noticePath))
createDir(@($logPath))

if ($debug) {
	Remove-Item -Path "$noticePath\*.*"
	Copy-Item -Path "$curDir\OUT1\*.*" -Destination $noticePath
}

$findFiles = Get-ChildItem -Path $noticePath | Where-Object { ! $_.PSIsContainer } | Where-Object { ($_.Name -match $311MaskFiz) -or ($_.Name -match $311MaskJur) -or ($_.Name -match $311MaskJur2) }
$count = ($findFiles | Measure-Object).count
if ($count -eq 0) {
	exit
}

Write-Log -EntryType Information -Message "Начинаем обработку..."

$tmpDir = "$curDir\tmp"
if (!(Test-Path -Path $tmpDir )) {
	New-Item -ItemType directory $tmpDir -Force | out-null
}
else {
	Remove-Item $tmpDir -Force -Recurse
	New-Item -ItemType directory $tmpDir -Force | out-null
}

ForEach ($file in $findFiles) {
	if ($file -match $311MaskFiz) {
		if (!$flag311Fiz) {
			init311Fiz
		}
		#$result = 440Handler -file $file
	}
	if (($file -match $311MaskJur) -or ($file -match $311MaskJur2)) {
		#$result = 311Handler -file $file

	}
	#sendEmail -result $result
}

Remove-Item $tmpDir -Force -Recurse

Write-Log -EntryType Information -Message "Завершение обработки..."

Stop-FileLog
Stop-HostLog