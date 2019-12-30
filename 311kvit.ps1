#Программа проверки квитанций по форме 311 для юр. лиц
#от 21.12.2019

#текущий путь
$curDir = Split-Path -Path $myInvocation.MyCommand.Path -Parent

[string]$lib = "$curDir\lib"
. $curDir/variables.ps1
. $lib/PSMultiLog.ps1
. $lib/libs.ps1
. $lib/311lib.ps1

#Set-Location $curDir
Clear-Host

[boolean]$flag311Fiz = $false
[string]$cur311Archive = ''
[string]$kvit311Archive = ''
[string]$kvit311ArchiveErr = ''
[string]$kvit311ArchiveMsk = ''
[hashtable]$311Fiz = @{
	errCount = 0
	sucCount = 0
	allCount = 0
}

[boolean]$flag311Jur = $false
[string]$cur311JurArchive = ''
[string]$kvit311JurArchive = ''
[string]$kvit311JurArchiveErr = ''
[hashtable]$311Jur = @{
	errCount = 0
	sucCount = 0
	allCount = 0
	errFiles = @()
}

Start-HostLog -LogLevel Information
Start-FileLog -LogLevel Information -FilePath $logName -Append

testDir(@($noticePath))
createDir(@($logPath))
testFiles(@($arj32))

if ($debug) {
	Remove-Item -Path "$noticePath\*.*"
	Copy-Item -Path "$curDir\temp\OUT1\*.*" -Destination $noticePath
}

$findFiles = Get-ChildItem -Path $noticePath | Where-Object { ! $_.PSIsContainer } | Where-Object { ($_.Name -match $311MaskFiz) -or ($_.Name -match $311MaskJur) -or ($_.Name -match $311MaskJur2) }
$count = ($findFiles | Measure-Object).count
if ($count -eq 0) {
	exit
}

Write-Log -EntryType Information -Message "Начинаем обработку..."

#проверяем есть ли диск А
$disks = (Get-PSDrive -PSProvider FileSystem).Name
if ($disks -notcontains "a") {
	Write-Log -EntryType Error -Message "Диск А не найден!"
	exit
}

#сохраняем текущею ключевую дискету
Write-Log -EntryType Information -Message "Сохраняем текущею ключевую дискету"
$tmpKeys = "$curDir\tmp_keys"
if (!(Test-Path $tmpKeys)) {
	New-Item -ItemType directory -Path $tmpKeys | out-Null
}
copyDirs -from 'a:' -to $tmpKeys
Remove-Item 'a:' -Recurse -ErrorAction "SilentlyContinue"

Write-Log -EntryType Information -Message "Загружаем ключевую дискету $vdkeys"
copyDirs -from $vdkeys -to 'a:'

ForEach ($file in $findFiles) {
	if ($file -match $311MaskFiz) {
		if (!$flag311Fiz) {
			init311Fiz
		}
		311FizHandler -file $file

		$msg = Move-Item -Path $($file.FullName) -Destination $cur311Archive -ErrorAction "SilentlyContinue" -Verbose -Force *>&1
		Write-Log -EntryType Information -Message ($msg | Out-String)
	}
	if (($file -match $311MaskJur) -or ($file -match $311MaskJur2)) {
		if (!$flag311Jur) {
			init311Jur
		}
		311JurHandler -file $file

		$msg = Move-Item -Path $($file.FullName) -Destination $cur311JurArchive -ErrorAction "SilentlyContinue" -Verbose -Force *>&1
		Write-Log -EntryType Information -Message ($msg | Out-String)
	}
}

sendEmail
#Remove-Item $tmpDir -Force -Recurse

Write-Log -EntryType Information -Message "Загружаем исходную ключевую дискету"
Remove-Item 'a:' -Recurse -ErrorAction "SilentlyContinue"
copyDirs -from $tmpKeys -to 'a:'
#Remove-Item $tmpKeys -Recurse -Force

Write-Log -EntryType Information -Message "Завершение обработки..."

Stop-FileLog
Stop-HostLog