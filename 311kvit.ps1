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

#копируем каталоги рекурсивно на "волшебный" диск А: - туда и обратно
function copyDirs {
	Param(
		[string]$from,
		[string]$to)

	Get-ChildItem -Path $from -Recurse |
	Copy-Item -Destination {
		if ($_.PSIsContainer) {
			Join-Path $to $_.Parent.FullName.Substring($from.length)
		}
		else {
			Join-Path $to $_.FullName.Substring($from.length)
		}
	} -Force
}

function init311Fiz {
	$curDate = Get-Date -Format "ddMMyyyy"

	[string]$global:cur311Archive = $311Archive + '\' + $curDate
	if (!(Test-Path -Path $global:cur311Archive )) {
		New-Item -ItemType directory $global:cur311Archive -Force | out-null
	}

	[string]$global:kvit311Archive = $cur311Archive + '\' + 'KVIT'
	if (!(Test-Path -Path $global:kvit311Archive )) {
		New-Item -ItemType directory $global:kvit311Archive -Force | out-null
	}

	[string]$global:kvit311ArchiveMsk = $311DirKvit + '\' + $curDate
	if (!(Test-Path -Path $global:kvit311ArchiveMsk )) {
		New-Item -ItemType directory $global:kvit311ArchiveMsk -Force | out-null
	}

	$global:flag311Fiz = $true
}
function 311FizHandler {
	Param($file)
	Write-Log -EntryType Information -Message "Разархивация файла $file"
	$argList = "e -y $($file.FullName) $tmpDir"
	Start-Process -FilePath $arj32 -ArgumentList $argList -Wait -NoNewWindow

	$tmpFiles = Get-ChildItem "$tmpDir\*.xml"
	[int]$allCount = ($tmpFiles | Measure-Object).Count
	if ($allCount -gt 0) {
		$311Fiz.allCount += $allCount
		[int]$errCount = (Get-ChildItem "$tmpDir\SFE*.xml" | Measure-Object).Count
		$311Fiz.errCount = $errCount
		[int]$sucCount = (Get-ChildItem "$tmpDir\SFF*.xml" | Measure-Object).Count
		$311Fiz.sucCount += $sucCount

		foreach ($curFile in $tmpFiles) {
			$testFile = $curFile.FullName + '.test'
			$arguments = "-verify -delete -1 -profile $profile -registry -in ""$($curFile.FullName)"" -out ""$testFile"" -silent $logSpki"
			Start-Process $spki $arguments -NoNewWindow -Wait
			Write-Log -EntryType Information -Message "Снимаем подпись с файла $($curFile.Name)"

			if (Test-Path $testFile) {
				$msg = Remove-Item $($curFile.FullName) -Verbose -Force *>&1
				Write-Log -EntryType Information -Message ($msg | Out-String)

				$msg = Get-ChildItem $testFile | Rename-Item -NewName { $_.Name -replace '.test$', '' } -Verbose *>&1
				Write-Log -EntryType Information -Message ($msg | Out-String)

				Write-Log -EntryType Information -Message "Форматируем xml-файл $($curFile.Name)"
				[xml]$xml = Get-Content $curFile
				$xml.Save($curFile)
			}
			else {
				$msg = "С файла $($curFile.BaseName) не удалось снять подпись"
				Write-Log -EntryType Error -Message $msg
			}
		}
		if ($sucCount -gt 0) {
			$msg = Copy-Item -Path "$tmpDir\SFF*.xml" -Destination $kvit311Archive -ErrorAction "SilentlyContinue" -Verbose -Force *>&1
			Write-Log -EntryType Information -Message ($msg | Out-String)
		}
		if ($errCount -gt 0) {
			[string]$global:kvit311ArchiveErr = $cur311Archive + '\' + 'KVITERR'
			if (!(Test-Path -Path $kvit311ArchiveErr )) {
				New-Item -ItemType directory $kvit311ArchiveErr -Force | out-null
			}
			$msg = Copy-Item -Path "$tmpDir\SFE*.xml" -Destination $kvit311ArchiveErr -ErrorAction "SilentlyContinue" -Verbose -Force *>&1
			Write-Log -EntryType Information -Message ($msg | Out-String)
		}
		$msg = Move-Item -Path "$tmpDir\SF*.xml" -Destination $kvit311ArchiveMsk -ErrorAction "SilentlyContinue" -Verbose -Force *>&1
		Write-Log -EntryType Information -Message ($msg | Out-String)
	}
	else {
		Write-Log -EntryType Error -Message "Ошибка при разархивация файла $($file.FullName)"
	}
}
function init311Jur {
	$curDate = Get-Date -Format "ddMMyyyy"

	[string]$global:cur311JurArchive = $311JurArchive + '\' + $curDate
	if (!(Test-Path -Path $global:cur311JurArchive )) {
		New-Item -ItemType directory $cur311JurArchive -Force | out-null
	}

	[string]$global:kvit311JurArchive = $cur311JurArchive + '\' + 'KVIT'
	if (!(Test-Path -Path $global:kvit311JurArchive )) {
		New-Item -ItemType directory $global:kvit311JurArchive -Force | out-null
	}

	$global:flag311Jur = $true
}
function 311JurHandler {
	Param($file)

	[string]$typeFile = ''
	if ($file -match $311MaskJur) {
		$typeFile = 'ON'
	}
	elseif ($file -match $311MaskJur2) {
		$typeFile = 'S0'
	}

	Write-Log -EntryType Information -Message "Разархивация файла $file"
	$argList = "e -y $($file.FullName) $tmpDir"
	Start-Process -FilePath $arj32 -ArgumentList $argList -Wait -NoNewWindow

	$tmpFiles = Get-ChildItem "$tmpDir\*.xml"
	[int]$allCount = ($tmpFiles | Measure-Object).Count
	if ($allCount -gt 0) {
		$311Jur.allCount += $allCount
		foreach ($curFile in $tmpFiles) {
			$testFile = $curFile.FullName + '.test'
			$arguments = "-verify -delete -1 -profile $profile -registry -in ""$($curFile.FullName)"" -out ""$testFile"" -silent $logSpki"
			Start-Process $spki $arguments -NoNewWindow -Wait
			Write-Log -EntryType Information -Message "Снимаем подпись с файла $($curFile.Name)"

			if (Test-Path $testFile) {
				$msg = Remove-Item $($curFile.FullName) -Verbose -Force *>&1
				Write-Log -EntryType Information -Message ($msg | Out-String)

				$msg = Get-ChildItem $testFile | Rename-Item -NewName { $_.Name -replace '.test$', '' } -Verbose *>&1
				Write-Log -EntryType Information -Message ($msg | Out-String)

				Write-Log -EntryType Information -Message "Форматируем xml-файл $($curFile.Name)"
				[xml]$xml = Get-Content $curFile
				$xml.Save($curFile)

				if ($typeFile -eq 'ON') {
					if ($xml.Файл.Документ.РезОбр -eq 'Сообщение принято1') {
						$311Jur.sucCount += 1
					}
					else {
						$311Jur.errCount += 1
						$311Jur.errFiles += $curFile
					}
				}
				elseif ($typeFile -eq 'S0') {
					if ($xml.Файл.Документ.Ошибки.КодОшибки -eq '000') {
						$311Jur.sucCount += 1
					}
					else {
						$311Jur.errCount += 1
						$311Jur.errFiles += $curFile
					}
				}
			}
			else {
				$msg = "С файла $($curFile.BaseName) не удалось снять подпись"
				Write-Log -EntryType Error -Message $msg
			}
		}
		if ($311Jur.errCount -gt 0) {
			[string]$global:kvit311JurArchiveErr = $cur311JurArchive + '\' + 'KVITERR'
			if (!(Test-Path -Path $global:kvit311JurArchiveErr )) {
				New-Item -ItemType directory $global:kvit311JurArchiveErr -Force | out-null
			}
			$msg = $311Jur.errFiles | Move-Item -Destination $kvit311JurArchiveErr -ErrorAction "SilentlyContinue" -Verbose -Force *>&1
			Write-Log -EntryType Information -Message ($msg | Out-String)
		}
		if ($311Jur.sucCount -gt 0) {
			$msg = Move-Item -Path "$tmpDir\*.xml" -Destination $kvit311JurArchive -ErrorAction "SilentlyContinue" -Verbose -Force *>&1
			Write-Log -EntryType Information -Message ($msg | Out-String)
		}
	}
	else {
		Write-Log -EntryType Error -Message "Ошибка при разархивация файла $($file.FullName)"
	}
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

$tmpDir = "$curDir\tmp"
if (!(Test-Path -Path $tmpDir )) {
	New-Item -ItemType directory $tmpDir -Force | out-null
}
else {
	Remove-Item $tmpDir -Force -Recurse
	New-Item -ItemType directory $tmpDir -Force | out-null
}

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

#sendEmail -result $result
#Remove-Item $tmpDir -Force -Recurse

Write-Log -EntryType Information -Message "Загружаем исходную ключевую дискету"
Remove-Item 'a:' -Recurse -ErrorAction "SilentlyContinue"
copyDirs -from $tmpKeys -to 'a:'
Remove-Item $tmpKeys -Recurse -Force

Write-Log -EntryType Information -Message "Завершение обработки..."

Stop-FileLog
Stop-HostLog