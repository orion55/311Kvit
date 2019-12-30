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

	$tmpDir = getTmpDir

	Write-Log -EntryType Information -Message "Разархивация файла $file"
	$argList = "e -y ""$($file.FullName)"" $tmpDir"
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

	$tmpDir = getTmpDir

	Write-Log -EntryType Information -Message "Разархивация файла $file"
	$argList = "e -y ""$($file.FullName)"" $tmpDir"
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
					if ($xml.Файл.Документ.РезОбр -eq 'Сообщение принято') {
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
		if ($311Jur.errFiles.Length -gt 0) {
			[string]$global:kvit311JurArchiveErr = $cur311JurArchive + '\' + 'KVITERR'
			if (!(Test-Path -Path $global:kvit311JurArchiveErr )) {
				New-Item -ItemType directory $global:kvit311JurArchiveErr -Force | out-null
			}
			$msg = $311Jur.errFiles | Move-Item -Destination $kvit311JurArchiveErr -ErrorAction "SilentlyContinue" -Verbose -Force *>&1
			Write-Log -EntryType Information -Message ($msg | Out-String)
			$311Jur.errFiles = @()
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

function sendEmail {
	if ($311Fiz.allCount -gt 0) {
		$body = ''
		if ($311Fiz.sucCount -gt 0) {
			$body += "Пришли успешные подтверждения - $($311Fiz.sucCount) шт.`n"
			$body += "Потверждения находятся в каталоге $kvit311Archive`n"
			$body += "`n"
			$title = "Пришли подтверждения по 311-Физ"
		}
		if ($311Fiz.errCount -gt 0) {
			$body += "Пришли подтверждения с ошибками - $($311Fiz.errCount) шт.`n"
			$body += "Потверждения находятся в каталоге $kvit311ArchiveErr `n"
			$body += "`n"
			$title = "Пришли подтверждения с ошибками по 311-Физ"
		}
		$body += "Всего пришло подтверждений - $($311Fiz.allCount) шт.`n"
		if (Test-Connection $mailServer -Quiet -Count 2) {
			$encoding = [System.Text.Encoding]::UTF8
			Send-MailMessage -To $311mailAddrFiz -Body $body -Encoding $encoding -From $mailFrom -Subject $title -SmtpServer $mailServer -ErrorAction SilentlyContinue
		}
		else {
			Write-Log -EntryType Error -Message "Не удалось соединиться с почтовым сервером $mailServer"
		}
		Write-Log -EntryType Information -Message $body
	}
	if ($311Jur.allCount -gt 0) {
		$body = ''
		if ($311Jur.sucCount -gt 0) {
			$body += "Пришли успешные подтверждения - $($311Jur.sucCount) шт.`n"
			$body += "Потверждения находятся в каталоге $kvit311JurArchive `n"
			$body += "`n"
			$title = "Пришли подтверждения по 311-Юр"
		}
		if ($311Jur.errCount -gt 0) {
			$body += "Пришли подтверждения с ошибками - $($311Jur.errCount) шт.`n"
			$body += "Потверждения находятся в каталоге $kvit311JurArchiveErr `n"
			$body += "`n"
			$title = "Пришли подтверждения с ошибками по 311-Юр"
		}
		$body += "Всего пришло подтверждений - $($311Jur.allCount) шт.`n"
		if (Test-Connection $mailServer -Quiet -Count 2) {
			$encoding = [System.Text.Encoding]::UTF8
			Send-MailMessage -To $311mailAddrJur -Body $body -Encoding $encoding -From $mailFrom -Subject $title -SmtpServer $mailServer -ErrorAction SilentlyContinue
		}
		else {
			Write-Log -EntryType Error -Message "Не удалось соединиться с почтовым сервером $mailServer"
		}
		Write-Log -EntryType Information -Message $body
	}
}

function getTmpDir {
	$tmpDir = "$curDir\tmp"
	if (!(Test-Path -Path $tmpDir )) {
		New-Item -ItemType directory $tmpDir -Force | out-null
	}
	else {
		Remove-Item $tmpDir -Force -Recurse
		New-Item -ItemType directory $tmpDir -Force | out-null
	}
    Set-Location $tmpDir
	return $tmpDir
}