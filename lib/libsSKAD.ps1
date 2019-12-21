#копируем каталоги рекурсивно на "волшебный" диск А: - туда и обратно
function Copy_dirs {
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
function SKAD_Encrypt {
    Param(
        $encrypt = $false,
        [string]$maskFiles = "*.*")


    $mask = Get-ChildItem -path $work $maskFiles

    foreach ($file in $mask) {
        $tmpFile = $file.FullName + '.test'

        $arguments = ''
        if ($encrypt) {
            $arguments = "-sign -encrypt -profile $profile -registry -algorithm 1.2.643.7.1.1.2.2 -in $($file.FullName) -out $tmpFile -reclist $recList -silent $logSpki"
        }
        else {
            $arguments = "-sign -profile $profile -registry -algorithm 1.2.643.7.1.1.2.2 -data $($file.FullName) -out $tmpFile -reclist $recList -silent $logSpki"
        }

        Write-Log -EntryType Information -Message "Обрабатываем файл $($file.Name)"
        Start-Process $spki $arguments -NoNewWindow -Wait
    }

    $testFiles = Get-ChildItem "$work\*.test"
    if (($testFiles | Measure-Object).count -gt 0) {
        $msg = $mask | Remove-Item -Verbose -Force *>&1
        Write-Log -EntryType Information -Message ($msg | Out-String)
        $msg = Get-ChildItem -path $work '*.test' | Rename-Item -NewName { $_.Name -replace '.test$', '' } -Verbose *>&1
        Write-Log -EntryType Information -Message ($msg | Out-String)
    }
    else {
        Write-Log -EntryType Error -Message "Ошибка при работе программы $spki"
        exit
    }
}
function SKAD_Decrypt {
    Param(
        $decrypt = $false,
        [string]$maskFiles = "*.*")

    Write-Log -EntryType Information -Message "Начинаем преобразование..."
    $mask = Get-ChildItem -path $work $maskFiles

    foreach ($file in $mask) {
        $tmpFile = $file.FullName + '.test'

        $arguments = ''
        if ($decrypt) {
            $arguments = "-decrypt -verify -delete -1 -profile $profile -registry -in $($file.FullName) -out $tmpFile -silent $logSpki"
        }
        else {
            $arguments = "-verify -delete -1 -profile $profile -registry -in $($file.FullName) -out $tmpFile -silent $logSpki"
        }

        Write-Log -EntryType Information -Message "Обрабатываем файл $($file.Name)"
        Start-Process $spki $arguments -NoNewWindow -Wait
    }

    $testFiles = Get-ChildItem "$work\*.test"
    if (($testFiles | Measure-Object).count -gt 0) {
        $msg = $mask | Remove-Item -Verbose -Force *>&1
        Write-Log -EntryType Information -Message ($msg | Out-String)
        $msg = Get-ChildItem -path $work '*.test' | Rename-Item -NewName { $_.Name -replace '.test$', '' } -Verbose *>&1
        Write-Log -EntryType Information -Message ($msg | Out-String)
    }
    else {
        Write-Log -EntryType Error -Message "Ошибка при работе программы $spki"
        exit
    }

}

function SKAD_archive {
    Param([string]$maskFiles = "*.*")

    $arguments = "-f -k $work\$maskFiles"
    Start-Process $archiver $arguments -NoNewWindow -Wait

    $gzFiles = Get-ChildItem "$work\*.gz"
    if (($gzFiles | Measure-Object).count -gt 0) {
        $msg = "$work\$maskFiles" | Remove-Item -Verbose -Force -Exclude "$work\*.gz" *>&1
        Write-Log -EntryType Information -Message ($msg | Out-String)
        $msg = Get-ChildItem -path $work '*.gz' | Rename-Item -NewName { $_.Name -replace '.gz$', '' } -Verbose *>&1
        Write-Log -EntryType Information -Message ($msg | Out-String)
    }
    else {
        Write-Log -EntryType Error -Message "Ошибка при работе программы $archiver"
        exit
    }
}