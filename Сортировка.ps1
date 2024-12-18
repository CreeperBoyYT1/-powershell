# Установка пути к рабочему столу
$desktopPath = [Environment]::GetFolderPath('Desktop')

# Проверка существования папки рабочего стола
if (-Not (Test-Path $desktopPath)) {
    Write-Output "Ошибка: папка рабочего стола не найдена!"
    pause
    exit
}

# Базовый путь для перемещения файлов
$basePath = "$desktopPath\ВСЁ ТУТА\2024"

# Карта месяцев
$monthMap = @{
    "01" = "01 январь"
    "02" = "02 февраль"
    "03" = "03 март"
    "04" = "04 апрель"
    "05" = "05 май"
    "06" = "06 июнь"
    "07" = "07 июль"
    "08" = "08 август"
    "09" = "09 сентябрь"
    "10" = "10 октябрь"
    "11" = "11 ноябрь"
    "12" = "12 декабрь"
}

# Функция для расчёта процентного совпадения строк
function Get-Similarity {
    param (
        [string]$string1,
        [string]$string2
    )
    $string1 = $string1.ToLower()
    $string2 = $string2.ToLower()

    $length1 = $string1.Length
    $length2 = $string2.Length
    $maxLength = [Math]::Max($length1, $length2)

    if ($maxLength -eq 0) { return 0 }

    $distance = (New-Object System.Data.DataTable).Compute("LEVENSHTEIN('$string1', '$string2')", $null)
    return [Math]::Round((1 - ($distance / $maxLength)) * 100, 2)
}

# Поиск и обработка файлов с расширениями .xls и .xlsx
Write-Output "Поиск файлов с расширениями .xls и .xlsx на рабочем столе..."
Write-Output "--------------------------------------"

try {
    # Получение списка файлов
    $files = Get-ChildItem -Path $desktopPath -Filter "*.xls" -File
    $files += Get-ChildItem -Path $desktopPath -Filter "*.xlsx" -File

    if ($files.Count -eq 0) {
        Write-Output "Нет файлов с расширениями .xls или .xlsx на рабочем столе."
    } else {
        Write-Output "Найдены файлы:"
        $files | ForEach-Object { Write-Output $_.FullName }

        foreach ($file in $files) {
            $fileName = $file.Name
            $fileBaseName = $file.BaseName
            $creationDate = $file.CreationTime
            $creationMonth = $creationDate.ToString("MM")
            $creationYear = $creationDate.ToString("yyyy")

            if ($creationYear -ne "2024") {
                Write-Output "Файл '$fileName' пропущен, так как год создания не 2024."
                continue
            }

            if (-not $monthMap.ContainsKey($creationMonth)) {
                Write-Output "Ошибка: некорректный месяц в файле '$fileName' (месяц: $creationMonth)."
                continue
            }

            $monthFolder = $monthMap[$creationMonth]
            $monthFolderPath = Join-Path -Path $basePath -ChildPath $monthFolder

            if (-not (Test-Path -Path $monthFolderPath)) {
                New-Item -ItemType Directory -Path $monthFolderPath | Out-Null
                Write-Output "Папка '$monthFolderPath' создана."
            }

            $subfolders = Get-ChildItem -Path $monthFolderPath -Directory
            $destinationFolder = $null
            $maxSimilarity = 0

            foreach ($subfolder in $subfolders) {
                $similarity = Get-Similarity -string1 $fileBaseName -string2 $subfolder.Name
                if ($similarity -ge 50 -and $similarity -gt $maxSimilarity) {
                    $destinationFolder = $subfolder.FullName
                    $maxSimilarity = $similarity
                }
            }

            if (-not $destinationFolder) {
                $destinationFolder = Join-Path -Path $monthFolderPath -ChildPath $fileBaseName.ToLower()
                if (-not (Test-Path -Path $destinationFolder)) {
                    New-Item -ItemType Directory -Path $destinationFolder | Out-Null
                    Write-Output "Создана новая папка: $destinationFolder"
                }
            } else {
                Write-Output "Файл '$fileName' совпал с подпапкой '$destinationFolder' на $maxSimilarity%."
            }

            $destinationPath = Join-Path -Path $destinationFolder -ChildPath $fileName
            Move-Item -Path $file.FullName -Destination $destinationPath
            Write-Output "Файл '$fileName' перемещен в '$destinationPath'."
        }
    }
} catch {
    Write-Output "Произошла ошибка: $_"
}

Write-Output "--------------------------------------"
Write-Output "Обработка завершена."
Write-Output "Нажмите любую клавишу для завершения..."
pause
