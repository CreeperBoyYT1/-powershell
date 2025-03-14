PowerShell Script для организации файлов на рабочем столе.

Этот скрипт предназначен для автоматической организации файлов с расширениями .xls и .xlsx, находящихся на рабочем столе, в папки по месяцам и клиентам внутри заранее заданной структуры.

Как работает скрипт:
1. Поиск файлов: Скрипт сканирует рабочий стол пользователя на наличие файлов с расширениями .xls и .xlsx.


2. Определение папки месяца: Для каждого найденного файла определяется папка месяца в соответствии с датой создания файла.

3. Поиск подпапки клиента:
Если в папке месяца есть подпапка с именем, схожим с именем файла (совпадение 50% и выше), файл перемещается туда.

Если подходящей подпапки нет, создаётся новая подпапка на основе имени файла.

4. Перемещение файла: Файл перемещается в определённую подпапку.

Требования:
1. PowerShell: Скрипт предназначен для работы в Windows PowerShell.

2. Структура папок: Скрипт предполагает, что файлы будут перемещаться в директорию ВСЁ ТУТА/2024, которая находится на рабочем столе.

Установка и запуск:
1. Скачайте или скопируйте код скрипта в файл "Сортировка".

2. Убедитесь, что структура папок соответствует следующему формату:

Рабочий стол/
├── ВСЁ ТУТА/
    ├── 2024/
        ├── 01 январь/
        ├── 02 февраль/
        ├── ... (другие месяцы)

Если папки отсутствуют, скрипт создаст их автоматически.

3. Откройте PowerShell и перейдите в директорию, где находится скрипт:
   
cd "Путь/к/файлу"

5. Запустите скрипт:

.\Сортировка.ps1

Настройки скрипта
Изменение базового пути

Если вы хотите перемещать файлы в другую папку, измените значение переменной $basePath:

$basePath = "Ваш/путь/к/папке"

Изменение карты месяцев
Если необходимо изменить названия папок месяцев, отредактируйте переменную $monthMap:

$monthMap = @{
    "01" = "Январь"
    "02" = "Февраль"
    ...
}

Порог совпадения для подпапок

По умолчанию скрипт ищет подпапку с совпадением имени файла на 50% и выше. Чтобы изменить этот порог, отредактируйте значение в условии:

if ($similarity -ge 50) { ... }

Например, для 70%:
if ($similarity -ge 70) { ... }

Возможные ошибки и их решение
1. Папка рабочего стола не найдена:
Убедитесь, что ваш рабочий стол находится в стандартной директории.

Если у вас нестандартный путь к рабочему столу, замените следующую строку:

$desktopPath = [Environment]::GetFolderPath('Desktop')

Например:
$desktopPath = "C:\Пользователи\ВашПользователь\Рабочий стол"

2. Ошибка в обработке файлов:
Проверьте, что файлы .xls или .xlsx действительно находятся на рабочем столе.

Примечания:
Скрипт не обрабатывает файлы, дата создания которых не соответствует 2024 году.

При совпадении имён учитывается регистронезависимое сравнение.

---

Эта инструкция поможет любому пользователю настроить и запустить скрипт без лишних вопросов.

