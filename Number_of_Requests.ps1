<#
# *************************************************************
# 
# Author: <Gusev Alexandr>
# Creation date: <08/07/2024>
# Description: <Скрипт для подсчета числа запросов>
# Заявка: <XXX0000000>
# История изменений:
# 1) Создание скрипта согласно ТЗ
# 
# ************************************************************* 
#>

<# *************************************************************
# Global variable section
#>

# Загрузка настроек из конфигурационного файла
$configFile = Join-Path -Path $PSScriptRoot -ChildPath "config.json"   # Получение пути файла конфигурации
$config = Get-Content -Path $configFile | ConvertFrom-Json

# Настройки
[string]$logFilePath = $config.logFilePath           # Путь к файлу лога
[string]$tempDir = $config.tempDir                   # Временный каталог для распаковки

[string]$archiveBaseDir = $config.archiveDir         # Каталог с архивами
#[string]$archiveBaseDir = $config.archiveDirTest     # Каталог с архивами тестовый
[string]$matchString = $config.matchString           # Строка для поиска

[string]$smtpServer = $config.smtpServer             # Почтовый сервер
[string]$smtpUsername = $config.smtpUsername         # Логин для отправки писем
[string]$smtpPassword = $config.smtpPassword         # Пароль для отпарвки писем

[string]$senderEmail = $config.senderEmail           # Адрес отправителя

[string]$receiver = $config.receiver                 # Тестовая почта

[string[]] $receiverInfo = $config.receiverInfo      # Список адресов для отправки [INFO] сообщения
[string[]] $receiverWarn = $config.receiverWarn      # Список адресов для отправки [WARN] сообщения
[string[]] $receiverError = $config.receiverError    # Список адресов для отправки [ERROR] сообщения

# Дополнительные каталоги для zip архива и лог-файлов
[string]$zipDir = Join-Path -Path $tempDir -ChildPath "Zip"    # Каталог для копирования zip архивов
[string]$logDir = Join-Path -Path $tempDir -ChildPath "Log"    # Каталог для распаковки zip архивов

<# *************************************************************
# Function section
#>


# Функция для записи лога
function Write-Log {
    param(
        [string]$message
    )
    $date = (Get-Date).ToString("dd.MM.yyyy hh:mm:ss")
    $logMessage = "$date - $message"
    Add-Content -Path $logFilePath -Value $logMessage
    Write-Host $logMessage
}


# Функция для отправки письма
function Send-Mail {
    param(
        [string]$subject,
        [string]$body,
        [string[]]$to
    )
   try {
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = $senderEmail
        $mailMessage.Subject = $subject
        $mailMessage.Body = $body
        foreach ($recipient in $to) {
            $mailMessage.To.Add($recipient)
        }
        $smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer)
        $smtpClient.Credentials = New-Object System.Net.NetworkCredential($smtpUsername, $smtpPassword)
        $smtpClient.Send($mailMessage)
    } catch {
        Write-Log ("[ERROR] Ошибка отправки письма: $($_.Exception.Message)")
    }
}

<# *************************************************************
# Code section
#>

# Логирование запуска
Write-Log "Запуск скрипта"


# Получение даты прошлого месяца
[string]$lastMonth = (Get-Date).AddMonths(-1).ToString("MM.yyyy")
#$lastMonth = "10.2024"        # Меняем дату для тестирования логики    
Write-Log "Рабочая дата: $lastMonth"


# Получение корректного месяца
[string]$month = ($lastMonth -split '\.')[0]
[string]$year = $lastMonth -split '\.' | Select-Object -Last 1
if ($month -lt 10) {
   $month = $month.TrimStart('0')
}
$archiveDir = "$archiveBaseDir\$year\$month"


# Поиск каталога с архивами
if (!(Test-Path -Path $archiveDir)) {
    Write-Log "[WARN] Каталог с архивами не найден: $archiveDir"
    Send-Mail -Subject "[WARN] Отчет о запросах" -Body "Каталог $archiveDir не найден." -To $receiverWarn
    exit
}
Write-Log "Каталог с архивами найден: $archiveDir"


# Проверка наличия архивных файлов
$archiveFiles = Get-ChildItem -Path $archiveDir -Filter "*.zip"
if ($archiveFiles.Count -eq 0) {
    Write-Log "[WARN] Архивные файлы не найдены в каталоге: $archiveDir"
    Send-Mail -Subject "[WARN] Отчет о запросах" -Body "В каталоге $archiveDir не найдены архивные файлы с расширением .zip." -To $receiverWarn
    exit
}
Write-Log "Найдено архивных файлов: $($archiveFiles.Count)"


# Создание каталогов для копирования и хранения файлов
if (!(Test-Path -Path $zipDir)) {
    New-Item -ItemType directory -Path $zipDir
}
if (!(Test-Path -Path $logDir)) {
    New-Item -ItemType directory -Path $logDir
}


# Копирование архивов в временный каталог для ZIP-файлов
foreach ($archiveFile in $archiveFiles) {
    #Write-Log "Копирование файла: $($archiveFile.FullName)"
    try {
        Copy-Item -Path $archiveFile.FullName -Destination $zipDir
    } catch {
        Write-Log ("[ERROR] Ошибка копирования файла: $($_.Exception.Message)")
        Send-Mail -Subject "[ERROR] Отчет о запросах" -Body "Ошибка копирования файла: $($_.Exception.Message)" -To $receiverError
        exit
    }
}


# Распаковка архивов
Write-Log "Распаковка архивов"
foreach ($archiveFile in Get-ChildItem -Path $zipDir -Filter "*.zip") {
    try {
        #Write-Log "Распаковка архива: $($archiveFile.FullName)"
        Expand-Archive -Path $archiveFile.FullName -DestinationPath $logDir -Force -ErrorAction Stop
    } catch {
        Write-Log "[ERROR] Ошибка распаковки архива: $($_.Exception.Message)"
        Send-Mail -Subject "[ERROR] Отчет о запросах" -Body "Ошибка распаковки архива: $($_.Exception.Message)" -To $receiverError
        exit
    }
}


# Поиск строки в файлах
Write-Log "Поиск строки '$matchString' в файлах, начинающихся с 'info.log'"

# Инициализируем переменную для подсчета количества совпадений
$count = 0

# Ищем файлы, начинающиеся с "info.log"
$logFiles = Get-ChildItem -Path $logDir -Filter "info.log*" -File

foreach ($file in $logFiles) {
    #Write-Log "Обработка файла: $($file.FullName)"
    try {
        # Подсчитываем количество совпадений в текущем файле
        $count += (Get-Content -Path $file.FullName | Where-Object { $_ -match $matchString }).Count
    } catch {
        Write-Log ("[ERROR] Ошибка обработки файла: $($_.Exception.Message)")
        Send-Mail -Subject "[ERROR] Отчет о запросах" -Body "Ошибка обработки файла: $($_.Exception.Message)" -To $receiverError
    }
}

Write-Log "Обнаружено совпадений: $count"


# Отправка письма с уведомлением
$message = "Количество запросов через XML-шлюз за $lastMonth : $count"
Write-Log "Отправка письма с уведомлением"
Send-Mail -Subject "[INFO] Отчет о запросах" -Body $message -To $receiverInfo


# Очистка содержимого каталогов Zip и Log
Remove-Item -Path $zipDir\* -Recurse -Force -Confirm:$false
Remove-Item -Path $logDir\* -Recurse -Force -Confirm:$false


# Логирование остановки
Write-Log "Остановка скрипта"