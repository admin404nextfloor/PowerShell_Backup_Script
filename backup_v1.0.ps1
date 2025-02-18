# Указываем пути к исходной папке, временной папке и папке для архивов
$source_folder = "C:\отчеты"
$temp_folder = [System.IO.Path]::GetTempPath()
$archive_folder = "$temp_folder\archives"

# Указываем данные для подключения к FTP-серверу
$ftp_host = "ваш_ftp_сервер"
$ftp_user = "ваш_логин"
$ftp_password = "ваш_пароль"
$ftp_directory = "/путь/к/папке/на/ftp/сервере"

# Получаем текущую дату и время для создания уникального имени архива
$now = Get-Date -Format "yyyyMMdd_HHmmss"
$archive_name = "archive_$now.zip"
$archive_path = Join-Path $archive_folder $archive_name

# Создаем папку для архивов, если она не существует
if (!(Test-Path $archive_folder)) {
  New-Item -ItemType Directory -Path $archive_folder
}

# Получаем список всех файлов в исходной папке
$files = Get-ChildItem -Path $source_folder -File

# Создаем архив
try {
  # Создаем объект сжатия
  $compression = [System.IO.Compression.ZipFile]::Open($archive_path, "Create")

  # Добавляем каждый файл в архив
  foreach ($file in $files) {
    [System.IO.Compression.ZipFile]::CreateEntryFromFile($compression, $file.FullName, $file.Name)
  }

  # Закрываем объект сжатия
  $compression.Close()

  Write-Host "Архив '$archive_name' успешно создан."

} catch {
  Write-Error "Ошибка создания архива: $_"
  return
}

# Загружаем архив на FTP-сервер
try {
  # Создаем объект FTP-подключения
  $ftp = [System.Net.FtpWebRequest]::Create("ftp://$ftp_host/$ftp_directory/$archive_name")
  $ftp.Credentials = New-Object System.Net.NetworkCredential($ftp_user, $ftp_password)
  $ftp.Method = [System.Net.WebRequestMethods.Ftp]::UploadFile
  $ftp.UseBinary = $true

  # Открываем поток для чтения архива
  $fileStream = Get-ChildItem -Path $archive_path
  $requestStream = $ftp.GetRequestStream()

  # Копируем содержимое архива в поток FTP-запроса
  $buffer = New-Object byte[] 4096
  while (($bytesRead = $fileStream.Read($buffer, 0, $buffer.Length)) -ne 0) {
    $requestStream.Write($buffer, 0, $bytesRead)
  }

  # Закрываем потоки
  $requestStream.Close()
  $fileStream.Close()

  # Получаем ответ от FTP-сервера
  $response = $ftp.GetResponse()
  $response.Close()

  Write-Host "Архив '$archive_name' успешно загружен на FTP-сервер."

} catch {
  Write-Error "Ошибка загрузки на FTP-сервер: $_"
  return
}

# Удаляем архив из временной папки
Remove-Item -Path $archive_path -Force

Write-Host "Архив '$archive_name' успешно удален из временной папки."

# Проверяем, есть ли задача в планировщике
$taskName = "backup_buch"
$task = Get-ScheduledTask -TaskName $taskName

# Если задача не существует, создаем ее
if (-not $task) {
  # Получаем путь к текущему скрипту
  $scriptPath = $MyInvocation.MyCommand.Source

  # Создаем задачу
  $trigger = New-ScheduledTaskTrigger -Daily -At "21:00"
  $action = New-ScheduledTaskAction -Execute $scriptPath
  Register-ScheduledTask -TaskName $taskName -Trigger $trigger -Action $action -Description "Резервное копирование на FTP"
  Write-Host "Задача '$taskName' успешно создана в планировщике задач."
} else {
  Write-Host "Задача '$taskName' уже существует в планировщике задач."
}