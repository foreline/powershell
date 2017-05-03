

ADD-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

################## Работа с почтовыми ящиками ##################

# Информация о ящиках в базе данных
Get-Mailbox -database "DB01"

# Информация о ящиках в базе данных, включая количество сообщений и занимаемый размер ящика

Get-Mailbox -database "DB01" | Get-MailboxStatistics | Sort-Object TotalitemSize -Descending | Select-Object DisplayName, ItemCount, @{name="MailboxSize";exp={$_.totalitemsize}}

Get-Mailbox -database "DB01" | Sort-Object Alias

Get-Mailbox -Arbitration -database "DB01"
Get-Mailbox -database "DB01" -Archive
Get-Mailbox -database "DB01" -PublicFolder
Get-MailboxPlan
Get-Mailbox -Arbitration -database "DB01" | New-MoveRequest -TargetDatabase "DB02"
Get-Mailbox -database "DB01" | New-MoveRequest -TargetDatabase "DB02"

Get-Mailbox -database "DB01" | FT DisplayName, TotalItemSize, IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,UseDatabaseQuotaDefaults

################## Работа с базами данных ##################

Get-MailboxDatabase

Get-MailboxDatabase DB01 -Status | FT Name,DatabaseSize,AvailableNewMailboxSpace -Auto

Get-MailboxDatabase DB01 | FL Name, *Path*

# Перемещение базы данных
Move-DatabasePath DB01 -EdbFilePath "D:\DB01\DB01.edb" -LogFolderPath "F:\DB01"

# Создание базы данных для восстановления
New-MailboxDatabase -Recovery:$true -Name "RECOVERY" -Server "Exchange" -EdbFilePath "E:\RECOVERY\RECOVERY.edb" -LogFolderPath "F:\RECOVERY"

################## DAG / Копии баз данных (Группа высокой доступности) ##################

# Информация о копиях баз данных на текущем сервере
Get-MailboxDatabaseCopyStatus

# Информация о всех копиях баз данных на указанном сервере
Get-MailboxDatabaseCopyStatus -Server "MAIL01"

# Информация по конкретной базе
Get-MailboxDatabaseCopyStatus DB01

# Перемещение активной копии базы данных на (другой) сервер
Move-ActiveMailboxDatabase -Identity "DB01" -ActivateOnServer "MAIL02"

# Перемещение всех активных копий на указанный сервер
Move-ActiveMailboxDatabase -Server "MAIL01" -ActivateOnServer "MAIL02"

# Возобновление репликации базы данных [Status -> FailedAndSuspended]
# Приостановка копии
Suspend-MailboxDatabaseCopy DB01\MAIL02
# Reseed базы данных
Update-MailboxDatabaseCopy DB01\MAIL02 -SourceServer MAIL01
# Возобновление работы копии
Resume-MailboxDatabaseCopy DB01\MAIL02

################## Index Catalog ##################

Get-MailboxDatabaseCopyStatus DB01 | FL Name,*Index*

# Вывод баз данных с поврежденным индексом [ContentIndexState -> Failed]
Get-MailboxDatabaseCopyStatus * | where {$_.ContentIndexState -eq "Failed"}
Get-MailboxDatabaseCopyStatus * | where {$_.ContentIndexState -eq "FailedAndSuspended"}
Get-MailboxDatabaseCopyStatus * | where {$_.ContentIndexState -eq "Failed"} | Update-MailboxDatabaseCopy -CatalogOnly
Get-MailboxDatabaseCopyStatus * | where {$_.ContentIndexState -eq "FailedAndSuspended"} | Update-MailboxDatabaseCopy

# Для пересоздания индекса необходимо удалить/переименовать папку "GUID.Single", которая находится в директории с базой данных

# Остановка служб
Stop-Service MSExchangeFastSearch
Stop-Service HostControllerService

# Запуск служб
Start-Service MSExchangeFastSearch
Start-Service HostControllerService

################## Запросы миграции ##################

# Проверить текущие запросы миграции
Get-MoveRequest
# Проверить текущие запросы миграции с подробной информацией
Get-MoveRequest | fl
# Проверить текущие запросы миграции с фильтрацией по статусу
Get-MoveRequest | where {$_.Status -notlike "Completed"}
# Статистика по запросу миграции
Get-MoveRequest | where {$_.Status -notlike "Completed"} | Get-MoveRequestStatistics

# Создание нового запроса миграции
#New-MoveRequest -Identity "mail01@contoso.com" -TargetDatabase "DB02" -ArchiveTargetDatabase "DB02" -BatchName "mail01@contoso.com DB01->DB02"

################## Монтирование/Размонтирование базы данных ##################

#Mount-Database DB01
#Dismount-Database DB01

################## Provisioning ##################

Get-MailboxDatabase | select Name,ServerName,IsExcludedFromProvisioning

Set-MailboxDatabase DB01 -IsExcludedFromProvisioning $true

################## Очередь отправки ##################

Get-Queue

Get-Message -Queue "Unreachable”
Get-Message -Queue “Remote Delivery Queue”

