Attribute VB_Name = "FSO_Demo_Improved"

Option Explicit

' Демонстрация работы с файлами в новой архитектуре
Sub TestImprovedFSOFile()
    Dim fso As Object
    Dim fileObj As Object
    Dim myFile As clsFSOFile
    Dim logger As clsFSOErrorLogger
    
    ' Создаем объект FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Получаем экземпляр логгера
    Set logger = GetFSOErrorLoggerInstance()
    
    On Error GoTo ErrorHandler
    
    ' Замените путь на существующий файл на вашем компьютере
    Set fileObj = fso.GetFile("C:\Users\dalis\Библиотека\11. Тестирование\Новая папка\№3.txt")
    
    ' Создаем экземпляр нашей улучшенной модели и загружаем данные из FSO
    Set myFile = New clsFSOFile
    myFile.LoadFromFSO fileObj
    
    logger.LogInfo "TestImprovedFSOFile", "Файл загружен с помощью LoadFromFSO."
    
    ' Вывод базовых данных
    Debug.Print "=== ИНФОРМАЦИЯ О ФАЙЛЕ ==="
    Debug.Print "Имя файла: " & myFile.fileName
    Debug.Print "Путь файла: " & myFile.filePath
    
    ' Вывод данных валидации
    Debug.Print vbCrLf & "=== ВАЛИДАЦИЯ ФАЙЛА ==="
    Debug.Print "Файл существует? " & myFile.Validation.fileExists
    Debug.Print "Доступен для записи? " & myFile.Validation.IsWritable
    Debug.Print "Доступен для чтения? " & myFile.Validation.IsReadable
    Debug.Print "Уровень доступа: " & myFile.Validation.AccessLevel
    Debug.Print "Родительская папка существует? " & myFile.Validation.ParentFolderExists
    Debug.Print "Имя родительской папки: " & myFile.Validation.ParentFolderName
    
    ' Вывод данных содержимого
    Debug.Print vbCrLf & "=== СОДЕРЖИМОЕ ФАЙЛА ==="
    Debug.Print "Размер файла (байт): " & myFile.content.FileSize
    Debug.Print "Размер в читаемом формате: " & myFile.GetFormattedSize()
    Debug.Print "Дата создания: " & myFile.content.DateCreated
    Debug.Print "Дата последнего изменения: " & myFile.content.DateLastModified
    Debug.Print "Дата последнего доступа: " & myFile.content.DateLastAccessed
    
    ' Вывод атрибутов
    Debug.Print vbCrLf & "=== АТРИБУТЫ ФАЙЛА ==="
    Debug.Print "ReadOnly: " & myFile.content.IsReadOnly
    Debug.Print "Hidden: " & myFile.content.IsHidden
    Debug.Print "System: " & myFile.content.IsSystem
    Debug.Print "Archive: " & myFile.content.IsArchived
    Debug.Print "Compressed: " & myFile.content.IsCompressed
    Debug.Print "Indexed: " & myFile.content.IsIndexed
    Debug.Print "Encrypted: " & myFile.content.IsEncrypted
    
    ' Чтение содержимого файла
    If myFile.Validation.IsReadable Then
        Debug.Print vbCrLf & "=== ПЕРВЫЕ 200 СИМВОЛОВ СОДЕРЖИМОГО ==="
        Dim content As String
        content = myFile.GetContent()
        Debug.Print Left$(content, 200)
    End If
    
    ' Обновление информации о файле
    myFile.UpdateContent
    logger.LogInfo "TestImprovedFSOFile", "Информация о файле обновлена."
    
    Debug.Print vbCrLf & "=== ПОСЛЕ ОБНОВЛЕНИЯ ==="
    Debug.Print "Размер файла (байт): " & myFile.content.FileSize
    
    Exit Sub
    
ErrorHandler:
    logger.LogError "TestImprovedFSOFile", Err.Number, Err.Description, "Ошибка при тестировании файла."
    MsgBox "Ошибка: " & Err.Description, vbCritical, "Ошибка теста"
End Sub

' Демонстрация работы с папками в новой архитектуре
Sub TestImprovedFSOFolder()
    Dim fso As Object
    Dim folderObj As Object
    Dim myFolder As clsFSOFolder
    Dim logger As clsFSOErrorLogger
    
    ' Создаем объект FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Получаем экземпляр логгера
    Set logger = GetFSOErrorLoggerInstance()
    
    On Error GoTo ErrorHandler
    
    ' Замените путь на существующую папку на вашем компьютере
    Set folderObj = fso.GetFolder("C:\Users\dalis\Библиотека\11. Тестирование\Новая папка")
    
    ' Создаем экземпляр улучшенной модели и загружаем данные из FSO
    Set myFolder = New clsFSOFolder
    myFolder.LoadFromFSO folderObj
    
    logger.LogInfo "TestImprovedFSOFolder", "Папка загружена с помощью LoadFromFSO."
    
    ' Вывод базовых данных
    Debug.Print "=== ИНФОРМАЦИЯ О ПАПКЕ ==="
    Debug.Print "Имя папки: " & myFolder.FolderName
    Debug.Print "Путь папки: " & myFolder.FolderPath
    
    ' Вывод данных валидации
    Debug.Print vbCrLf & "=== ВАЛИДАЦИЯ ПАПКИ ==="
    Debug.Print "Папка существует? " & myFolder.Validation.FolderExists
    Debug.Print "Доступна для записи? " & myFolder.Validation.IsWritable
    Debug.Print "Доступна для чтения? " & myFolder.Validation.IsReadable
    Debug.Print "Уровень доступа: " & myFolder.Validation.AccessLevel
    Debug.Print "Родительская папка существует? " & myFolder.Validation.ParentFolderExists
    Debug.Print "Имя родительской папки: " & myFolder.Validation.ParentFolderName
    
    ' Вывод данных содержимого
    Debug.Print vbCrLf & "=== СОДЕРЖИМОЕ ПАПКИ ==="
    Debug.Print "Количество файлов: " & myFolder.content.FileCount
    Debug.Print "Размер папки (байт): " & myFolder.content.Size
    Debug.Print "Размер в читаемом формате: " & myFolder.GetFormattedSize()
    Debug.Print "Дата создания: " & myFolder.content.DateCreated
    Debug.Print "Дата последнего изменения: " & myFolder.content.DateLastModified
    Debug.Print "Дата последнего доступа: " & myFolder.content.DateLastAccessed
    
    ' Вывод атрибутов
    Debug.Print vbCrLf & "=== АТРИБУТЫ ПАПКИ ==="
    Debug.Print "Пустая? " & myFolder.content.IsEmpty
    Debug.Print "ReadOnly: " & myFolder.content.IsReadOnly
    Debug.Print "Hidden: " & myFolder.content.IsHidden
    Debug.Print "System: " & myFolder.content.IsSystem
    Debug.Print "Archive: " & myFolder.content.IsArchived
    Debug.Print "Compressed: " & myFolder.content.IsCompressed
    Debug.Print "Indexed: " & myFolder.content.IsIndexed
    Debug.Print "Encrypted: " & myFolder.content.IsEncrypted
    
    ' Вывод списка файлов через словарь контента
    Dim dict As Object, fileDict As Object, key As Variant
    Set dict = myFolder.content.ContentDictionary
    
    If dict.Exists("Files") Then
        Set fileDict = dict("Files")
        Debug.Print vbCrLf & "=== СПИСОК ФАЙЛОВ В ПАПКЕ ==="
        For Each key In fileDict.Keys
            Debug.Print " " & key & " => " & fileDict(key)
        Next key
    End If
    
    ' Получение всех файлов в папке с использованием нового метода
    Debug.Print vbCrLf & "=== ПОЛУЧЕНИЕ ФАЙЛОВ ЧЕРЕЗ GetFiles ==="
    Dim files As Collection
    Set files = myFolder.GetFiles()
    
    Dim fileObj As clsFSOFile
    For Each fileObj In files
        Debug.Print fileObj.fileName & " (" & fileObj.GetFormattedSize() & ")"
    Next fileObj
    
    ' Обновление содержимого папки
    myFolder.UpdateContent
    logger.LogInfo "TestImprovedFSOFolder", "Содержимое папки обновлено."
    
    Debug.Print vbCrLf & "=== ПОСЛЕ ОБНОВЛЕНИЯ ==="
    Debug.Print "Количество файлов: " & myFolder.content.FileCount
    Debug.Print "Размер папки (байт): " & myFolder.content.Size
    
    Exit Sub
    
ErrorHandler:
    logger.LogError "TestImprovedFSOFolder", Err.Number, Err.Description, "Ошибка при тестировании папки."
    MsgBox "Ошибка: " & Err.Description, vbCritical, "Ошибка теста"
End Sub

' Демонстрация работы с информацией о дисках
Sub TestImprovedFSOGeneral()
    Dim gen As clsFSOGeneral
    Dim logger As clsFSOErrorLogger
    
    Set gen = New clsFSOGeneral
    Set logger = GetFSOErrorLoggerInstance()
    
    ' Устанавливаем диск для проверки
    gen.DriveLetter = "C:\"
    
    Debug.Print "=== ИНФОРМАЦИЯ О ДИСКЕ " & gen.DriveLetter & " ==="
    
    ' Проверка готовности диска
    If Not gen.IsDriveReady() Then
        Debug.Print "Диск не готов!"
        Exit Sub
    End If
    
    ' Получение информации о диске
    Dim avail As Currency, total As Currency, used As Currency
    avail = gen.GetAvailableSpace()
    total = gen.GetTotalSize()
    used = gen.GetUsedSpace()
    
    Debug.Print "Тип диска: " & gen.GetDriveTypeText()
    Debug.Print "Серийный номер: " & gen.GetDriveSerialNumber()
    Debug.Print "Свободно: " & gen.FormatBytes(avail) & " (" & Format(gen.GetFreeSpacePercent(), "0.00") & "%)"
    Debug.Print "Общий размер: " & gen.FormatBytes(total)
    Debug.Print "Занято: " & gen.FormatBytes(used)
    
    ' Проверка наличия свободного места
    If gen.IsEnoughFreeSpace(104857600) Then ' 100 МБ
        Debug.Print "Достаточно свободного места для 100 МБ."
    Else
        Debug.Print "Недостаточно свободного места для 100 МБ."
    End If
    
    ' Получение списка всех доступных дисков
    Debug.Print vbCrLf & "=== СПИСОК ДОСТУПНЫХ ДИСКОВ ==="
    Dim drives As Collection
    Set drives = gen.GetAllDrives()
    
    Dim drive As Variant
    For Each drive In drives
        gen.DriveLetter = drive & ":\"
        Debug.Print drive & ":\ - " & gen.GetDriveTypeText()
    Next drive
End Sub

' Демонстрация создания и манипуляции с файлами и папками
Sub TestFSOFileOperations()
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    
    On Error GoTo ErrorHandler
    
    ' Создание тестовой папки
' Создание тестовой папки
Dim testFolder As clsFSOFolder
Set testFolder = New clsFSOFolder
testFolder.FolderPath = "C:\Users\dalis\Библиотека\11. Тестирование\FSO_Test_Folder"
    
    Debug.Print "=== ОПЕРАЦИИ С ПАПКАМИ ==="
    
    ' Проверка и создание папки
    If Not testFolder.CheckExists() Then
        Debug.Print "Папка не существует, создаем..."
        If testFolder.CreateIfNotExists() Then
            Debug.Print "Папка успешно создана."
        Else
            Debug.Print "Ошибка при создании папки!"
            Exit Sub
        End If
    Else
        Debug.Print "Папка уже существует."
    End If
    
    ' Создание тестового файла
    Dim testFile As clsFSOFile
    Set testFile = New clsFSOFile
    testFile.filePath = testFolder.FolderPath & "\test_file.txt"
    
    Debug.Print vbCrLf & "=== ОПЕРАЦИИ С ФАЙЛАМИ ==="
    
    ' Запись данных в файл
    Debug.Print "Создаем тестовый файл..."
    If testFile.WriteContent("Это тестовый файл, созданный с помощью clsFSOFile." & vbCrLf & _
                           "Текущая дата: " & Now) Then
        Debug.Print "Файл успешно создан и записан."
    Else
        Debug.Print "Ошибка при создании файла!"
        Exit Sub
    End If
    
    ' Чтение данных из файла
    Debug.Print vbCrLf & "Содержимое файла:"
    Debug.Print testFile.GetContent()
    
    ' Копирование файла
    Debug.Print vbCrLf & "Копирование файла..."
    If testFile.CopyTo(testFolder.FolderPath & "\test_file_copy.txt", True) Then
        Debug.Print "Файл успешно скопирован."
    Else
        Debug.Print "Ошибка при копировании файла!"
    End If
    
    ' Получение файла из папки
    Debug.Print vbCrLf & "Получение файла из папки..."
    Dim copiedFile As clsFSOFile
    Set copiedFile = testFolder.GetFile("test_file_copy.txt")
    
    If Not copiedFile Is Nothing Then
        Debug.Print "Найден файл: " & copiedFile.fileName
        Debug.Print "Размер: " & copiedFile.GetFormattedSize()
    Else
        Debug.Print "Файл не найден!"
    End If
    
    ' Получение всех файлов в папке
    Debug.Print vbCrLf & "Все файлы в папке:"
    Dim files As Collection
    Set files = testFolder.GetFiles()
    
    Dim file As clsFSOFile
    For Each file In files
        Debug.Print "- " & file.fileName & " (" & file.GetFormattedSize() & ")"
    Next file
    
    ' Удаление файлов и папок
    If MsgBox("Удалить созданные тестовые файлы и папку?", vbYesNo + vbQuestion, "Подтверждение") = vbYes Then
        Debug.Print vbCrLf & "Удаление тестовых файлов и папки..."
        
        ' Удаление всех файлов в папке
        For Each file In files
            If file.Delete() Then
                Debug.Print "Файл " & file.fileName & " удален."
            Else
                Debug.Print "Ошибка при удалении файла " & file.fileName
            End If
        Next file
        
        ' Удаление папки
        If testFolder.Delete(True) Then
            Debug.Print "Папка " & testFolder.FolderPath & " удалена."
        Else
            Debug.Print "Ошибка при удалении папки " & testFolder.FolderPath
        End If
    End If
    
    Debug.Print vbCrLf & "Тестирование завершено."
    
    Exit Sub
    
ErrorHandler:
    logger.LogError "TestFSOFileOperations", Err.Number, Err.Description, "Ошибка при выполнении операций с файлами."
    MsgBox "Ошибка: " & Err.Description, vbCritical, "Ошибка теста"
End Sub

