VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSOFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsFSOFolder - класс для работы с папками
Option Explicit

' Наследуем базовый класс FSO
Private pFSOBase As clsFSOBase

' Специфические объекты для папки
Private pValidation As clsFolderValidation
Private pContent As clsFolderContent

Private Sub Class_Initialize()
    Set pFSOBase = New clsFSOBase
    Set pValidation = New clsFolderValidation
    Set pContent = New clsFolderContent
End Sub

' Проксирование свойств базового класса
Public Property Get FolderName() As String
    FolderName = pFSOBase.Name
End Property
Public Property Let FolderName(ByVal Value As String)
    pFSOBase.Name = Value
End Property

Public Property Get FolderPath() As String
    FolderPath = pFSOBase.Path
End Property
Public Property Let FolderPath(ByVal Value As String)
    pFSOBase.Path = Value
End Property

' Доступ к вложенным объектам
Public Property Get Validation() As clsFolderValidation
    Set Validation = pValidation
End Property

Public Property Get content() As clsFolderContent
    Set content = pContent
End Property

Public Property Get PathHandler() As clsFSOPath
    Set PathHandler = pFSOBase.PathHandler
End Property

' Метод загрузки данных из объекта Folder (FSO)
Public Sub LoadFromFSO(ByVal fsoFolder As Object)
    On Error GoTo ErrorHandler
    
    ' Проверка входных данных
    If fsoFolder Is Nothing Then
        Err.Raise vbObjectError + 1200, "clsFSOFolder.LoadFromFSO", "Объект Folder не может быть Nothing"
        Exit Sub
    End If
    
    ' Загрузка базовых свойств
    pFSOBase.LoadFromFSO fsoFolder
    
    ' Валидация папки
    pValidation.ValidateFolder fsoFolder
    
    ' Загрузка содержимого папки
    pContent.LoadFromFSO fsoFolder
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.LoadFromFSO", Err.Number, Err.Description, "Ошибка при загрузке папки: " & fsoFolder.Path
End Sub

' Метод обновления содержимого папки
Public Sub UpdateContent()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not pValidation.FolderExists Then
        Exit Sub
    End If
    
    Dim folder As Object
    Set folder = fso.GetFolder(Me.FolderPath)
    
    If Not folder Is Nothing Then
        pContent.UpdateContent folder
    End If
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.UpdateContent", Err.Number, Err.Description, _
                   "Ошибка при обновлении информации о папке: " & Me.FolderPath
End Sub

' Метод проверки существования папки
Public Function CheckExists() As Boolean
    CheckExists = pValidation.CheckFolderExists(Me.FolderPath)
End Function

' Метод создания папки, если она не существует
Public Function CreateIfNotExists() As Boolean
    On Error GoTo ErrorHandler
    
    If pValidation.FolderExists Then
        CreateIfNotExists = True
        Exit Function
    End If
    
    ' Проверяем, существует ли файл с таким именем
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.fileExists(Me.FolderPath) Then
        Dim logger As clsFSOErrorLogger
        Set logger = GetFSOErrorLoggerInstance()
        logger.LogError "clsFSOFolder.CreateIfNotExists", 58, "File already exists", _
                      "Невозможно создать папку, т.к. существует файл: " & Me.FolderPath
        CreateIfNotExists = False
        Exit Function
    End If
    
    ' Создаем саму папку
    fso.CreateFolder Me.FolderPath
    pValidation.FolderExists = True
    CreateIfNotExists = True
    
    Exit Function
    
ErrorHandler:
    'Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.CreateIfNotExists", Err.Number, Err.Description, "Ошибка создания папки: " & Me.FolderPath
    CreateIfNotExists = False
End Function

' Метод для копирования папки в новое место
Public Function CopyTo(ByVal destinationPath As String, Optional ByVal overwrite As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    If Not pValidation.FolderExists Then
        CopyTo = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.CopyFolder Me.FolderPath, destinationPath, overwrite
    CopyTo = True
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.CopyTo", Err.Number, Err.Description, _
                   "Ошибка копирования папки из " & Me.FolderPath & " в " & destinationPath
    CopyTo = False
End Function

' Метод для удаления папки
Public Function Delete(Optional ByVal force As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    If Not pValidation.FolderExists Then
        Delete = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.DeleteFolder Me.FolderPath, force ' True - принудительное удаление
    Delete = True
    
    ' Обновляем состояние объекта
    pValidation.FolderExists = False
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.Delete", Err.Number, Err.Description, "Ошибка удаления папки: " & Me.FolderPath
    Delete = False
End Function

' Метод для получения форматированного размера папки
Public Function GetFormattedSize() As String
    GetFormattedSize = pContent.GetFormattedSize()
End Function

' Метод для получения файла в папке по имени
Public Function GetFile(ByVal fileName As String) As clsFSOFile
    On Error GoTo ErrorHandler
    
    If Not pValidation.FolderExists Then
        Set GetFile = Nothing
        Exit Function
    End If
    
    ' Проверяем существование файла
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim filePath As String
    filePath = Me.FolderPath & "\" & fileName
    
    If Not fso.fileExists(filePath) Then
        Set GetFile = Nothing
        Exit Function
    End If
    
    ' Создаем объект файла
    Dim fileObj As clsFSOFile
    Set fileObj = New clsFSOFile
    fileObj.filePath = filePath
    
    ' Загружаем данные
    fileObj.LoadFromFSO fso.GetFile(filePath)
    
    Set GetFile = fileObj
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.GetFile", Err.Number, Err.Description, _
                   "Ошибка получения файла " & fileName & " в папке " & Me.FolderPath
    Set GetFile = Nothing
End Function

' Метод для получения массива всех файлов в папке
Public Function GetFiles(Optional ByVal filePattern As String = "*.*") As Collection
    On Error GoTo ErrorHandler
    
    Dim result As Collection
    Set result = New Collection
    
    If Not pValidation.FolderExists Then
        Set GetFiles = result
        Exit Function
    End If
    
    ' Получаем все файлы
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(Me.FolderPath)
    
    Dim file As Object
    For Each file In folder.files
        ' Проверяем соответствие шаблону
        If file.Name Like filePattern Then
            ' Создаем объект файла
            Dim fileObj As clsFSOFile
            Set fileObj = New clsFSOFile
            fileObj.filePath = file.Path
            fileObj.LoadFromFSO file
            
            ' Добавляем в коллекцию
            result.Add fileObj
        End If
    Next file
    
    Set GetFiles = result
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.GetFiles", Err.Number, Err.Description, _
                   "Ошибка получения файлов в папке " & Me.FolderPath
    Set GetFiles = New Collection
End Function

