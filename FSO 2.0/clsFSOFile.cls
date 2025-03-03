VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSOFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsFSOFile - класс для работы с файлами
Option Explicit

' Наследуем базовый класс FSO
Private pFSOBase As clsFSOBase

' Специфические объекты для файла
Private pValidation As clsFileValidation
Private pContent As clsFileContent

Private Sub Class_Initialize()
    Set pFSOBase = New clsFSOBase
    Set pValidation = New clsFileValidation
    Set pContent = New clsFileContent
End Sub

' Проксирование свойств базового класса
Public Property Get fileName() As String
    fileName = pFSOBase.Name
End Property
Public Property Let fileName(ByVal Value As String)
    pFSOBase.Name = Value
End Property

Public Property Get filePath() As String
    filePath = pFSOBase.Path
End Property
Public Property Let filePath(ByVal Value As String)
    pFSOBase.Path = Value
End Property

' Доступ к вложенным объектам
Public Property Get Validation() As clsFileValidation
    Set Validation = pValidation
End Property

Public Property Get content() As clsFileContent
    Set content = pContent
End Property

Public Property Get PathHandler() As clsFSOPath
    Set PathHandler = pFSOBase.PathHandler
End Property

' Метод загрузки данных из объекта File (FSO)
Public Sub LoadFromFSO(ByVal fsoFile As Object)
    On Error GoTo ErrorHandler
    
    ' Проверка входных данных
    If fsoFile Is Nothing Then
        Err.Raise vbObjectError + 1100, "clsFSOFile.LoadFromFSO", "Объект File не может быть Nothing"
        Exit Sub
    End If
    
    ' Загрузка базовых свойств
    pFSOBase.LoadFromFSO fsoFile
    
    ' Валидация файла
    pValidation.ValidateFile fsoFile
    
    ' Загрузка содержимого файла
    pContent.LoadFromFSO fsoFile
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFile.LoadFromFSO", Err.Number, Err.Description, "Ошибка при загрузке файла: " & fsoFile.Path
End Sub

' Метод обновления информации о файле
Public Sub UpdateContent()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not pValidation.fileExists Then
        Exit Sub
    End If
    
    Dim file As Object
    Set file = fso.GetFile(Me.filePath)
    
    If Not file Is Nothing Then
        pContent.LoadFromFSO file
    End If
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFile.UpdateContent", Err.Number, Err.Description, "Ошибка при обновлении информации о файле: " & Me.filePath
End Sub

' Метод получения содержимого файла с поддержкой корректной кодировки
Public Function GetContent(Optional ByVal charset As String = "utf-8") As String
    On Error GoTo ErrorHandler
    
    If Not pValidation.fileExists Then
        GetContent = ""
        Exit Function
    End If
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Type = 2          ' Текстовый режим
        .charset = charset ' Устанавливаем нужную кодировку
        .Open
        .LoadFromFile Me.filePath
        GetContent = .ReadText(-1)  ' Читаем весь текст
        .Close
    End With
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFile.GetContent", Err.Number, Err.Description, "Ошибка чтения файла: " & Me.filePath
    GetContent = ""
End Function

' Метод записи содержимого в файл с поддержкой кодировки
' Метод записи содержимого в файл с поддержкой кодировки
Public Function WriteContent(ByVal content As String, Optional ByVal charset As String = "utf-8") As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверяем, существует ли файл
    Dim fileExists As Boolean
    fileExists = pValidation.CheckFileExists(Me.filePath)
    
    ' Если файл существует, проверяем права на запись
    If fileExists And Not pValidation.IsWritable Then
        WriteContent = False
        Exit Function
    End If
    
    ' Если файл не существует, проверяем доступность родительской папки
    If Not fileExists Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        ' Получаем путь к родительской папке
        Dim parentFolder As String
        parentFolder = PathHandler.GetParentPath()
        
        If Not fso.FolderExists(parentFolder) Then
            Dim logger As clsFSOErrorLogger
            Set logger = GetFSOErrorLoggerInstance()
            logger.LogError "clsFSOFile.WriteContent", 76, "Path not found", _
                           "Родительская папка не существует: " & parentFolder
            WriteContent = False
            Exit Function
        End If
    End If
    
    ' Записываем содержимое
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Type = 2          ' Текстовый режим
        .charset = charset ' Устанавливаем нужную кодировку
        .Open
        .WriteText content
        .SaveToFile Me.filePath, 2  ' 2 = adSaveCreateOverWrite
        .Close
    End With
    
    ' Обновляем валидацию после создания файла
    pValidation.fileExists = True
    pValidation.IsWritable = True
    pValidation.IsReadable = True
    
    ' Обновляем информацию о файле после записи
    UpdateContent
    
    WriteContent = True
    Exit Function
    
ErrorHandler:
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFile.WriteContent", Err.Number, Err.Description, "Ошибка записи в файл: " & Me.filePath
    WriteContent = False
End Function

' Метод проверки существования файла
Public Function CheckExists() As Boolean
    CheckExists = pValidation.CheckFileExists(Me.filePath)
End Function

' Метод для копирования файла в новое место
Public Function CopyTo(ByVal destinationPath As String, Optional ByVal overwrite As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    If Not pValidation.fileExists Then
        CopyTo = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.CopyFile Me.filePath, destinationPath, overwrite
    CopyTo = True
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFile.CopyTo", Err.Number, Err.Description, _
                   "Ошибка копирования файла из " & Me.filePath & " в " & destinationPath
    CopyTo = False
End Function

' Метод для удаления файла
Public Function Delete() As Boolean
    On Error GoTo ErrorHandler
    
    If Not pValidation.fileExists Then
        Delete = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.DeleteFile Me.filePath, True ' True - принудительное удаление
    Delete = True
    
    ' Обновляем состояние объекта
    pValidation.fileExists = False
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFile.Delete", Err.Number, Err.Description, "Ошибка удаления файла: " & Me.filePath
    Delete = False
End Function

' Метод для получения форматированного размера файла
Public Function GetFormattedSize() As String
    GetFormattedSize = pContent.GetFormattedSize()
End Function

