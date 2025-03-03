VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsFileValidation - класс для валидации файлов
Option Explicit

' Наследуем от базового класса валидации
Private pBaseValidation As clsBaseValidation
Private pFileExists As Boolean

Private Sub Class_Initialize()
    Set pBaseValidation = New clsBaseValidation
End Sub

' Основное свойство - существование файла
Public Property Get fileExists() As Boolean
    fileExists = pFileExists
End Property
Public Property Let fileExists(ByVal Value As Boolean)
    pFileExists = Value
End Property

' Проксирование свойств базового класса
Public Property Get IsWritable() As Boolean
    IsWritable = pBaseValidation.IsWritable
End Property
Public Property Let IsWritable(ByVal Value As Boolean)
    pBaseValidation.IsWritable = Value
    pBaseValidation.UpdateAccessLevel
End Property

Public Property Get IsReadable() As Boolean
    IsReadable = pBaseValidation.IsReadable
End Property
Public Property Let IsReadable(ByVal Value As Boolean)
    pBaseValidation.IsReadable = Value
    pBaseValidation.UpdateAccessLevel
End Property

Public Property Get AccessLevel() As Integer
    AccessLevel = pBaseValidation.AccessLevel
End Property
Public Property Let AccessLevel(ByVal Value As Integer)
    pBaseValidation.AccessLevel = Value
End Property

Public Property Get ParentFolderExists() As Boolean
    ParentFolderExists = pBaseValidation.ParentFolderExists
End Property
Public Property Let ParentFolderExists(ByVal Value As Boolean)
    pBaseValidation.ParentFolderExists = Value
End Property

Public Property Get ParentFolderName() As String
    ParentFolderName = pBaseValidation.ParentFolderName
End Property
Public Property Let ParentFolderName(ByVal Value As String)
    pBaseValidation.ParentFolderName = Value
End Property

' Метод для проверки существования файла
Public Function CheckFileExists(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    CheckFileExists = fso.fileExists(filePath)
    pFileExists = CheckFileExists
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFileValidation.CheckFileExists", Err.Number, Err.Description, _
                   "Ошибка при проверке существования файла: " & filePath
    CheckFileExists = False
    pFileExists = False
End Function

' Полная валидация файла
Public Sub ValidateFile(ByVal fileObject As Object)
    On Error GoTo ErrorHandler
    
    If fileObject Is Nothing Then
        pFileExists = False
        pBaseValidation.IsReadable = False
        pBaseValidation.IsWritable = False
        pBaseValidation.ParentFolderExists = False
        pBaseValidation.ParentFolderName = ""
        pBaseValidation.UpdateAccessLevel
        Exit Sub
    End If
    
    ' Файл существует
    pFileExists = True
    
    ' Проверка записи
    TestWritePermission fileObject
    
    ' Проверка чтения
    TestReadPermission fileObject
    
    ' Обновление уровня доступа
    pBaseValidation.UpdateAccessLevel
    
    ' Загрузка информации о родительской папке
    pBaseValidation.LoadParentFolderInfo fileObject
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFileValidation.ValidateFile", Err.Number, Err.Description, _
                   "Ошибка при валидации файла"
End Sub

' Проверка возможности записи в файл
Private Sub TestWritePermission(ByVal fileObject As Object)
    On Error Resume Next
    
    Dim ts As Object
    Set ts = fileObject.OpenAsTextStream(8, 0)  ' 8 = ForAppending
    
    If Err.Number = 0 Then
        ts.Close
        pBaseValidation.IsWritable = True
    Else
        pBaseValidation.IsWritable = False
    End If
    
    On Error GoTo 0
End Sub

' Проверка возможности чтения файла
Private Sub TestReadPermission(ByVal fileObject As Object)
    On Error Resume Next
    
    Dim ts As Object
    Set ts = fileObject.OpenAsTextStream(1, 0)  ' 1 = ForReading
    
    If Err.Number = 0 Then
        ts.Close
        pBaseValidation.IsReadable = True
    Else
        pBaseValidation.IsReadable = False
    End If
    
    On Error GoTo 0
End Sub

