VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBaseValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsBaseValidation - базовый класс для валидации файловых объектов
Option Explicit

' Общие свойства для валидации
Private pIsWritable As Boolean
Private pIsReadable As Boolean
Private pAccessLevel As FSOAccessLevel
Private pParentFolderExists As Boolean
Private pParentFolderName As String

' Свойства доступа
Public Property Get IsWritable() As Boolean
    IsWritable = pIsWritable
End Property
Public Property Let IsWritable(ByVal Value As Boolean)
    pIsWritable = Value
End Property

Public Property Get IsReadable() As Boolean
    IsReadable = pIsReadable
End Property
Public Property Let IsReadable(ByVal Value As Boolean)
    pIsReadable = Value
End Property

Public Property Get AccessLevel() As FSOAccessLevel
    AccessLevel = pAccessLevel
End Property
Public Property Let AccessLevel(ByVal Value As FSOAccessLevel)
    pAccessLevel = Value
End Property

' Свойства родительской папки
Public Property Get ParentFolderExists() As Boolean
    ParentFolderExists = pParentFolderExists
End Property
Public Property Let ParentFolderExists(ByVal Value As Boolean)
    pParentFolderExists = Value
End Property

Public Property Get ParentFolderName() As String
    ParentFolderName = pParentFolderName
End Property
Public Property Let ParentFolderName(ByVal Value As String)
    pParentFolderName = Value
End Property

' Метод загрузки родительской папки из FSO объекта
Public Sub LoadParentFolderInfo(ByVal fsoObject As Object)
    On Error GoTo ErrorHandler
    
    If fsoObject Is Nothing Then
        pParentFolderExists = False
        pParentFolderName = ""
        Exit Sub
    End If
    
    If Not fsoObject.parentFolder Is Nothing Then
        pParentFolderExists = True
        pParentFolderName = fsoObject.parentFolder.Name
    Else
        pParentFolderExists = False
        pParentFolderName = ""
    End If
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsBaseValidation.LoadParentFolderInfo", Err.Number, Err.Description, _
                   "Ошибка при загрузке информации о родительской папке"
    pParentFolderExists = False
    pParentFolderName = ""
End Sub

' Обновление уровня доступа на основе прав чтения и записи
Public Sub UpdateAccessLevel()
    If pIsWritable Then
        pAccessLevel = ACCESS_LEVEL_FULL
    ElseIf pIsReadable Then
        pAccessLevel = ACCESS_LEVEL_READONLY
    Else
        pAccessLevel = ACCESS_LEVEL_NONE
    End If
End Sub

