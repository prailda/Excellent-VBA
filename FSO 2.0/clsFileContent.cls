VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsFileContent - класс для работы с содержимым файла
Option Explicit

' Наследуем от базового класса содержимого
Private pBaseContent As clsBaseContent
Private pFileSize As Currency

Private Sub Class_Initialize()
    Set pBaseContent = New clsBaseContent
End Sub

Public Property Get FileSize() As Currency
    FileSize = pFileSize
End Property
Public Property Let FileSize(ByVal Value As Currency)
    pFileSize = Value
End Property

' Проксирование свойств базового класса
Public Property Get DateCreated() As Date
    DateCreated = pBaseContent.DateCreated
End Property
Public Property Let DateCreated(ByVal Value As Date)
    pBaseContent.DateCreated = Value
End Property

Public Property Get DateLastModified() As Date
    DateLastModified = pBaseContent.DateLastModified
End Property
Public Property Let DateLastModified(ByVal Value As Date)
    pBaseContent.DateLastModified = Value
End Property

Public Property Get DateLastAccessed() As Date
    DateLastAccessed = pBaseContent.DateLastAccessed
End Property
Public Property Let DateLastAccessed(ByVal Value As Date)
    pBaseContent.DateLastAccessed = Value
End Property

Public Property Get IsReadOnly() As Boolean
    IsReadOnly = pBaseContent.IsReadOnly
End Property
Public Property Let IsReadOnly(ByVal Value As Boolean)
    pBaseContent.IsReadOnly = Value
End Property

Public Property Get IsHidden() As Boolean
    IsHidden = pBaseContent.IsHidden
End Property
Public Property Let IsHidden(ByVal Value As Boolean)
    pBaseContent.IsHidden = Value
End Property

Public Property Get IsSystem() As Boolean
    IsSystem = pBaseContent.IsSystem
End Property
Public Property Let IsSystem(ByVal Value As Boolean)
    pBaseContent.IsSystem = Value
End Property

Public Property Get IsArchived() As Boolean
    IsArchived = pBaseContent.IsArchived
End Property
Public Property Let IsArchived(ByVal Value As Boolean)
    pBaseContent.IsArchived = Value
End Property

Public Property Get IsCompressed() As Boolean
    IsCompressed = pBaseContent.IsCompressed
End Property
Public Property Let IsCompressed(ByVal Value As Boolean)
    pBaseContent.IsCompressed = Value
End Property

Public Property Get IsIndexed() As Boolean
    IsIndexed = pBaseContent.IsIndexed
End Property
Public Property Let IsIndexed(ByVal Value As Boolean)
    pBaseContent.IsIndexed = Value
End Property

Public Property Get IsEncrypted() As Boolean
    IsEncrypted = pBaseContent.IsEncrypted
End Property
Public Property Let IsEncrypted(ByVal Value As Boolean)
    pBaseContent.IsEncrypted = Value
End Property

' Метод для загрузки всех атрибутов из FSO объекта
Public Sub LoadFromFSO(ByVal fileObject As Object)
    On Error GoTo ErrorHandler
    If fileObject Is Nothing Then
        Err.Raise vbObjectError + 1001, "clsFileContent.LoadFromFSO", "Объект File не может быть Nothing"
        Exit Sub
    End If
    
    ' Загружаем общие атрибуты из базового класса
    pBaseContent.LoadCommonAttributes fileObject
    
    ' Специфичные для файла свойства
    pFileSize = fileObject.Size
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFileContent.LoadFromFSO", Err.Number, Err.Description, "Ошибка при загрузке содержимого файла"
End Sub

' Форматированный размер файла в человекочитаемом виде
Public Function GetFormattedSize() As String
    Dim fsoBase As New clsFSOBase
    GetFormattedSize = fsoBase.FormatSize(Me.FileSize)
End Function
