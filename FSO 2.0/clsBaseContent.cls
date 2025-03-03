VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBaseContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsBaseContent - базовый класс для контента файлов и папок
Option Explicit

' Общие свойства для файлов и папок
Private pDateCreated As Date
Private pDateLastModified As Date
Private pDateLastAccessed As Date

' Общие атрибуты
Private pIsReadOnly As Boolean
Private pIsHidden As Boolean
Private pIsSystem As Boolean
Private pIsArchived As Boolean
Private pIsCompressed As Boolean
Private pIsIndexed As Boolean
Private pIsEncrypted As Boolean

' Свойства дат
Public Property Get DateCreated() As Date
    DateCreated = pDateCreated
End Property
Public Property Let DateCreated(ByVal Value As Date)
    pDateCreated = Value
End Property

Public Property Get DateLastModified() As Date
    DateLastModified = pDateLastModified
End Property
Public Property Let DateLastModified(ByVal Value As Date)
    pDateLastModified = Value
End Property

Public Property Get DateLastAccessed() As Date
    DateLastAccessed = pDateLastAccessed
End Property
Public Property Let DateLastAccessed(ByVal Value As Date)
    pDateLastAccessed = Value
End Property

' Свойства атрибутов
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = pIsReadOnly
End Property
Public Property Let IsReadOnly(ByVal Value As Boolean)
    pIsReadOnly = Value
End Property

Public Property Get IsHidden() As Boolean
    IsHidden = pIsHidden
End Property
Public Property Let IsHidden(ByVal Value As Boolean)
    pIsHidden = Value
End Property

Public Property Get IsSystem() As Boolean
    IsSystem = pIsSystem
End Property
Public Property Let IsSystem(ByVal Value As Boolean)
    pIsSystem = Value
End Property

Public Property Get IsArchived() As Boolean
    IsArchived = pIsArchived
End Property
Public Property Let IsArchived(ByVal Value As Boolean)
    pIsArchived = Value
End Property

Public Property Get IsCompressed() As Boolean
    IsCompressed = pIsCompressed
End Property
Public Property Let IsCompressed(ByVal Value As Boolean)
    pIsCompressed = Value
End Property

Public Property Get IsIndexed() As Boolean
    IsIndexed = pIsIndexed
End Property
Public Property Let IsIndexed(ByVal Value As Boolean)
    pIsIndexed = Value
End Property

Public Property Get IsEncrypted() As Boolean
    IsEncrypted = pIsEncrypted
End Property
Public Property Let IsEncrypted(ByVal Value As Boolean)
    pIsEncrypted = Value
End Property

' Метод для загрузки общих атрибутов из FSO объекта
' Метод для загрузки общих атрибутов из FSO объекта
Public Sub LoadCommonAttributes(ByVal fsoObject As Object)
    On Error GoTo ErrorHandler
    ' Даты
    pDateCreated = fsoObject.DateCreated
    pDateLastModified = fsoObject.DateLastModified
    pDateLastAccessed = fsoObject.DateLastAccessed
    
    ' Атрибуты
    With New clsFSOBase
        pIsReadOnly = .HasAttribute(fsoObject, ATTR_READONLY)
        pIsHidden = .HasAttribute(fsoObject, ATTR_HIDDEN)
        pIsSystem = .HasAttribute(fsoObject, ATTR_SYSTEM)
        pIsArchived = .HasAttribute(fsoObject, ATTR_ARCHIVE)
        pIsCompressed = .HasAttribute(fsoObject, ATTR_COMPRESSED)
        pIsIndexed = .HasAttribute(fsoObject, ATTR_INDEXED)
        pIsEncrypted = .HasAttribute(fsoObject, ATTR_ENCRYPTED)
    End With
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsBaseContent.LoadCommonAttributes", Err.Number, Err.Description, "Ошибка при загрузке атрибутов"
End Sub
