VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFolderContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsFolderContent - класс для работы с содержимым папки
Option Explicit

' Наследуем от базового класса содержимого
Private pBaseContent As clsBaseContent

' Специфичные для папки свойства
Private pFileList As String
Private pContentDictionary As Object
Private pSize As Currency
Private pFileCount As Long
Private pIsEmpty As Boolean

Private Sub Class_Initialize()
    Set pBaseContent = New clsBaseContent
    Set pContentDictionary = CreateObject("Scripting.Dictionary")
End Sub

' Специфичные для папки свойства
Public Property Get FileList() As String
    FileList = pFileList
End Property
Public Property Let FileList(ByVal Value As String)
    pFileList = Value
End Property

Public Property Get ContentDictionary() As Object
    Set ContentDictionary = pContentDictionary
End Property
Public Property Set ContentDictionary(ByVal Value As Object)
    Set pContentDictionary = Value
End Property

Public Property Get Size() As Currency
    Size = pSize
End Property
Public Property Let Size(ByVal Value As Currency)
    pSize = Value
End Property

Public Property Get FileCount() As Long
    FileCount = pFileCount
End Property
Public Property Let FileCount(ByVal Value As Long)
    pFileCount = Value
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = pIsEmpty
End Property
Public Property Let IsEmpty(ByVal Value As Boolean)
    pIsEmpty = Value
End Property

' Проксирование свойств базового класса для совместимости
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

' Метод загрузки содержимого из объекта FSO
Public Sub LoadFromFSO(ByVal folderObject As Object)
    On Error GoTo ErrorHandler
    
    If folderObject Is Nothing Then
        Err.Raise vbObjectError + 1002, "clsFolderContent.LoadFromFSO", "Объект Folder не может быть Nothing"
        Exit Sub
    End If
    
    ' Загружаем общие атрибуты из базового класса
    pBaseContent.LoadCommonAttributes folderObject
    
    ' Специфичные для папки данные
    pIsEmpty = ((folderObject.files.Count = 0) And (folderObject.SubFolders.Count = 0))
    
    ' Сбор информации о файлах
    RefreshFilesInfo folderObject
    
    ' Создание словаря содержимого
    RefreshContentDictionary folderObject
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFolderContent.LoadFromFSO", Err.Number, Err.Description, "Ошибка при загрузке содержимого папки"
End Sub

' Метод обновления информации о файлах в папке
Private Sub RefreshFilesInfo(ByVal folderObject As Object)
    On Error GoTo ErrorHandler
    
    Dim f As Object
    Dim fileNames As String: fileNames = ""
    Dim countFiles As Long: countFiles = 0
    Dim totalSize As Currency: totalSize = 0
    
    For Each f In folderObject.files
        fileNames = fileNames & f.Name & vbCrLf
        countFiles = countFiles + 1
        totalSize = totalSize + f.Size
    Next f
    
    pFileList = fileNames
    pFileCount = countFiles
    pSize = totalSize
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFolderContent.RefreshFilesInfo", Err.Number, Err.Description, "Ошибка при обновлении информации о файлах"
End Sub

' Метод обновления словаря содержимого папки
Private Sub RefreshContentDictionary(ByVal folderObject As Object)
    On Error GoTo ErrorHandler
    
    Set pContentDictionary = CreateObject("Scripting.Dictionary")
    
    ' Добавляем файлы
    Dim fileDict As Object
    Set fileDict = CreateObject("Scripting.Dictionary")
    Dim f As Object
    
    For Each f In folderObject.files
        fileDict.Add f.Name, f.Path
    Next f
    pContentDictionary.Add "Files", fileDict
    
    ' Добавляем подпапки
    Dim subFolderDict As Object
    Set subFolderDict = CreateObject("Scripting.Dictionary")
    Dim subFldr As Object
    
    For Each subFldr In folderObject.SubFolders
        subFolderDict.Add subFldr.Name, subFldr.Path
    Next subFldr
    pContentDictionary.Add "Folders", subFolderDict
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFolderContent.RefreshContentDictionary", Err.Number, Err.Description, "Ошибка при обновлении словаря содержимого"
End Sub

' Форматированный размер папки в человекочитаемом виде
Public Function GetFormattedSize() As String
    Dim fsoBase As New clsFSOBase
    GetFormattedSize = fsoBase.FormatSize(Me.Size)
End Function

' Обновление всего содержимого папки
Public Sub UpdateContent(ByVal folderObject As Object)
    On Error GoTo ErrorHandler
    
    If folderObject Is Nothing Then
        Err.Raise vbObjectError + 1003, "clsFolderContent.UpdateContent", "Объект Folder не может быть Nothing"
        Exit Sub
    End If
    
    ' Обновляем все данные
    pBaseContent.LoadCommonAttributes folderObject
    RefreshFilesInfo folderObject
    RefreshContentDictionary folderObject
    pIsEmpty = ((folderObject.files.Count = 0) And (folderObject.SubFolders.Count = 0))
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFolderContent.UpdateContent", Err.Number, Err.Description, "Ошибка при обновлении содержимого папки"
End Sub
