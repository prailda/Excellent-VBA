VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSOBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsFSOBase - базовый класс для работы с объектами файловой системы
Option Explicit

' Общие свойства для файлов/папок
Private pName As String
Private pPath As String
Private pPathHandler As clsFSOPath

' Инициализация объекта
Private Sub Class_Initialize()
    Set pPathHandler = New clsFSOPath
End Sub

' Свойство для имени объекта
Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(ByVal Value As String)
    pName = Value
End Property

' Свойство для пути к объекту
Public Property Get Path() As String
    Path = pPath
End Property
Public Property Let Path(ByVal Value As String)
    pPath = Value
    pPathHandler.RawPath = Value
End Property

' Доступ к обработчику пути
Public Property Get PathHandler() As clsFSOPath
    Set PathHandler = pPathHandler
End Property

' Общий метод для проверки наличия атрибута
Public Function HasAttribute(ByVal fsoObject As Object, ByVal attributeMask As Integer) As Boolean
    On Error GoTo ErrorHandler
    HasAttribute = ((fsoObject.Attributes And attributeMask) <> 0)
    Exit Function
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOBase.HasAttribute", Err.Number, Err.Description, "Ошибка при проверке атрибута"
    HasAttribute = False
End Function

' Метод получения форматированного размера
Public Function FormatSize(ByVal sizeInBytes As Currency) As String
    Dim sizes As Variant
    sizes = Array("Б", "КБ", "МБ", "ГБ", "ТБ")
    Dim i As Integer
    
    If sizeInBytes = 0 Then
        FormatSize = "0 Б"
        Exit Function
    End If
    
    i = Int(Log(sizeInBytes) / Log(1024))
    If i > UBound(sizes) Then i = UBound(sizes)
    
    FormatSize = Format$(sizeInBytes / (1024 ^ i), "0.00") & " " & sizes(i)
End Function

' Базовый метод загрузки из FSO - должен быть переопределен в наследниках
Public Sub LoadFromFSO(ByVal fsoObject As Object)
    ' Проверка входных параметров
    If fsoObject Is Nothing Then
        Err.Raise vbObjectError + 1000, "clsFSOBase.LoadFromFSO", "Объект FSO не может быть Nothing"
        Exit Sub
    End If
    
    ' Базовые свойства, общие для всех объектов
    Me.Name = fsoObject.Name
    Me.Path = fsoObject.Path
End Sub

' Базовый метод обновления - должен быть переопределен в наследниках
Public Sub UpdateContent()
    ' Базовая реализация - должна быть переопределена
    Debug.Print "Base UpdateContent called for " & Me.Path
End Sub
