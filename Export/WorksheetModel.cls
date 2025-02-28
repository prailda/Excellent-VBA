VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "WorksheetModel"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'@Folder "MVVM.Model"
Option Explicit

' Структура данных для хранения информации о листе
Private Type TWorksheetInfo
    Index As Long     ' Используем Long вместо Integer
    name As String
    BookName As String
    Worksheet As Object  ' Ссылка на объект Worksheet
End Type

' Получает коллекцию листов из всех открытых книг
Public Function GetWorksheets() As Collection
    Dim result As New Collection
    
    On Error Resume Next
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsInfo As TWorksheetInfo
    
    ' Перебираем все открытые книги
    For Each wb In Workbooks
        ' Перебираем все листы в текущей книге
        For Each ws In wb.worksheets
            wsInfo.Index = result.Count + 1
            wsInfo.name = ws.name
            wsInfo.BookName = wb.name
            Set wsInfo.Worksheet = ws
            
            ' Добавляем информацию о листе в результат
            result.Add ws
        Next ws
    Next wb
    
    Set GetWorksheets = result
    
    On Error GoTo 0
End Function

' Получает информацию о листе по индексу
Public Function GetWorksheetByIndex(ByVal Index As Long) As Object
    Dim worksheets As Collection
    
    On Error Resume Next
    
    Set worksheets = GetWorksheets()
    
    If Index > 0 And Index <= worksheets.Count Then
        Set GetWorksheetByIndex = worksheets(Index)
    Else
        Set GetWorksheetByIndex = Nothing
    End If
    
    On Error GoTo 0
End Function

' Получает количество листов
Public Function GetWorksheetCount() As Long
    Dim worksheets As Collection
    Set worksheets = GetWorksheets()
    GetWorksheetCount = worksheets.Count
End Function

' Получает имя листа по индексу
Public Function GetWorksheetName(ByVal Index As Long) As String
    On Error Resume Next
    
    Dim ws As Object
    Set ws = GetWorksheetByIndex(Index)
    
    If Not ws Is Nothing Then
        GetWorksheetName = ws.name
    Else
        GetWorksheetName = ""
    End If
    
    On Error GoTo 0
End Function

' Получает имя книги, содержащей лист с указанным индексом
Public Function GetWorksheetBookName(ByVal Index As Long) As String
    On Error Resume Next
    
    Dim ws As Object
    Set ws = GetWorksheetByIndex(Index)
    
    If Not ws Is Nothing Then
        GetWorksheetBookName = ws.Parent.name
    Else
        GetWorksheetBookName = ""
    End If
    
    On Error GoTo 0
End Function

' Добавляет новый лист в активную книгу
Public Function AddWorksheet(ByVal name As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверяем, есть ли активная книга
    If Workbooks.Count = 0 Then
        ' Создаем новую книгу, если нет открытых
        Workbooks.Add
    End If
    
    ' Добавляем лист в активную книгу
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.worksheets.Add(After:=ActiveWorkbook.worksheets(ActiveWorkbook.worksheets.Count))
    
    ' Задаем имя листу
    ws.name = name
    
    ' Активируем новый лист
    ws.Activate
    
    AddWorksheet = True
    Exit Function
    
ErrorHandler:
    AddWorksheet = False
End Function

' Удаляет лист по индексу
Public Function DeleteWorksheet(ByVal Index As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Object
    Set ws = GetWorksheetByIndex(Index)
    
    If Not ws Is Nothing Then
        ' Убеждаемся, что после удаления останется хотя бы 1 лист в книге
        If ws.Parent.worksheets.Count > 1 Then
            Application.DisplayAlerts = False  ' Отключаем предупреждение
            ws.Delete
            Application.DisplayAlerts = True   ' Включаем предупреждения обратно
            DeleteWorksheet = True
        Else
            ' Нельзя удалить последний лист в книге
            DeleteWorksheet = False
        End If
    Else
        DeleteWorksheet = False
    End If
    
    Exit Function
    
ErrorHandler:
    Application.DisplayAlerts = True  ' На всякий случай включаем предупреждения
    DeleteWorksheet = False
End Function

' Переименовывает лист по индексу
Public Function RenameWorksheet(ByVal Index As Long, ByVal newName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Object
    Set ws = GetWorksheetByIndex(Index)
    
    If Not ws Is Nothing Then
        ws.name = newName
        RenameWorksheet = True
    Else
        RenameWorksheet = False
    End If
    
    Exit Function
    
ErrorHandler:
    RenameWorksheet = False
End Function
