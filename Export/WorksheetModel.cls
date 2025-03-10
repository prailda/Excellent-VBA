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

' ��������� ������ ��� �������� ���������� � �����
Private Type TWorksheetInfo
    Index As Long     ' ���������� Long ������ Integer
    name As String
    BookName As String
    Worksheet As Object  ' ������ �� ������ Worksheet
End Type

' �������� ��������� ������ �� ���� �������� ����
Public Function GetWorksheets() As Collection
    Dim result As New Collection
    
    On Error Resume Next
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsInfo As TWorksheetInfo
    
    ' ���������� ��� �������� �����
    For Each wb In Workbooks
        ' ���������� ��� ����� � ������� �����
        For Each ws In wb.worksheets
            wsInfo.Index = result.Count + 1
            wsInfo.name = ws.name
            wsInfo.BookName = wb.name
            Set wsInfo.Worksheet = ws
            
            ' ��������� ���������� � ����� � ���������
            result.Add ws
        Next ws
    Next wb
    
    Set GetWorksheets = result
    
    On Error GoTo 0
End Function

' �������� ���������� � ����� �� �������
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

' �������� ���������� ������
Public Function GetWorksheetCount() As Long
    Dim worksheets As Collection
    Set worksheets = GetWorksheets()
    GetWorksheetCount = worksheets.Count
End Function

' �������� ��� ����� �� �������
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

' �������� ��� �����, ���������� ���� � ��������� ��������
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

' ��������� ����� ���� � �������� �����
Public Function AddWorksheet(ByVal name As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' ���������, ���� �� �������� �����
    If Workbooks.Count = 0 Then
        ' ������� ����� �����, ���� ��� ��������
        Workbooks.Add
    End If
    
    ' ��������� ���� � �������� �����
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.worksheets.Add(After:=ActiveWorkbook.worksheets(ActiveWorkbook.worksheets.Count))
    
    ' ������ ��� �����
    ws.name = name
    
    ' ���������� ����� ����
    ws.Activate
    
    AddWorksheet = True
    Exit Function
    
ErrorHandler:
    AddWorksheet = False
End Function

' ������� ���� �� �������
Public Function DeleteWorksheet(ByVal Index As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Object
    Set ws = GetWorksheetByIndex(Index)
    
    If Not ws Is Nothing Then
        ' ����������, ��� ����� �������� ��������� ���� �� 1 ���� � �����
        If ws.Parent.worksheets.Count > 1 Then
            Application.DisplayAlerts = False  ' ��������� ��������������
            ws.Delete
            Application.DisplayAlerts = True   ' �������� �������������� �������
            DeleteWorksheet = True
        Else
            ' ������ ������� ��������� ���� � �����
            DeleteWorksheet = False
        End If
    Else
        DeleteWorksheet = False
    End If
    
    Exit Function
    
ErrorHandler:
    Application.DisplayAlerts = True  ' �� ������ ������ �������� ��������������
    DeleteWorksheet = False
End Function

' ��������������� ���� �� �������
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
