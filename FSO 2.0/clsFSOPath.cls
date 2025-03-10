VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSOPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsFSOPath - ����� ��� ������ � ������ � �������� �������
Option Explicit

' ��������� ��� ��������� �����
Private Const MAX_PATH_LENGTH As Integer = 260
Private Const PATH_DELIMITER As String = "\"

Private pRawPath As String  ' �������� ����
Private pCleanPath As String ' ��������� ���� ����� ���������

' �������� ��������� ����
Public Property Get RawPath() As String
    RawPath = pRawPath
End Property

' ��� ��������� RawPath ������������� ������������ ����
Public Property Let RawPath(ByVal Value As String)
    pRawPath = Value
    pCleanPath = ProcessPath(Value)
End Property

' ��������, ������������ ��������� ����
Public Property Get CleanPath() As String
    CleanPath = pCleanPath
End Property

' ������� ��������� ����: ��������� ������� �������� ������.
Public Function ProcessPath(ByVal pathStr As String) As String
    On Error GoTo ErrorHandler
    
    ' ������� �������
    Dim result As String
    result = Trim(pathStr)
    result = Replace(result, "/", PATH_DELIMITER)
    result = Replace(result, """", "")
    
    ' ������� ����������� ����� � �������
    If Right(result, 1) = ";" Then
        result = Left(result, Len(result) - 1)
    End If
    
    ' ���������� �������� ����������: ������� ��������� (��������, "C:")
    If InStr(1, result, ":", vbTextCompare) = 0 Then
        result = ""
    End If
    
    ProcessPath = result
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOPath.ProcessPath", Err.Number, Err.Description, "������ ��� ��������� ����: " & pathStr
    ProcessPath = ""
End Function

' �������� ���� �� ����������
Public Function SplitPath() As Variant
    If pCleanPath = "" Then
        SplitPath = Array()
    Else
        SplitPath = Split(pCleanPath, PATH_DELIMITER)
    End If
End Function

' ��������� ����� ������� (��������� ��������� ����)
Public Function GetObjectName() As String
    Dim parts As Variant
    parts = SplitPath()
    If UBound(parts) >= 0 Then
        GetObjectName = parts(UBound(parts))
    Else
        GetObjectName = ""
    End If
End Function

' ��������� ������������� ���� (��� ����������, ����� ����������)
Public Function GetParentPath() As String
    Dim parts As Variant
    parts = SplitPath()
    If UBound(parts) < 1 Then
        GetParentPath = ""
    Else
        Dim result As String: result = ""
        Dim i As Long
        For i = LBound(parts) To UBound(parts) - 1
            result = result & parts(i) & PATH_DELIMITER
        Next i
        If Right(result, 1) = PATH_DELIMITER Then result = Left(result, Len(result) - 1)
        GetParentPath = result
    End If
End Function

' �������� ���� �� ������� �����������
Public Function BuildPath(ByVal parts As Variant) As String
    Dim result As String
    Dim i As Long
    result = ""
    For i = LBound(parts) To UBound(parts)
        If result = "" Then
            result = parts(i)
        Else
            If Right(result, 1) <> PATH_DELIMITER Then result = result & PATH_DELIMITER
            result = result & parts(i)
        End If
    Next i
    BuildPath = result
End Function

' �������� ����� ���� (True ���� ����� <= 260 ��������)
Public Function IsPathLengthValid() As Boolean
    IsPathLengthValid = (Len(pCleanPath) <= MAX_PATH_LENGTH)
End Function

' ���������� ��������� ����� ���� (����������� �����)
Public Function BuildPathTree() As String
    Dim fso As Object, folder As Object
    Dim result As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set folder = fso.GetFolder(pCleanPath)
    On Error GoTo 0
    If folder Is Nothing Then
        BuildPathTree = "������������ ����: " & pCleanPath
        Exit Function
    End If
    result = BuildTreeForFolder(folder, "")
    BuildPathTree = result
End Function

Private Function BuildTreeForFolder(ByVal folder As Object, ByVal indent As String) As String
    Dim result As String
    result = indent & folder.Name & vbCrLf
    Dim f As Object, subFldr As Object
    For Each f In folder.files
        result = result & indent & "  " & f.Name & vbCrLf
    Next f
    For Each subFldr In folder.SubFolders
        result = result & BuildTreeForFolder(subFldr, indent & "  ")
    Next subFldr
    BuildTreeForFolder = result
End Function
