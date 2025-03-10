VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFolderValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsFolderValidation - ����� ��� ��������� �����
Option Explicit

' ��������� �� �������� ������ ���������
Private pBaseValidation As clsBaseValidation
Private pFolderExists As Boolean

Private Sub Class_Initialize()
    Set pBaseValidation = New clsBaseValidation
End Sub

' �������� �������� - ������������� �����
Public Property Get FolderExists() As Boolean
    FolderExists = pFolderExists
End Property
Public Property Let FolderExists(ByVal Value As Boolean)
    pFolderExists = Value
End Property

' ������������� ������� �������� ������
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

' ����� ��� �������� ������������� �����
Public Function CheckFolderExists(ByVal FolderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    CheckFolderExists = fso.FolderExists(FolderPath)
    pFolderExists = CheckFolderExists
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFolderValidation.CheckFolderExists", Err.Number, Err.Description, _
                   "������ ��� �������� ������������� �����: " & FolderPath
    CheckFolderExists = False
    pFolderExists = False
End Function

' ������ ��������� �����
Public Sub ValidateFolder(ByVal folderObject As Object)
    On Error GoTo ErrorHandler
    
    If folderObject Is Nothing Then
        pFolderExists = False
        pBaseValidation.IsReadable = False
        pBaseValidation.IsWritable = False
        pBaseValidation.ParentFolderExists = False
        pBaseValidation.ParentFolderName = ""
        pBaseValidation.UpdateAccessLevel
        Exit Sub
    End If
    
    ' ����� ����������
    pFolderExists = True
    
    ' �������� ������
    TestWritePermission folderObject
    
    ' �������� ������
    TestReadPermission folderObject
    
    ' ���������� ������ �������
    pBaseValidation.UpdateAccessLevel
    
    ' �������� ���������� � ������������ �����
    pBaseValidation.LoadParentFolderInfo folderObject
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFolderValidation.ValidateFolder", Err.Number, Err.Description, _
                   "������ ��� ��������� �����"
End Sub

' �������� ����������� ������ � �����
Private Sub TestWritePermission(ByVal folderObject As Object)
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ������� ��������� ���� ��� �������� ���� ������
    Dim tmpFileName As String
    tmpFileName = folderObject.Path & "\" & "temp_access_" & Format(Now, "yyyymmdd_hhnnss") & ".tmp"
    
    Dim tmpFile As Object
    Set tmpFile = fso.CreateTextFile(tmpFileName, True)
    
    If Err.Number = 0 Then
        tmpFile.Close
        fso.DeleteFile tmpFileName, True
        pBaseValidation.IsWritable = True
    Else
        pBaseValidation.IsWritable = False
    End If
    
    On Error GoTo 0
End Sub

' �������� ����������� ������ �����
Private Sub TestReadPermission(ByVal folderObject As Object)
    On Error Resume Next
    
    ' �������� �������� ���������� ������ � �����
    Dim testCount As Long
    testCount = folderObject.files.Count
    
    If Err.Number = 0 Then
        pBaseValidation.IsReadable = True
    Else
        pBaseValidation.IsReadable = False
    End If
    
    On Error GoTo 0
End Sub

