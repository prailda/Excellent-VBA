VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSOFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsFSOFolder - ����� ��� ������ � �������
Option Explicit

' ��������� ������� ����� FSO
Private pFSOBase As clsFSOBase

' ������������� ������� ��� �����
Private pValidation As clsFolderValidation
Private pContent As clsFolderContent

Private Sub Class_Initialize()
    Set pFSOBase = New clsFSOBase
    Set pValidation = New clsFolderValidation
    Set pContent = New clsFolderContent
End Sub

' ������������� ������� �������� ������
Public Property Get FolderName() As String
    FolderName = pFSOBase.Name
End Property
Public Property Let FolderName(ByVal Value As String)
    pFSOBase.Name = Value
End Property

Public Property Get FolderPath() As String
    FolderPath = pFSOBase.Path
End Property
Public Property Let FolderPath(ByVal Value As String)
    pFSOBase.Path = Value
End Property

' ������ � ��������� ��������
Public Property Get Validation() As clsFolderValidation
    Set Validation = pValidation
End Property

Public Property Get content() As clsFolderContent
    Set content = pContent
End Property

Public Property Get PathHandler() As clsFSOPath
    Set PathHandler = pFSOBase.PathHandler
End Property

' ����� �������� ������ �� ������� Folder (FSO)
Public Sub LoadFromFSO(ByVal fsoFolder As Object)
    On Error GoTo ErrorHandler
    
    ' �������� ������� ������
    If fsoFolder Is Nothing Then
        Err.Raise vbObjectError + 1200, "clsFSOFolder.LoadFromFSO", "������ Folder �� ����� ���� Nothing"
        Exit Sub
    End If
    
    ' �������� ������� �������
    pFSOBase.LoadFromFSO fsoFolder
    
    ' ��������� �����
    pValidation.ValidateFolder fsoFolder
    
    ' �������� ����������� �����
    pContent.LoadFromFSO fsoFolder
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.LoadFromFSO", Err.Number, Err.Description, "������ ��� �������� �����: " & fsoFolder.Path
End Sub

' ����� ���������� ����������� �����
Public Sub UpdateContent()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not pValidation.FolderExists Then
        Exit Sub
    End If
    
    Dim folder As Object
    Set folder = fso.GetFolder(Me.FolderPath)
    
    If Not folder Is Nothing Then
        pContent.UpdateContent folder
    End If
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.UpdateContent", Err.Number, Err.Description, _
                   "������ ��� ���������� ���������� � �����: " & Me.FolderPath
End Sub

' ����� �������� ������������� �����
Public Function CheckExists() As Boolean
    CheckExists = pValidation.CheckFolderExists(Me.FolderPath)
End Function

' ����� �������� �����, ���� ��� �� ����������
Public Function CreateIfNotExists() As Boolean
    On Error GoTo ErrorHandler
    
    If pValidation.FolderExists Then
        CreateIfNotExists = True
        Exit Function
    End If
    
    ' ���������, ���������� �� ���� � ����� ������
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.fileExists(Me.FolderPath) Then
        Dim logger As clsFSOErrorLogger
        Set logger = GetFSOErrorLoggerInstance()
        logger.LogError "clsFSOFolder.CreateIfNotExists", 58, "File already exists", _
                      "���������� ������� �����, �.�. ���������� ����: " & Me.FolderPath
        CreateIfNotExists = False
        Exit Function
    End If
    
    ' ������� ���� �����
    fso.CreateFolder Me.FolderPath
    pValidation.FolderExists = True
    CreateIfNotExists = True
    
    Exit Function
    
ErrorHandler:
    'Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.CreateIfNotExists", Err.Number, Err.Description, "������ �������� �����: " & Me.FolderPath
    CreateIfNotExists = False
End Function

' ����� ��� ����������� ����� � ����� �����
Public Function CopyTo(ByVal destinationPath As String, Optional ByVal overwrite As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    If Not pValidation.FolderExists Then
        CopyTo = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.CopyFolder Me.FolderPath, destinationPath, overwrite
    CopyTo = True
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.CopyTo", Err.Number, Err.Description, _
                   "������ ����������� ����� �� " & Me.FolderPath & " � " & destinationPath
    CopyTo = False
End Function

' ����� ��� �������� �����
Public Function Delete(Optional ByVal force As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    If Not pValidation.FolderExists Then
        Delete = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.DeleteFolder Me.FolderPath, force ' True - �������������� ��������
    Delete = True
    
    ' ��������� ��������� �������
    pValidation.FolderExists = False
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.Delete", Err.Number, Err.Description, "������ �������� �����: " & Me.FolderPath
    Delete = False
End Function

' ����� ��� ��������� ���������������� ������� �����
Public Function GetFormattedSize() As String
    GetFormattedSize = pContent.GetFormattedSize()
End Function

' ����� ��� ��������� ����� � ����� �� �����
Public Function GetFile(ByVal fileName As String) As clsFSOFile
    On Error GoTo ErrorHandler
    
    If Not pValidation.FolderExists Then
        Set GetFile = Nothing
        Exit Function
    End If
    
    ' ��������� ������������� �����
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim filePath As String
    filePath = Me.FolderPath & "\" & fileName
    
    If Not fso.fileExists(filePath) Then
        Set GetFile = Nothing
        Exit Function
    End If
    
    ' ������� ������ �����
    Dim fileObj As clsFSOFile
    Set fileObj = New clsFSOFile
    fileObj.filePath = filePath
    
    ' ��������� ������
    fileObj.LoadFromFSO fso.GetFile(filePath)
    
    Set GetFile = fileObj
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.GetFile", Err.Number, Err.Description, _
                   "������ ��������� ����� " & fileName & " � ����� " & Me.FolderPath
    Set GetFile = Nothing
End Function

' ����� ��� ��������� ������� ���� ������ � �����
Public Function GetFiles(Optional ByVal filePattern As String = "*.*") As Collection
    On Error GoTo ErrorHandler
    
    Dim result As Collection
    Set result = New Collection
    
    If Not pValidation.FolderExists Then
        Set GetFiles = result
        Exit Function
    End If
    
    ' �������� ��� �����
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(Me.FolderPath)
    
    Dim file As Object
    For Each file In folder.files
        ' ��������� ������������ �������
        If file.Name Like filePattern Then
            ' ������� ������ �����
            Dim fileObj As clsFSOFile
            Set fileObj = New clsFSOFile
            fileObj.filePath = file.Path
            fileObj.LoadFromFSO file
            
            ' ��������� � ���������
            result.Add fileObj
        End If
    Next file
    
    Set GetFiles = result
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFolder.GetFiles", Err.Number, Err.Description, _
                   "������ ��������� ������ � ����� " & Me.FolderPath
    Set GetFiles = New Collection
End Function

