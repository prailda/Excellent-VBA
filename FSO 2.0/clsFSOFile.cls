VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSOFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsFSOFile - ����� ��� ������ � �������
Option Explicit

' ��������� ������� ����� FSO
Private pFSOBase As clsFSOBase

' ������������� ������� ��� �����
Private pValidation As clsFileValidation
Private pContent As clsFileContent

Private Sub Class_Initialize()
    Set pFSOBase = New clsFSOBase
    Set pValidation = New clsFileValidation
    Set pContent = New clsFileContent
End Sub

' ������������� ������� �������� ������
Public Property Get fileName() As String
    fileName = pFSOBase.Name
End Property
Public Property Let fileName(ByVal Value As String)
    pFSOBase.Name = Value
End Property

Public Property Get filePath() As String
    filePath = pFSOBase.Path
End Property
Public Property Let filePath(ByVal Value As String)
    pFSOBase.Path = Value
End Property

' ������ � ��������� ��������
Public Property Get Validation() As clsFileValidation
    Set Validation = pValidation
End Property

Public Property Get content() As clsFileContent
    Set content = pContent
End Property

Public Property Get PathHandler() As clsFSOPath
    Set PathHandler = pFSOBase.PathHandler
End Property

' ����� �������� ������ �� ������� File (FSO)
Public Sub LoadFromFSO(ByVal fsoFile As Object)
    On Error GoTo ErrorHandler
    
    ' �������� ������� ������
    If fsoFile Is Nothing Then
        Err.Raise vbObjectError + 1100, "clsFSOFile.LoadFromFSO", "������ File �� ����� ���� Nothing"
        Exit Sub
    End If
    
    ' �������� ������� �������
    pFSOBase.LoadFromFSO fsoFile
    
    ' ��������� �����
    pValidation.ValidateFile fsoFile
    
    ' �������� ����������� �����
    pContent.LoadFromFSO fsoFile
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFile.LoadFromFSO", Err.Number, Err.Description, "������ ��� �������� �����: " & fsoFile.Path
End Sub

' ����� ���������� ���������� � �����
Public Sub UpdateContent()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not pValidation.fileExists Then
        Exit Sub
    End If
    
    Dim file As Object
    Set file = fso.GetFile(Me.filePath)
    
    If Not file Is Nothing Then
        pContent.LoadFromFSO file
    End If
    
    Exit Sub
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFile.UpdateContent", Err.Number, Err.Description, "������ ��� ���������� ���������� � �����: " & Me.filePath
End Sub

' ����� ��������� ����������� ����� � ���������� ���������� ���������
Public Function GetContent(Optional ByVal charset As String = "utf-8") As String
    On Error GoTo ErrorHandler
    
    If Not pValidation.fileExists Then
        GetContent = ""
        Exit Function
    End If
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Type = 2          ' ��������� �����
        .charset = charset ' ������������� ������ ���������
        .Open
        .LoadFromFile Me.filePath
        GetContent = .ReadText(-1)  ' ������ ���� �����
        .Close
    End With
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFile.GetContent", Err.Number, Err.Description, "������ ������ �����: " & Me.filePath
    GetContent = ""
End Function

' ����� ������ ����������� � ���� � ���������� ���������
' ����� ������ ����������� � ���� � ���������� ���������
Public Function WriteContent(ByVal content As String, Optional ByVal charset As String = "utf-8") As Boolean
    On Error GoTo ErrorHandler
    
    ' ���������, ���������� �� ����
    Dim fileExists As Boolean
    fileExists = pValidation.CheckFileExists(Me.filePath)
    
    ' ���� ���� ����������, ��������� ����� �� ������
    If fileExists And Not pValidation.IsWritable Then
        WriteContent = False
        Exit Function
    End If
    
    ' ���� ���� �� ����������, ��������� ����������� ������������ �����
    If Not fileExists Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        ' �������� ���� � ������������ �����
        Dim parentFolder As String
        parentFolder = PathHandler.GetParentPath()
        
        If Not fso.FolderExists(parentFolder) Then
            Dim logger As clsFSOErrorLogger
            Set logger = GetFSOErrorLoggerInstance()
            logger.LogError "clsFSOFile.WriteContent", 76, "Path not found", _
                           "������������ ����� �� ����������: " & parentFolder
            WriteContent = False
            Exit Function
        End If
    End If
    
    ' ���������� ����������
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Type = 2          ' ��������� �����
        .charset = charset ' ������������� ������ ���������
        .Open
        .WriteText content
        .SaveToFile Me.filePath, 2  ' 2 = adSaveCreateOverWrite
        .Close
    End With
    
    ' ��������� ��������� ����� �������� �����
    pValidation.fileExists = True
    pValidation.IsWritable = True
    pValidation.IsReadable = True
    
    ' ��������� ���������� � ����� ����� ������
    UpdateContent
    
    WriteContent = True
    Exit Function
    
ErrorHandler:
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFile.WriteContent", Err.Number, Err.Description, "������ ������ � ����: " & Me.filePath
    WriteContent = False
End Function

' ����� �������� ������������� �����
Public Function CheckExists() As Boolean
    CheckExists = pValidation.CheckFileExists(Me.filePath)
End Function

' ����� ��� ����������� ����� � ����� �����
Public Function CopyTo(ByVal destinationPath As String, Optional ByVal overwrite As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    If Not pValidation.fileExists Then
        CopyTo = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.CopyFile Me.filePath, destinationPath, overwrite
    CopyTo = True
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFile.CopyTo", Err.Number, Err.Description, _
                   "������ ����������� ����� �� " & Me.filePath & " � " & destinationPath
    CopyTo = False
End Function

' ����� ��� �������� �����
Public Function Delete() As Boolean
    On Error GoTo ErrorHandler
    
    If Not pValidation.fileExists Then
        Delete = False
        Exit Function
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.DeleteFile Me.filePath, True ' True - �������������� ��������
    Delete = True
    
    ' ��������� ��������� �������
    pValidation.fileExists = False
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOFile.Delete", Err.Number, Err.Description, "������ �������� �����: " & Me.filePath
    Delete = False
End Function

' ����� ��� ��������� ���������������� ������� �����
Public Function GetFormattedSize() As String
    GetFormattedSize = pContent.GetFormattedSize()
End Function

