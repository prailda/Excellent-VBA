Attribute VB_Name = "FSO_Demo_Improved"

Option Explicit

' ������������ ������ � ������� � ����� �����������
Sub TestImprovedFSOFile()
    Dim fso As Object
    Dim fileObj As Object
    Dim myFile As clsFSOFile
    Dim logger As clsFSOErrorLogger
    
    ' ������� ������ FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' �������� ��������� �������
    Set logger = GetFSOErrorLoggerInstance()
    
    On Error GoTo ErrorHandler
    
    ' �������� ���� �� ������������ ���� �� ����� ����������
    Set fileObj = fso.GetFile("C:\Users\dalis\����������\11. ������������\����� �����\�3.txt")
    
    ' ������� ��������� ����� ���������� ������ � ��������� ������ �� FSO
    Set myFile = New clsFSOFile
    myFile.LoadFromFSO fileObj
    
    logger.LogInfo "TestImprovedFSOFile", "���� �������� � ������� LoadFromFSO."
    
    ' ����� ������� ������
    Debug.Print "=== ���������� � ����� ==="
    Debug.Print "��� �����: " & myFile.fileName
    Debug.Print "���� �����: " & myFile.filePath
    
    ' ����� ������ ���������
    Debug.Print vbCrLf & "=== ��������� ����� ==="
    Debug.Print "���� ����������? " & myFile.Validation.fileExists
    Debug.Print "�������� ��� ������? " & myFile.Validation.IsWritable
    Debug.Print "�������� ��� ������? " & myFile.Validation.IsReadable
    Debug.Print "������� �������: " & myFile.Validation.AccessLevel
    Debug.Print "������������ ����� ����������? " & myFile.Validation.ParentFolderExists
    Debug.Print "��� ������������ �����: " & myFile.Validation.ParentFolderName
    
    ' ����� ������ �����������
    Debug.Print vbCrLf & "=== ���������� ����� ==="
    Debug.Print "������ ����� (����): " & myFile.content.FileSize
    Debug.Print "������ � �������� �������: " & myFile.GetFormattedSize()
    Debug.Print "���� ��������: " & myFile.content.DateCreated
    Debug.Print "���� ���������� ���������: " & myFile.content.DateLastModified
    Debug.Print "���� ���������� �������: " & myFile.content.DateLastAccessed
    
    ' ����� ���������
    Debug.Print vbCrLf & "=== �������� ����� ==="
    Debug.Print "ReadOnly: " & myFile.content.IsReadOnly
    Debug.Print "Hidden: " & myFile.content.IsHidden
    Debug.Print "System: " & myFile.content.IsSystem
    Debug.Print "Archive: " & myFile.content.IsArchived
    Debug.Print "Compressed: " & myFile.content.IsCompressed
    Debug.Print "Indexed: " & myFile.content.IsIndexed
    Debug.Print "Encrypted: " & myFile.content.IsEncrypted
    
    ' ������ ����������� �����
    If myFile.Validation.IsReadable Then
        Debug.Print vbCrLf & "=== ������ 200 �������� ����������� ==="
        Dim content As String
        content = myFile.GetContent()
        Debug.Print Left$(content, 200)
    End If
    
    ' ���������� ���������� � �����
    myFile.UpdateContent
    logger.LogInfo "TestImprovedFSOFile", "���������� � ����� ���������."
    
    Debug.Print vbCrLf & "=== ����� ���������� ==="
    Debug.Print "������ ����� (����): " & myFile.content.FileSize
    
    Exit Sub
    
ErrorHandler:
    logger.LogError "TestImprovedFSOFile", Err.Number, Err.Description, "������ ��� ������������ �����."
    MsgBox "������: " & Err.Description, vbCritical, "������ �����"
End Sub

' ������������ ������ � ������� � ����� �����������
Sub TestImprovedFSOFolder()
    Dim fso As Object
    Dim folderObj As Object
    Dim myFolder As clsFSOFolder
    Dim logger As clsFSOErrorLogger
    
    ' ������� ������ FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' �������� ��������� �������
    Set logger = GetFSOErrorLoggerInstance()
    
    On Error GoTo ErrorHandler
    
    ' �������� ���� �� ������������ ����� �� ����� ����������
    Set folderObj = fso.GetFolder("C:\Users\dalis\����������\11. ������������\����� �����")
    
    ' ������� ��������� ���������� ������ � ��������� ������ �� FSO
    Set myFolder = New clsFSOFolder
    myFolder.LoadFromFSO folderObj
    
    logger.LogInfo "TestImprovedFSOFolder", "����� ��������� � ������� LoadFromFSO."
    
    ' ����� ������� ������
    Debug.Print "=== ���������� � ����� ==="
    Debug.Print "��� �����: " & myFolder.FolderName
    Debug.Print "���� �����: " & myFolder.FolderPath
    
    ' ����� ������ ���������
    Debug.Print vbCrLf & "=== ��������� ����� ==="
    Debug.Print "����� ����������? " & myFolder.Validation.FolderExists
    Debug.Print "�������� ��� ������? " & myFolder.Validation.IsWritable
    Debug.Print "�������� ��� ������? " & myFolder.Validation.IsReadable
    Debug.Print "������� �������: " & myFolder.Validation.AccessLevel
    Debug.Print "������������ ����� ����������? " & myFolder.Validation.ParentFolderExists
    Debug.Print "��� ������������ �����: " & myFolder.Validation.ParentFolderName
    
    ' ����� ������ �����������
    Debug.Print vbCrLf & "=== ���������� ����� ==="
    Debug.Print "���������� ������: " & myFolder.content.FileCount
    Debug.Print "������ ����� (����): " & myFolder.content.Size
    Debug.Print "������ � �������� �������: " & myFolder.GetFormattedSize()
    Debug.Print "���� ��������: " & myFolder.content.DateCreated
    Debug.Print "���� ���������� ���������: " & myFolder.content.DateLastModified
    Debug.Print "���� ���������� �������: " & myFolder.content.DateLastAccessed
    
    ' ����� ���������
    Debug.Print vbCrLf & "=== �������� ����� ==="
    Debug.Print "������? " & myFolder.content.IsEmpty
    Debug.Print "ReadOnly: " & myFolder.content.IsReadOnly
    Debug.Print "Hidden: " & myFolder.content.IsHidden
    Debug.Print "System: " & myFolder.content.IsSystem
    Debug.Print "Archive: " & myFolder.content.IsArchived
    Debug.Print "Compressed: " & myFolder.content.IsCompressed
    Debug.Print "Indexed: " & myFolder.content.IsIndexed
    Debug.Print "Encrypted: " & myFolder.content.IsEncrypted
    
    ' ����� ������ ������ ����� ������� ��������
    Dim dict As Object, fileDict As Object, key As Variant
    Set dict = myFolder.content.ContentDictionary
    
    If dict.Exists("Files") Then
        Set fileDict = dict("Files")
        Debug.Print vbCrLf & "=== ������ ������ � ����� ==="
        For Each key In fileDict.Keys
            Debug.Print " " & key & " => " & fileDict(key)
        Next key
    End If
    
    ' ��������� ���� ������ � ����� � �������������� ������ ������
    Debug.Print vbCrLf & "=== ��������� ������ ����� GetFiles ==="
    Dim files As Collection
    Set files = myFolder.GetFiles()
    
    Dim fileObj As clsFSOFile
    For Each fileObj In files
        Debug.Print fileObj.fileName & " (" & fileObj.GetFormattedSize() & ")"
    Next fileObj
    
    ' ���������� ����������� �����
    myFolder.UpdateContent
    logger.LogInfo "TestImprovedFSOFolder", "���������� ����� ���������."
    
    Debug.Print vbCrLf & "=== ����� ���������� ==="
    Debug.Print "���������� ������: " & myFolder.content.FileCount
    Debug.Print "������ ����� (����): " & myFolder.content.Size
    
    Exit Sub
    
ErrorHandler:
    logger.LogError "TestImprovedFSOFolder", Err.Number, Err.Description, "������ ��� ������������ �����."
    MsgBox "������: " & Err.Description, vbCritical, "������ �����"
End Sub

' ������������ ������ � ����������� � ������
Sub TestImprovedFSOGeneral()
    Dim gen As clsFSOGeneral
    Dim logger As clsFSOErrorLogger
    
    Set gen = New clsFSOGeneral
    Set logger = GetFSOErrorLoggerInstance()
    
    ' ������������� ���� ��� ��������
    gen.DriveLetter = "C:\"
    
    Debug.Print "=== ���������� � ����� " & gen.DriveLetter & " ==="
    
    ' �������� ���������� �����
    If Not gen.IsDriveReady() Then
        Debug.Print "���� �� �����!"
        Exit Sub
    End If
    
    ' ��������� ���������� � �����
    Dim avail As Currency, total As Currency, used As Currency
    avail = gen.GetAvailableSpace()
    total = gen.GetTotalSize()
    used = gen.GetUsedSpace()
    
    Debug.Print "��� �����: " & gen.GetDriveTypeText()
    Debug.Print "�������� �����: " & gen.GetDriveSerialNumber()
    Debug.Print "��������: " & gen.FormatBytes(avail) & " (" & Format(gen.GetFreeSpacePercent(), "0.00") & "%)"
    Debug.Print "����� ������: " & gen.FormatBytes(total)
    Debug.Print "������: " & gen.FormatBytes(used)
    
    ' �������� ������� ���������� �����
    If gen.IsEnoughFreeSpace(104857600) Then ' 100 ��
        Debug.Print "���������� ���������� ����� ��� 100 ��."
    Else
        Debug.Print "������������ ���������� ����� ��� 100 ��."
    End If
    
    ' ��������� ������ ���� ��������� ������
    Debug.Print vbCrLf & "=== ������ ��������� ������ ==="
    Dim drives As Collection
    Set drives = gen.GetAllDrives()
    
    Dim drive As Variant
    For Each drive In drives
        gen.DriveLetter = drive & ":\"
        Debug.Print drive & ":\ - " & gen.GetDriveTypeText()
    Next drive
End Sub

' ������������ �������� � ����������� � ������� � �������
Sub TestFSOFileOperations()
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    
    On Error GoTo ErrorHandler
    
    ' �������� �������� �����
' �������� �������� �����
Dim testFolder As clsFSOFolder
Set testFolder = New clsFSOFolder
testFolder.FolderPath = "C:\Users\dalis\����������\11. ������������\FSO_Test_Folder"
    
    Debug.Print "=== �������� � ������� ==="
    
    ' �������� � �������� �����
    If Not testFolder.CheckExists() Then
        Debug.Print "����� �� ����������, �������..."
        If testFolder.CreateIfNotExists() Then
            Debug.Print "����� ������� �������."
        Else
            Debug.Print "������ ��� �������� �����!"
            Exit Sub
        End If
    Else
        Debug.Print "����� ��� ����������."
    End If
    
    ' �������� ��������� �����
    Dim testFile As clsFSOFile
    Set testFile = New clsFSOFile
    testFile.filePath = testFolder.FolderPath & "\test_file.txt"
    
    Debug.Print vbCrLf & "=== �������� � ������� ==="
    
    ' ������ ������ � ����
    Debug.Print "������� �������� ����..."
    If testFile.WriteContent("��� �������� ����, ��������� � ������� clsFSOFile." & vbCrLf & _
                           "������� ����: " & Now) Then
        Debug.Print "���� ������� ������ � �������."
    Else
        Debug.Print "������ ��� �������� �����!"
        Exit Sub
    End If
    
    ' ������ ������ �� �����
    Debug.Print vbCrLf & "���������� �����:"
    Debug.Print testFile.GetContent()
    
    ' ����������� �����
    Debug.Print vbCrLf & "����������� �����..."
    If testFile.CopyTo(testFolder.FolderPath & "\test_file_copy.txt", True) Then
        Debug.Print "���� ������� ����������."
    Else
        Debug.Print "������ ��� ����������� �����!"
    End If
    
    ' ��������� ����� �� �����
    Debug.Print vbCrLf & "��������� ����� �� �����..."
    Dim copiedFile As clsFSOFile
    Set copiedFile = testFolder.GetFile("test_file_copy.txt")
    
    If Not copiedFile Is Nothing Then
        Debug.Print "������ ����: " & copiedFile.fileName
        Debug.Print "������: " & copiedFile.GetFormattedSize()
    Else
        Debug.Print "���� �� ������!"
    End If
    
    ' ��������� ���� ������ � �����
    Debug.Print vbCrLf & "��� ����� � �����:"
    Dim files As Collection
    Set files = testFolder.GetFiles()
    
    Dim file As clsFSOFile
    For Each file In files
        Debug.Print "- " & file.fileName & " (" & file.GetFormattedSize() & ")"
    Next file
    
    ' �������� ������ � �����
    If MsgBox("������� ��������� �������� ����� � �����?", vbYesNo + vbQuestion, "�������������") = vbYes Then
        Debug.Print vbCrLf & "�������� �������� ������ � �����..."
        
        ' �������� ���� ������ � �����
        For Each file In files
            If file.Delete() Then
                Debug.Print "���� " & file.fileName & " ������."
            Else
                Debug.Print "������ ��� �������� ����� " & file.fileName
            End If
        Next file
        
        ' �������� �����
        If testFolder.Delete(True) Then
            Debug.Print "����� " & testFolder.FolderPath & " �������."
        Else
            Debug.Print "������ ��� �������� ����� " & testFolder.FolderPath
        End If
    End If
    
    Debug.Print vbCrLf & "������������ ���������."
    
    Exit Sub
    
ErrorHandler:
    logger.LogError "TestFSOFileOperations", Err.Number, Err.Description, "������ ��� ���������� �������� � �������."
    MsgBox "������: " & Err.Description, vbCritical, "������ �����"
End Sub

