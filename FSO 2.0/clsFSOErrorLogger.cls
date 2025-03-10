VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSOErrorLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsFSOErrorLogger - ����� ��� ����������� ������ � ���������
Option Explicit

' ��������� �������� ����� ����������� ���������� ������ ������� GetInstance
'Public Function GetInstance() As clsFSOErrorLogger
'    Static logger As clsFSOErrorLogger
'    If logger Is Nothing Then
'        Set logger = New clsFSOErrorLogger
'    End If
'    Set GetInstance = logger
'End Function

' ����������� ��������������� ���������: ������:
' [���� �����] [INFO] [������������] [������] ���������
Public Sub LogInfo(ByVal objName As String, ByVal message As String)
    LogMessage LOG_LEVEL_INFO, objName, message
End Sub

' ����������� ��������������
Public Sub LogWarning(ByVal objName As String, ByVal message As String)
    LogMessage LOG_LEVEL_WARNING, objName, message
End Sub

' ����������� ��������� �� ������: ������:
' [���� �����] [ERROR] [������������] [��������� �����] Error Code: X, Description: � [, Info: �]
Public Sub LogError(ByVal methodName As String, ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal additionalInfo As String = "")
    Dim logStr As String
    logStr = FormatLogHeader(LOG_LEVEL_ERROR, methodName) & _
             "Error Code: " & errNumber & ", Description: " & errDescription
    If additionalInfo <> "" Then logStr = logStr & ", Info: " & additionalInfo
    Debug.Print logStr
End Sub

' ������������� ����� ����������� � ��������� ������
Public Sub LogMessage(ByVal level As String, ByVal objName As String, ByVal message As String)
    Dim logStr As String
    logStr = FormatLogHeader(level, objName) & message
    Debug.Print logStr
End Sub

' ����������� ����������� ������ � ��������� ���������� (��������� "����� �������")
Public Sub LogDetailedError(ByVal methodName As String, ByVal errNumber As Long, ByVal errDescription As String, ByVal variableState As String)
    Dim logStr As String
    logStr = FormatLogHeader(LOG_LEVEL_ERROR, methodName) & _
             "Error Code: " & errNumber & ", Description: " & errDescription & ", Variable State: " & variableState
    Debug.Print logStr
End Sub

' ����������� ��������� ��������� ����
Private Function FormatLogHeader(ByVal level As String, ByVal objName As String) As String
    FormatLogHeader = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [" & level & "] [" & Environ("USERNAME") & "] [" & objName & "] "
End Function

' ���������� ������� ��� ��������� ���������� ������� - �������� GetInstance ��� �������� �������������
Public Function GetFSOErrorLoggerInstance() As clsFSOErrorLogger
    Set GetFSOErrorLoggerInstance = GetInstance()
End Function

Public Static Function GetInstance() As clsFSOErrorLogger
    Static logger As clsFSOErrorLogger
    If logger Is Nothing Then
        Set logger = New clsFSOErrorLogger
    End If
    Set GetInstance = logger
End Function
