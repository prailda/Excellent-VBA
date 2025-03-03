Attribute VB_Name = "FSO_Constants"

Option Explicit

' ��������� ��������� ������/����� ��� ����� �������
Public Const ATTR_READONLY As Integer = 1
Public Const ATTR_HIDDEN As Integer = 2
Public Const ATTR_SYSTEM As Integer = 4
Public Const ATTR_DIRECTORY As Integer = 16
Public Const ATTR_ARCHIVE As Integer = 32
Public Const ATTR_NORMAL As Integer = 128
Public Const ATTR_TEMPORARY As Integer = 256
Public Const ATTR_COMPRESSED As Integer = 2048
Public Const ATTR_INDEXED As Integer = 8192
Public Const ATTR_ENCRYPTED As Integer = 16384

' ��������� ��� ������� �����������
Public Const LOG_LEVEL_INFO As String = "INFO"
Public Const LOG_LEVEL_WARNING As String = "WARNING"
Public Const LOG_LEVEL_ERROR As String = "ERROR"
Public Const LOG_LEVEL_DEBUG As String = "DEBUG"

' ��������� ��� ������� �������
Public Enum FSOAccessLevel
    ACCESS_LEVEL_FULL = 0        ' ������ ������ (������ � ������)
    ACCESS_LEVEL_READONLY = 1    ' ������ ������
    ACCESS_LEVEL_NONE = 2        ' ��� �������
End Enum

' ��������� ��� ����� ������
Public Enum FSODriveType
    DRIVE_TYPE_UNKNOWN = 0
    DRIVE_TYPE_REMOVABLE = 1
    DRIVE_TYPE_FIXED = 2
    DRIVE_TYPE_NETWORK = 3
    DRIVE_TYPE_CDROM = 4
    DRIVE_TYPE_RAMDISK = 5
End Enum


' ����� ������� ��� ��������� ������� - ������ �������� ��������
Public Function GetFSOErrorLoggerInstance() As clsFSOErrorLogger
    Static logger As clsFSOErrorLogger
    If logger Is Nothing Then
        Set logger = New clsFSOErrorLogger
    End If
    Set GetFSOErrorLoggerInstance = logger
End Function
