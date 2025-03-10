VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetNameValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "WorksheetNameValidator"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
'@Folder "MVVM.Validation"
'@PredeclaredId
Option Explicit
Implements IValueValidator

' ������� ����� ��������� ���������� ����� �����
Public Function Create() As IValueValidator
    GuardClauses.GuardNonDefaultInstance Me, WorksheetNameValidator
    Set Create = New WorksheetNameValidator
End Function

' ���������, �������� �� ��� ����� ����������
Private Function IsWorksheetNameValid(ByVal name As String) As Boolean
    ' �������� ����� �����
    If Len(Trim$(name)) = 0 Then
        IsWorksheetNameValid = False
        Exit Function
    End If
    
    ' �������� �� ����������� ������� � ����� �����
    Dim invalidChars As String
    invalidChars = "\/[]:*?""><|'"
    
    Dim i As Long
    For i = 1 To Len(invalidChars)
        If InStr(1, name, Mid$(invalidChars, i, 1)) > 0 Then
            IsWorksheetNameValid = False
            Exit Function
        End If
    Next i
    
    ' ����� ����� ����� �� ������ ��������� 31 ������
    If Len(name) > 31 Then
        IsWorksheetNameValid = False
        Exit Function
    End If
    
    IsWorksheetNameValid = True
End Function

' ���������� ���������� IValueValidator
Private Property Get IValueValidator_Message() As String
    IValueValidator_Message = "��� ����� �� ������ ���� ������, " & _
                              "��������� �������: [ ] \ / : * ? " & Chr(34) & " < > | ' " & _
                              "� ������ ���� �� ������� 31 �������."
End Property

Private Function IValueValidator_IsValid(ByVal value As Variant, ByVal Source As Object, ByVal Target As Object) As Boolean
    ' ��������� ������������ ��������
    If IsNull(value) Then
        IValueValidator_IsValid = False
        Exit Function
    End If
    
    ' ���� �������� �� ������, ����������� � ������
    Dim strValue As String
    If VarType(value) <> vbString Then
        On Error Resume Next
        strValue = CStr(value)
        If Err.Number <> 0 Then
            IValueValidator_IsValid = False
            Exit Function
        End If
        On Error GoTo 0
    Else
        strValue = value
    End If
    
    ' �������� �� ������������ �������� ���������� ������
    IValueValidator_IsValid = IsWorksheetNameValid(strValue)
End Function

