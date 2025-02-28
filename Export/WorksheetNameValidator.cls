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

' Создает новый экземпляр валидатора имени листа
Public Function Create() As IValueValidator
    GuardClauses.GuardNonDefaultInstance Me, WorksheetNameValidator
    Set Create = New WorksheetNameValidator
End Function

' Проверяет, является ли имя листа допустимым
Private Function IsWorksheetNameValid(ByVal name As String) As Boolean
    ' Проверка длины имени
    If Len(Trim$(name)) = 0 Then
        IsWorksheetNameValid = False
        Exit Function
    End If
    
    ' Проверка на запрещенные символы в имени листа
    Dim invalidChars As String
    invalidChars = "\/[]:*?""><|'"
    
    Dim i As Long
    For i = 1 To Len(invalidChars)
        If InStr(1, name, Mid$(invalidChars, i, 1)) > 0 Then
            IsWorksheetNameValid = False
            Exit Function
        End If
    Next i
    
    ' Длина имени листа не должна превышать 31 символ
    If Len(name) > 31 Then
        IsWorksheetNameValid = False
        Exit Function
    End If
    
    IsWorksheetNameValid = True
End Function

' Реализация интерфейса IValueValidator
Private Property Get IValueValidator_Message() As String
    IValueValidator_Message = "Имя листа не должно быть пустым, " & _
                              "содержать символы: [ ] \ / : * ? " & Chr(34) & " < > | ' " & _
                              "и должно быть не длиннее 31 символа."
End Property

Private Function IValueValidator_IsValid(ByVal value As Variant, ByVal Source As Object, ByVal Target As Object) As Boolean
    ' Проверяем передаваемое значение
    If IsNull(value) Then
        IValueValidator_IsValid = False
        Exit Function
    End If
    
    ' Если значение не строка, преобразуем в строку
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
    
    ' Проверка на соответствие правилам именования листов
    IValueValidator_IsValid = IsWorksheetNameValid(strValue)
End Function

