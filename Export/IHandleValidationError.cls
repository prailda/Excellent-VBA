VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHandleValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "IHandleValidationError"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' Сбрасывает состояние IsValid для указанного свойства
Public Sub ClearValidationErrors(ByVal PropertyName As String)
End Sub

' Получает массив, содержащий все сообщения об ошибках валидации
Public Property Get ValidationErrors() As Variant
End Property

' Вызывается при возникновении ошибки валидации при обновлении привязки свойства
Public Sub OnValidationError(ByVal Source As IPropertyBinding, ByVal Message As String)
End Sub

' Возвращает True, если нет ошибок валидации
Public Property Get IsValid() As Boolean
End Property

' Возвращает первое сообщение об ошибке
Public Property Get Message() As String
End Property
