VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValueValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' Возвращает понятное для пользователя сообщение, описывающее правила валидации
Public Property Get Message() As String
End Property

' Возвращает True, если указанное значение действительно само по себе или в контексте источника и цели привязки
Public Function IsValid(ByVal value As Variant, ByVal Source As Object, ByVal Target As Object) As Boolean
End Function
