VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValidationErrorFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "IValidationErrorFormatter"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' Получает целевой элемент управления
Public Property Get Target() As Object
End Property

' Применяет стили ошибки к целевому элементу
Public Sub Apply(ByVal PropertyName As String, ByVal Message As String)
End Sub

' Восстанавливает нормальный вид элемента
Public Sub Restore()
End Sub
