VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "IView"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' Получает объект ViewModel, связанный с представлением
Public Property Get viewModel() As Object
End Property

' Отображает представление
Public Sub Show()
End Sub

' Отображает представление в модальном режиме
Public Function ShowDialog() As Boolean
End Function

' Скрывает представление
Public Sub Hide()
End Sub
