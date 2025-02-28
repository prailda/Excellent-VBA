VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "ICommand"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' ¬озвращает значение, указывающее, может ли команда выполнитьс€ в текущем состо€нии
Public Function CanExecute(ByVal Context As Object) As Boolean
End Function

' ¬ыполн€ет команду
Public Sub Execute(ByVal Context As Object)
End Sub

' ѕолучает описание команды
Public Property Get Description() As String
End Property
