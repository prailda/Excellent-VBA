VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICancellable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' Получает значение, указывающее, был ли объект отменен
Public Property Get IsCancelled() As Boolean
End Property

' Устанавливает флаг IsCancelled в True
Public Sub OnCancel()
End Sub
