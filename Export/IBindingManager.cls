VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBindingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "IBindingManager"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' Привязывает путь к свойству на объекте-источнике к свойству элемента управления
Public Sub BindPropertyPath(ByVal Source As Object, ByVal PropertyPath As String, ByVal Target As Object, _
Optional ByVal TargetProperty As String = "", _
Optional ByVal Mode As BindingMode = TwoWayBinding, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal ErrorFormat As IValidationErrorFormatter)
End Sub

' Привязывает команду к элементу управления
Public Sub BindCommand(ByVal Source As Object, ByVal Target As Object, ByVal Command As ICommand)
End Sub

' Применяет все привязки и оценивает состояние команд
Public Sub ApplyBindings(ByVal Source As Object)
End Sub

' Оценивает состояние всех привязанных команд
Public Sub OnEvaluateCanExecute(ByVal Source As Object)
End Sub

