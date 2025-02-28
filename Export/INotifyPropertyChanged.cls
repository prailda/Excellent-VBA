VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INotifyPropertyChanged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "INotifyPropertyChanged"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' Регистрирует обработчик для события изменения свойств
Public Sub RegisterHandler(ByVal Handler As IHandlePropertyChanged)
End Sub

' Уведомляет всех зарегистрированных обработчиков об изменении свойства
Public Sub OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
End Sub
