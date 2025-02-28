VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' Перечисление для режимов привязки
Public Enum BindingMode
    TwoWayBinding         ' Двусторонняя привязка (изменения в обоих направлениях)
    OneWayBinding         ' Односторонняя привязка (от источника к цели)
    OneWayToSource        ' Односторонняя привязка (от цели к источнику)
    OneTimeBinding        ' Одноразовая привязка (только при инициализации)
End Enum

' Перечисление для триггеров обновления источника
Public Enum BindingUpdateSourceTrigger
    OnPropertyChanged     ' Обновлять при изменении свойства
    OnKeyPress            ' Обновлять при нажатии клавиши
    OnExit                ' Обновлять при выходе из элемента управления
End Enum

' Получает значение, указывающее режим/направление привязки
Public Property Get Mode() As BindingMode
End Property

' Получает значение, указывающее триггер обновления привязки
Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
End Property

' Получает объект ViewModel, который является источником привязки
Public Property Get Source() As Object
End Property

' Получает путь к процедуре Public Property Get, разрешаемой из источника привязки
Public Property Get SourcePropertyPath() As String
End Property

' Получает элемент пользовательского интерфейса, который является целью привязки
Public Property Get Target() As Object
End Property

' Получает имя привязанного свойства на целевом объекте
Public Property Get TargetProperty() As String
End Property

' Получает конвертер для использования (если есть) при применении привязки
Public Property Get Converter() As IValueConverter
End Property

' Применяет привязку
Public Sub Apply()
End Sub
