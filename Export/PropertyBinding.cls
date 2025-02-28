VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "PropertyBinding"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
'@Folder "MVVM.Binding"
'@PredeclaredId
Option Explicit
Implements IPropertyBinding
Implements IHandlePropertyChanged
Implements INotifyPropertyChanged

Private Type TState
    Handlers As Collection
    Manager As IBindingManager
    Mode As BindingMode
    UpdateSourceTrigger As BindingUpdateSourceTrigger
    Source As Object
    SourcePropertyPath As String
    Target As Object
    TargetProperty As String
    Converter As IValueConverter
    Validator As IValueValidator
    ValidationErrorHandler As IHandleValidationError
    ValidationErrorFormatter As IValidationErrorFormatter
    Applied As Boolean
    PropertyChangedHandler As IHandlePropertyChanged
End Type

Private WithEvents TextBoxEventSource As MSForms.TextBox
Attribute TextBoxEventSource.VB_VarHelpID = -1
Private this As TState

' Создает экземпляр привязки свойства для TextBox
Public Function ForTextBox(ByVal Manager As IBindingManager, ByVal Target As MSForms.TextBox, ByVal Source As Object, ByVal sourceProperty As String, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal UpdateSource As BindingUpdateSourceTrigger = OnPropertyChanged, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal ErrorFormat As IValidationErrorFormatter) As IPropertyBinding
    Set ForTextBox = Create(Manager, Target, "Text", Source, sourceProperty, Mode, UpdateSource, Nothing, Validator, Converter, ErrorFormat)
End Function

' Создает экземпляр привязки свойства
Public Function Create(ByVal Manager As IBindingManager, ByVal Target As Object, ByVal TargetProperty As String, ByVal Source As Object, ByVal sourceProperty As String, _
Optional ByVal Mode As BindingMode = BindingMode.TwoWayBinding, _
Optional ByVal UpdateSource As BindingUpdateSourceTrigger = BindingUpdateSourceTrigger.OnPropertyChanged, _
Optional ByVal Handler As IHandlePropertyChanged, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal ErrorFormat As IValidationErrorFormatter) As IPropertyBinding
    
    GuardClauses.GuardNonDefaultInstance Me, PropertyBinding
    
    Dim result As New PropertyBinding
    
    Set result.Target = Target
    result.TargetProperty = TargetProperty
    
    Set result.Source = Source
    result.SourcePropertyPath = sourceProperty
    
    Set result.Manager = Manager
    Set result.Converter = Converter
    Set result.Validator = Validator
    Set result.ValidationErrorFormatter = ErrorFormat
    Set result.PropertyChangedHandler = Handler
    result.Mode = Mode
    result.UpdateSourceTrigger = UpdateSource
    
    If Mode <> OneWayToSource And TypeOf Source Is INotifyPropertyChanged Then
        Dim notifier As INotifyPropertyChanged
        Set notifier = Source
        notifier.RegisterHandler result
    End If
    
    Set Create = result
End Function

Public Property Get Source() As Object
    Set Source = this.Source
End Property

Public Property Set Source(ByVal RHS As Object)
    GuardClauses.GuardDoubleInitialization this.Source, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    Set this.Source = RHS
    
    ' Проверяем поддержку IViewModel для обработки ошибок валидации
    If TypeOf RHS Is IViewModel Then
        Dim viewModel As IViewModel
        Set viewModel = RHS
        Set this.ValidationErrorHandler = viewModel.Validation
    End If
End Property

Public Property Get Target() As Object
    Set Target = this.Target
End Property

Public Property Set Target(ByVal RHS As Object)
    GuardClauses.GuardDoubleInitialization this.Target, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    
    Set this.Target = RHS
    
    ' Очищаем предыдущий источник событий
    Set TextBoxEventSource = Nothing
    
    ' Проверяем, является ли цель TextBox
    If TypeOf RHS Is MSForms.TextBox Then
        Set TextBoxEventSource = RHS
    End If
End Property

Public Property Get PropertyChangedHandler() As IHandlePropertyChanged
    Set PropertyChangedHandler = this.PropertyChangedHandler
End Property

Public Property Set PropertyChangedHandler(ByVal RHS As IHandlePropertyChanged)
    Set this.PropertyChangedHandler = RHS
End Property

Public Property Get SourcePropertyPath() As String
    SourcePropertyPath = this.SourcePropertyPath
End Property

Public Property Let SourcePropertyPath(ByVal RHS As String)
    this.SourcePropertyPath = RHS
End Property

Public Property Get TargetProperty() As String
    TargetProperty = this.TargetProperty
End Property

Public Property Let TargetProperty(ByVal RHS As String)
    this.TargetProperty = RHS
End Property

Public Property Get Mode() As BindingMode
    Mode = this.Mode
End Property

Public Property Let Mode(ByVal RHS As BindingMode)
    this.Mode = RHS
End Property

Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
    UpdateSourceTrigger = this.UpdateSourceTrigger
End Property

Public Property Let UpdateSourceTrigger(ByVal RHS As BindingUpdateSourceTrigger)
    this.UpdateSourceTrigger = RHS
End Property

Public Property Get Validator() As IValueValidator
    Set Validator = this.Validator
End Property

Public Property Set Validator(ByVal RHS As IValueValidator)
    Set this.Validator = RHS
End Property

Public Property Get ValidationErrorFormatter() As IValidationErrorFormatter
    Set ValidationErrorFormatter = this.ValidationErrorFormatter
End Property

Public Property Set ValidationErrorFormatter(ByVal RHS As IValidationErrorFormatter)
    Set this.ValidationErrorFormatter = RHS
End Property

Public Property Get Converter() As IValueConverter
    Set Converter = this.Converter
End Property

Public Property Set Converter(ByVal RHS As IValueConverter)
    Set this.Converter = RHS
End Property

Public Property Get Manager() As IBindingManager
    Set Manager = this.Manager
End Property

Public Property Set Manager(ByVal RHS As IBindingManager)
    GuardClauses.GuardDoubleInitialization this.Manager, TypeName(Me)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    Set this.Manager = RHS
End Property

' Пытается прочитать значение свойства с объекта
Private Function TryReadPropertyValue(ByVal Instance As Object, ByVal PropertyName As String, ByRef outValue As Variant) As Boolean
    On Error Resume Next
    outValue = CallByName(Instance, PropertyName, VbGet)
    TryReadPropertyValue = (Err.Number = 0)
    On Error GoTo 0
End Function

' Пытается записать значение в свойство объекта
Private Function TryWritePropertyValue(ByVal Instance As Object, ByVal PropertyName As String, ByVal value As Variant) As Boolean
    On Error Resume Next
    CallByName Instance, PropertyName, VbLet, value
    TryWritePropertyValue = (Err.Number = 0)
    On Error GoTo 0
End Function

' Применяет значение из цели к источнику
Private Sub ApplyToSource()
    ' Односторонние привязки не записывают значения в источник
    If this.Mode = OneTimeBinding Or this.Mode = OneWayBinding Then Exit Sub
    
    Dim targetValue As Variant
    Dim PropertyName As String
    
    ' Читаем значение из целевого элемента управления
    PropertyName = this.TargetProperty
    If Not TryReadPropertyValue(this.Target, PropertyName, targetValue) Then
        Exit Sub
    End If
    
    ' Преобразуем значение, если есть конвертер
    If Not this.Converter Is Nothing Then targetValue = this.Converter.ConvertBack(targetValue)
    
    ' Проверяем валидность значения
    If Not ValidateValue(targetValue) Then Exit Sub
    
    ' Обновляем значение в источнике
    Dim sourceProperty As String
    sourceProperty = this.SourcePropertyPath
    
    Dim currentValue As Variant
    If Not TryReadPropertyValue(this.Source, sourceProperty, currentValue) Then
        Exit Sub
    End If
    
    ' Применяем значение, только если оно изменилось
    If targetValue <> currentValue Then
        If Not this.ValidationErrorHandler Is Nothing Then
            this.ValidationErrorHandler.ClearValidationErrors sourceProperty
        End If
        
        If Not TryWritePropertyValue(this.Source, sourceProperty, targetValue) Then
            Exit Sub
        End If
        
        ' Уведомляем об обновлении свойства
        If Not this.PropertyChangedHandler Is Nothing Then
            this.PropertyChangedHandler.OnPropertyChanged this.Source, sourceProperty
        End If
    End If
End Sub

' Проверяет валидность значения
Private Function ValidateValue(ByVal value As Variant) As Boolean
    If this.Validator Is Nothing Then
        ValidateValue = True
        Exit Function
    End If
    
    ' Проверяем значение с помощью валидатора
    If Not this.Validator.IsValid(value, this.Source, this.Target) Then
        ' Уведомляем об ошибке валидации
        If Not this.ValidationErrorHandler Is Nothing Then
            Dim me_ As IPropertyBinding
            Set me_ = Me
            this.ValidationErrorHandler.OnValidationError me_, this.Validator.Message
        End If
        
        ' Применяем стиль ошибки к элементу
        If Not this.ValidationErrorFormatter Is Nothing Then
            this.ValidationErrorFormatter.Apply this.SourcePropertyPath, this.Validator.Message
        End If
        
        ValidateValue = False
        Exit Function
    End If
    
    ' Восстанавливаем нормальный вид элемента
    If Not this.ValidationErrorFormatter Is Nothing Then
        this.ValidationErrorFormatter.Restore
    End If
    
    ValidateValue = True
End Function

' Применяет значение из
' Применяет значение из источника к цели
Private Sub ApplyToTarget()
    ' Односторонние привязки в направлении источника не задают значения в цель
    If this.Mode = OneWayToSource Then Exit Sub
    
    Dim sourceValue As Variant
    Dim PropertyName As String
    
    ' Читаем значение из источника
    PropertyName = this.SourcePropertyPath
    If Not TryReadPropertyValue(this.Source, PropertyName, sourceValue) Then
        Exit Sub
    End If
    
    ' Преобразуем значение, если есть конвертер
    If Not this.Converter Is Nothing Then sourceValue = this.Converter.Convert(sourceValue)
    
    ' Обновляем значение в целевом элементе
    Dim TargetProperty As String
    TargetProperty = this.TargetProperty
    
    Dim currentValue As Variant
    If Not TryReadPropertyValue(this.Target, TargetProperty, currentValue) Then
        Exit Sub
    End If
    
    ' Применяем значение, только если оно изменилось
    If currentValue <> sourceValue Then
        If Not TryWritePropertyValue(this.Target, TargetProperty, sourceValue) Then
            Exit Sub
        End If
    End If
End Sub

Private Sub Class_Initialize()
    Set this.Handlers = New Collection
End Sub

Private Sub TextBoxEventSource_Change()
    ' Если настроено обновление при изменении свойства
    If this.UpdateSourceTrigger = OnPropertyChanged Then
        ApplyToSource
    End If
End Sub

Private Sub TextBoxEventSource_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Если настроено обновление при выходе из элемента
    If this.UpdateSourceTrigger = OnExit Then
        ApplyToSource
    End If
End Sub

Private Sub TextBoxEventSource_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Если настроено обновление при нажатии клавиши
    If this.UpdateSourceTrigger = OnKeyPress Then
        ApplyToSource
    End If
End Sub

Private Sub IHandlePropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    If Source Is this.Source And PropertyName = this.SourcePropertyPath Then
        ApplyToTarget
    End If
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    Dim Handler As IHandlePropertyChanged
    For Each Handler In this.Handlers
        Handler.OnPropertyChanged Source, PropertyName
    Next
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    If this.Handlers Is Nothing Then Set this.Handlers = New Collection
    this.Handlers.Add Handler
End Sub

' Применяет привязку
Public Sub Apply()
    If Not this.Applied Then
        this.Applied = True
        ApplyToTarget
    End If
End Sub

' Реализация IPropertyBinding
Private Property Get IPropertyBinding_Converter() As IValueConverter
    Set IPropertyBinding_Converter = this.Converter
End Property

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = this.Mode
End Property

Private Property Get IPropertyBinding_Source() As Object
    Set IPropertyBinding_Source = this.Source
End Property

Private Property Get IPropertyBinding_SourcePropertyPath() As String
    IPropertyBinding_SourcePropertyPath = this.SourcePropertyPath
End Property

Private Property Get IPropertyBinding_Target() As Object
    Set IPropertyBinding_Target = this.Target
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = this.TargetProperty
End Property

Private Property Get IPropertyBinding_UpdateSourceTrigger() As BindingUpdateSourceTrigger
    IPropertyBinding_UpdateSourceTrigger = this.UpdateSourceTrigger
End Property

Private Sub IPropertyBinding_Apply()
    Apply
End Sub

