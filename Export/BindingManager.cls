VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "BindingManager"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
'@Folder "MVVM.Binding"
'@PredeclaredId
Option Explicit
Implements IBindingManager
Implements IHandlePropertyChanged

Private Type TState
    Handlers As Collection
    CommandBindings As Collection
    PropertyBindings As Collection
End Type

Private this As TState

Public Function Create() As IBindingManager
    GuardClauses.GuardNonDefaultInstance Me, BindingManager
    Set Create = New BindingManager
End Function

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is BindingManager
End Property

Private Sub ApplyAllBindings(ByVal Source As Object)
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me), "Member call is invalid against stateless default instance."
    Dim Binding As IPropertyBinding
    For Each Binding In this.PropertyBindings
        Binding.Apply
    Next
    EvaluateAllCommands Source
End Sub

Private Sub EvaluateAllCommands(ByVal Source As Object)
    Dim Binding As ICommandBinding
    For Each Binding In this.CommandBindings
        Binding.EvaluateCanExecute Source
    Next
End Sub

Private Sub Class_Initialize()
    If Not IsDefaultInstance Then
        Set this.Handlers = New Collection
        Set this.CommandBindings = New Collection
        Set this.PropertyBindings = New Collection
    End If
End Sub

Private Sub IBindingManager_BindPropertyPath(ByVal Source As Object, ByVal PropertyPath As String, ByVal Target As Object, _
Optional ByVal TargetProperty As String = "", _
Optional ByVal Mode As BindingMode = TwoWayBinding, _
Optional ByVal Validator As IValueValidator, _
Optional ByVal Converter As IValueConverter, _
Optional ByVal ErrorFormat As IValidationErrorFormatter)
    
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me), "Member call is invalid against stateless default instance."
    
    Dim Binding As IPropertyBinding
    ' Здесь проблема: неправильный параметр.
    ' Вместо объекта OnPropertyChanged (которого нет), передаем Me как IHandlePropertyChanged
    Set Binding = PropertyBinding.Create(Me, Target, TargetProperty, Source, PropertyPath, Mode, Me, Validator, Converter, ErrorFormat)
    
    this.PropertyBindings.Add Binding
    this.Handlers.Add Binding
End Sub

Private Sub IBindingManager_BindCommand(ByVal Source As Object, ByVal Target As Object, ByVal Command As ICommand)
    
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me), "Member call is invalid against stateless default instance."
    
    Dim Binding As ICommandBinding
    Set Binding = CommandBinding.Create(Target, Command, Source)
    
    this.CommandBindings.Add Binding
    
End Sub

Private Sub IBindingManager_ApplyBindings(ByVal Source As Object)
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me), "Member call is invalid against stateless default instance."
    ApplyAllBindings Source
End Sub

Private Sub IBindingManager_OnEvaluateCanExecute(ByVal Source As Object)
    GuardClauses.GuardExpression IsDefaultInstance, TypeName(Me), "Member call is invalid against stateless default instance."
    EvaluateAllCommands Source
End Sub

Private Sub IHandlePropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    EvaluateAllCommands Source
End Sub

