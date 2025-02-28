VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder "MVVM.Binding"
'@PredeclaredId
Option Explicit
Implements ICommandBinding

Private Type TCommandBinding
    viewModel As Object
    Target As Object
    Command As ICommand
End Type

' ќбъ€вление переменной WithEvents без атрибута
Private WithEvents CommandButtonEvents As MSForms.CommandButton
Attribute CommandButtonEvents.VB_VarHelpID = -1
Private this As TCommandBinding

' —оздает и возвращает реализацию ICommandBinding, св€зывающую указанную ICommand с указанной целью MSForms.CommandButton
Public Function Create(ByVal Target As Object, ByVal Command As ICommand, ByVal viewModel As Object) As ICommandBinding
    GuardClauses.GuardNonDefaultInstance Me, CommandBinding
    
    Dim result As CommandBinding
    Set result = New CommandBinding
    
    Set result.viewModel = viewModel
    Set result.Target = Target
    Set result.Command = Command
    
    Set Create = result
End Function

Public Property Get viewModel() As Object
    Set viewModel = this.viewModel
End Property

Public Property Set viewModel(ByVal RHS As Object)
    Set this.viewModel = RHS
End Property

Public Property Get Target() As Object
    Set Target = this.Target
End Property

Public Property Set Target(ByVal RHS As Object)
    GuardClauses.GuardDoubleInitialization this.Target, TypeName(Me)
    Set this.Target = RHS
    
    ' ”станавливаем источник событий в зависимости от типа
    If TypeOf RHS Is MSForms.CommandButton Then
        Set CommandButtonEvents = RHS
    End If
End Property

Public Property Get Command() As ICommand
    Set Command = this.Command
End Property

Public Property Set Command(ByVal RHS As ICommand)
    Set this.Command = RHS
    If Not RHS Is Nothing And Not this.Target Is Nothing Then
        this.Target.ControlTipText = RHS.Description
    End If
End Property

Private Sub OnExecute()
    If Not this.Command Is Nothing Then
        this.Command.Execute this.viewModel
    End If
End Sub

Private Property Get ICommandBinding_Target() As Object
    Set ICommandBinding_Target = this.Target
End Property

Private Property Get ICommandBinding_Command() As ICommand
    Set ICommandBinding_Command = this.Command
End Property

Private Sub ICommandBinding_EvaluateCanExecute(ByVal Context As Object)
    EvaluateCanExecute Context
End Sub

Private Sub EvaluateCanExecute(ByVal Source As Object)
    If this.Target Is Nothing Then Exit Sub
    If this.Command Is Nothing Then
        this.Target.Enabled = False
    Else
        this.Target.Enabled = this.Command.CanExecute(Source)
    End If
End Sub

Private Sub CommandButtonEvents_Click()
    OnExecute
End Sub
