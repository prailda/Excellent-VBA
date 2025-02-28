VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CancelCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder "MVVM.Commands"
'@PredeclaredId
Option Explicit
Implements ICommand

Private Type TState
    view As ICancellable
End Type

Private this As TState

' —оздает новый экземпл€р этой команды
Public Function Create(ByVal view As ICancellable) As ICommand
    Dim result As CancelCommand
    Set result = New CancelCommand
    Set result.view = view
    Set Create = result
End Function

Public Property Get view() As ICancellable
    Set view = this.view
End Property

Public Property Set view(ByVal RHS As ICancellable)
    GuardClauses.GuardDoubleInitialization this.view, TypeName(Me)
    Set this.view = RHS
End Property

Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    ICommand_CanExecute = True
End Function

Private Property Get ICommand_Description() As String
    ICommand_Description = "ќтменить изменени€ и закрыть."
End Property

Private Sub ICommand_Execute(ByVal Context As Object)
    this.view.OnCancel
End Sub
