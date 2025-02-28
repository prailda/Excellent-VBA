VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeleteWorksheetCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder "MVVM.Commands"
'@PredeclaredId
Option Explicit
Implements ICommand

Private Type TState
    view As IView
End Type

Private this As TState

' Создает новый экземпляр команды
Public Function Create(ByVal view As IView) As ICommand
    Dim result As DeleteWorksheetCommand
    Set result = New DeleteWorksheetCommand
    Set result.view = view
    Set Create = result
End Function

' Свойство View
Public Property Get view() As IView
    Set view = this.view
End Property

Public Property Set view(ByVal RHS As IView)
    Set this.view = RHS
End Property

' Реализация ICommand
Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    Dim viewModel As WorksheetViewModel
    
    If TypeOf Context Is WorksheetViewModel Then
        Set viewModel = Context
        ' Команда доступна, только если выбран лист
        ICommand_CanExecute = (viewModel.SelectedWorksheetName <> "")
    Else
        ICommand_CanExecute = False
    End If
End Function

Private Property Get ICommand_Description() As String
    ICommand_Description = "Удалить выбранный лист"
End Property

Private Sub ICommand_Execute(ByVal Context As Object)
    On Error Resume Next
    
    Dim viewModel As WorksheetViewModel
    Set viewModel = Context
    
    ' Подтверждение удаления
    Dim response As VbMsgBoxResult
    response = MsgBox("Вы уверены, что хотите удалить лист '" & viewModel.SelectedWorksheetName & "'?", _
                     vbQuestion + vbYesNo, "Подтверждение удаления")
    
    If response = vbYes Then
        viewModel.DeleteSelectedWorksheet
    End If
End Sub

