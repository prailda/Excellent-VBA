VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddWorksheetCommand"
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

Private Property Get ICommand_Description() As String
    ICommand_Description = "Добавить новый лист в выбранную книгу"
End Property

' Свойство View
Public Property Get view() As IView
    Set view = this.view
End Property

Public Property Set view(ByVal RHS As IView)
    Set this.view = RHS
End Property

' Создает новый экземпляр команды
Public Function Create(ByVal view As IView) As ICommand
    Dim result As AddWorksheetCommand
    Set result = New AddWorksheetCommand
    Set result.view = view
    Set Create = result
End Function

' Реализация ICommand
Private Function ICommand_CanExecute(ByVal Context As Object) As Boolean
    Dim viewModel As WorksheetViewModel
    
    If TypeOf Context Is WorksheetViewModel Then
        Set viewModel = Context
        ' Команда доступна, если есть хотя бы одна рабочая книга
        ICommand_CanExecute = (viewModel.ActiveWorkbooks.Count > 0)
    Else
        ICommand_CanExecute = False
    End If
End Function

Private Sub ICommand_Execute(ByVal Context As Object)
    On Error Resume Next
    
    Dim viewModel As WorksheetViewModel
    Set viewModel = Context
    
    ' Запрашиваем имя для нового листа
    Dim newName As String
    If viewModel.NewWorksheetName <> "" Then
        newName = viewModel.NewWorksheetName
    Else
        newName = InputBox("Введите имя нового листа:", "Добавление листа")
        If newName = "" Then Exit Sub ' Пользователь отменил
    End If
    
    ' Если выбрана книга, добавляем лист в эту книгу
    If viewModel.SelectedWorkbookName <> "" Then
        viewModel.AddNewWorksheet viewModel.SelectedWorkbookName, newName
    Else
        ' Если книга не выбрана, но есть доступные книги, используем первую
        If viewModel.ActiveWorkbooks.Count > 0 Then
            viewModel.AddNewWorksheet viewModel.ActiveWorkbooks(1), newName
        Else
            viewModel.StatusMessage = "Нет доступных книг для добавления листа"
        End If
    End If
    
    ' Очищаем поле ввода нового имени
    viewModel.NewWorksheetName = ""
End Sub
