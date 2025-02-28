Attribute VB_Name = "Main"
'Attribute VB_Name = "Main"
'@Folder "MVVM"
Option Explicit

' Объявление всех используемых типов для устранения проблем с компиляцией
Private Type MODULECOMPONENTS
    WorksheetViewModel As WorksheetViewModel
    WorksheetManagerView As WorksheetManagerView
    BindingManager As BindingManager
    WorksheetModel As WorksheetModel
    ValidationManager As ValidationManager
End Type

' Точка входа в приложение
Public Sub ShowWorksheetManager()
    ' Создаем и запускаем приложение по паттерну MVVM
    Dim view As IView
    Set view = InitializeWorksheetManager
    view.Show
End Sub

' Инициализация компонентов MVVM
Private Function InitializeWorksheetManager() As IView
    ' Создание всех необходимых компонентов
    Dim viewModel As WorksheetViewModel
    Dim bindingManagerInstance As BindingManager
    
    ' Создаем ViewModel
    Set viewModel = WorksheetViewModel.Create
    
    ' Создаем BindingManager
    Set bindingManagerInstance = BindingManager.Create
    
    ' Создаем View с привязкой к ViewModel
    Dim result As IView
    Set result = WorksheetManagerView.Create(viewModel, bindingManagerInstance)
    
    ' Инициализируем данные
    viewModel.RefreshData
    viewModel.StatusMessage = "Готово к работе"
    
    Set InitializeWorksheetManager = result
End Function

' Функция для тестирования
Public Sub TestWorksheetManager()
    ShowWorksheetManager
End Sub
