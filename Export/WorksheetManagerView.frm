VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WorksheetManagerView 
   Caption         =   "UserForm1"
   ClientHeight    =   9750.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6990
   OleObjectBlob   =   "WorksheetManagerView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WorksheetManagerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Version 5#
'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WorksheetManagerView
'   Caption = "Менеджер листов"
'   ClientHeight = 5550
'   ClientLeft = 120
'   ClientTop = 456
'   ClientWidth = 6372
'   OleObjectBlob   =   "WorksheetManagerView.frx":0000
'   StartUpPosition = 1    'CenterOwner
'End
'Attribute VB_Name = "WorksheetManagerView"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
'@Folder "MVVM.View"
'@PredeclaredId
Option Explicit
Implements IView

Private Type TView
    viewModel As WorksheetViewModel
    BindingManager As IBindingManager
End Type

Private this As TView

' Создает экземпляр представления, используя указанные ViewModel и BindingManager
Public Function Create(ByVal viewModel As WorksheetViewModel, ByVal BindingManager As IBindingManager) As IView
    GuardClauses.GuardNonDefaultInstance Me, WorksheetManagerView
    
    Dim result As WorksheetManagerView
    Set result = New WorksheetManagerView
    
    Set result.viewModel = viewModel
    Set result.BindingManager = BindingManager
    
    result.InitializeBindings
    
    Set Create = result
End Function

' Инициализирует привязки данных
Public Sub InitializeBindings()
    ' Привязка списка листов
    this.BindingManager.BindPropertyPath this.viewModel, "WorksheetsList", lstWorksheets, "List"
    this.BindingManager.BindPropertyPath this.viewModel, "SelectedIndex", lstWorksheets, "ListIndex"
    
    ' Привязка текстовых полей
    this.BindingManager.BindPropertyPath this.viewModel, "NewWorksheetName", txtNewSheetName, "Text"
    this.BindingManager.BindPropertyPath this.viewModel, "SelectedWorksheetName", lblSelectedSheet, "Caption"
    this.BindingManager.BindPropertyPath this.viewModel, "SelectedWorkbookName", lblSelectedBook, "Caption"
    this.BindingManager.BindPropertyPath this.viewModel, "StatusMessage", lblStatus, "Caption"
    
    ' Привязка команд к кнопкам
    this.BindingManager.BindCommand this.viewModel, btnAddWorksheet, this.viewModel.AddCommand
    this.BindingManager.BindCommand this.viewModel, btnDeleteWorksheet, this.viewModel.DeleteCommand
    this.BindingManager.BindCommand this.viewModel, btnRenameWorksheet, this.viewModel.RenameCommand
    this.BindingManager.BindCommand this.viewModel, btnRefresh, this.viewModel.RefreshCommand
    
    ' Применяем все привязки
    this.BindingManager.ApplyBindings this.viewModel
End Sub

' Свойство для доступа к ViewModel
Public Property Get viewModel() As WorksheetViewModel
    Set viewModel = this.viewModel
End Property

Public Property Set viewModel(ByVal RHS As WorksheetViewModel)
    GuardClauses.GuardDoubleInitialization this.viewModel, TypeName(Me)
    Set this.viewModel = RHS
End Property

' Свойство для доступа к BindingManager
Public Property Get BindingManager() As IBindingManager
    Set BindingManager = this.BindingManager
End Property

Public Property Set BindingManager(ByVal RHS As IBindingManager)
    GuardClauses.GuardDoubleInitialization this.BindingManager, TypeName(Me)
    Set this.BindingManager = RHS
End Property

' Обработка событий представления
Private Sub btnClose_Click()
    Me.Hide
End Sub

' Обработка выбора в списке (дополнительное обновление визуальных элементов)
Private Sub lstWorksheets_Click()
    ' Обновляем состояние команд, так как изменился выбор
    this.BindingManager.OnEvaluateCanExecute this.viewModel
End Sub

' Реализация интерфейса IView
Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = this.viewModel
End Property

Private Sub IView_Show()
    Me.Show vbModeless
End Sub

Private Function IView_ShowDialog() As Boolean
    Me.Show vbModal
    IView_ShowDialog = True ' Result could be set based on dialog outcome
End Function

Private Sub IView_Hide()
    Me.Hide
End Sub
