Attribute VB_Name = "Main"
'Attribute VB_Name = "Main"
'@Folder "MVVM"
Option Explicit

' ���������� ���� ������������ ����� ��� ���������� ������� � �����������
Private Type MODULECOMPONENTS
    WorksheetViewModel As WorksheetViewModel
    WorksheetManagerView As WorksheetManagerView
    BindingManager As BindingManager
    WorksheetModel As WorksheetModel
    ValidationManager As ValidationManager
End Type

' ����� ����� � ����������
Public Sub ShowWorksheetManager()
    ' ������� � ��������� ���������� �� �������� MVVM
    Dim view As IView
    Set view = InitializeWorksheetManager
    view.Show
End Sub

' ������������� ����������� MVVM
Private Function InitializeWorksheetManager() As IView
    ' �������� ���� ����������� �����������
    Dim viewModel As WorksheetViewModel
    Dim bindingManagerInstance As BindingManager
    
    ' ������� ViewModel
    Set viewModel = WorksheetViewModel.Create
    
    ' ������� BindingManager
    Set bindingManagerInstance = BindingManager.Create
    
    ' ������� View � ��������� � ViewModel
    Dim result As IView
    Set result = WorksheetManagerView.Create(viewModel, bindingManagerInstance)
    
    ' �������������� ������
    viewModel.RefreshData
    viewModel.StatusMessage = "������ � ������"
    
    Set InitializeWorksheetManager = result
End Function

' ������� ��� ������������
Public Sub TestWorksheetManager()
    ShowWorksheetManager
End Sub
