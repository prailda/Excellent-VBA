VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "IView"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' �������� ������ ViewModel, ��������� � ��������������
Public Property Get viewModel() As Object
End Property

' ���������� �������������
Public Sub Show()
End Sub

' ���������� ������������� � ��������� ������
Public Function ShowDialog() As Boolean
End Function

' �������� �������������
Public Sub Hide()
End Sub
