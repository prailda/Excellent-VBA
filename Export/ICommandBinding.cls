VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' �������� ������-�������� �������, ����������� � �������
Public Property Get Target() As Object
End Property

' �������� �������, ����������� � ��������� �������
Public Property Get Command() As ICommand
End Property

' ���������, ����� �� ������� ����������� � ������ ��������� ��������
Public Sub EvaluateCanExecute(ByVal Context As Object)
End Sub
