VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' ������������ ��� ������� ��������
Public Enum BindingMode
    TwoWayBinding         ' ������������ �������� (��������� � ����� ������������)
    OneWayBinding         ' ������������� �������� (�� ��������� � ����)
    OneWayToSource        ' ������������� �������� (�� ���� � ���������)
    OneTimeBinding        ' ����������� �������� (������ ��� �������������)
End Enum

' ������������ ��� ��������� ���������� ���������
Public Enum BindingUpdateSourceTrigger
    OnPropertyChanged     ' ��������� ��� ��������� ��������
    OnKeyPress            ' ��������� ��� ������� �������
    OnExit                ' ��������� ��� ������ �� �������� ����������
End Enum

' �������� ��������, ����������� �����/����������� ��������
Public Property Get Mode() As BindingMode
End Property

' �������� ��������, ����������� ������� ���������� ��������
Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
End Property

' �������� ������ ViewModel, ������� �������� ���������� ��������
Public Property Get Source() As Object
End Property

' �������� ���� � ��������� Public Property Get, ����������� �� ��������� ��������
Public Property Get SourcePropertyPath() As String
End Property

' �������� ������� ����������������� ����������, ������� �������� ����� ��������
Public Property Get Target() As Object
End Property

' �������� ��� ������������ �������� �� ������� �������
Public Property Get TargetProperty() As String
End Property

' �������� ��������� ��� ������������� (���� ����) ��� ���������� ��������
Public Property Get Converter() As IValueConverter
End Property

' ��������� ��������
Public Sub Apply()
End Sub
