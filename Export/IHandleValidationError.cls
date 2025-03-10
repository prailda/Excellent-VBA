VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHandleValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "IHandleValidationError"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' ���������� ��������� IsValid ��� ���������� ��������
Public Sub ClearValidationErrors(ByVal PropertyName As String)
End Sub

' �������� ������, ���������� ��� ��������� �� ������� ���������
Public Property Get ValidationErrors() As Variant
End Property

' ���������� ��� ������������� ������ ��������� ��� ���������� �������� ��������
Public Sub OnValidationError(ByVal Source As IPropertyBinding, ByVal Message As String)
End Sub

' ���������� True, ���� ��� ������ ���������
Public Property Get IsValid() As Boolean
End Property

' ���������� ������ ��������� �� ������
Public Property Get Message() As String
End Property
