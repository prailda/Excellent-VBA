VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValueConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "IValueConverter"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'@Folder "MVVM.Interfaces"
'@Interface
Option Explicit

' ����������� �������� �� ��������� � �������� ��������
Public Function Convert(ByVal value As Variant) As Variant
End Function

' ����������� �������� �� ���� � ���������
Public Function ConvertBack(ByVal value As Variant) As Variant
End Function
