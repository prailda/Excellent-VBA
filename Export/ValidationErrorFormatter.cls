VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationErrorFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "ValidationErrorFormatter"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
'@Folder "MVVM.Validation"
'@PredeclaredId
Option Explicit
Implements IValidationErrorFormatter

Private Type TState
    Target As Object
    DefaultBackColor As Long
    DefaultForeColor As Long
    HasError As Boolean
End Type

Private this As TState

' ��������� ��� ������
Private Const ErrorBackColor As Long = &HC0C0FF   ' ������-�������
Private Const ErrorForeColor As Long = &H80&      ' �����-�������

' ������� ����� ��������� ValidationErrorFormatter
Public Function Create(ByVal Target As Object) As IValidationErrorFormatter
    GuardClauses.GuardNonDefaultInstance Me, ValidationErrorFormatter
    
    Dim result As New ValidationErrorFormatter
    Set result.Target = Target
    
    Set Create = result
End Function

' �������� ��� ������� � �������� ��������
Public Property Get Target() As Object
    Set Target = this.Target
End Property

Public Property Set Target(ByVal RHS As Object)
    GuardClauses.GuardNullReference RHS, TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.Target, TypeName(Me)
    
    Set this.Target = RHS
    
    ' ��������� ����������� ����� ��������
    SaveDefaultColors
End Property

' ��������� ����������� ����� ��������
Private Sub SaveDefaultColors()
    On Error Resume Next
    
    If TypeOf this.Target Is MSForms.TextBox Then
        this.DefaultBackColor = this.Target.BackColor
        this.DefaultForeColor = this.Target.ForeColor
    ElseIf TypeOf this.Target Is MSForms.ComboBox Then
        this.DefaultBackColor = this.Target.BackColor
        this.DefaultForeColor = this.Target.ForeColor
    End If
    
    On Error GoTo 0
End Sub

' ���������� IValidationErrorFormatter
Private Property Get IValidationErrorFormatter_Target() As Object
    Set IValidationErrorFormatter_Target = this.Target
End Property

' ��������� ����� ������ � ��������
Private Sub IValidationErrorFormatter_Apply(ByVal PropertyName As String, ByVal Message As String)
    On Error Resume Next
    
    If TypeOf this.Target Is MSForms.TextBox Then
        this.Target.BackColor = ErrorBackColor
        this.Target.ForeColor = ErrorForeColor
        this.Target.ControlTipText = Message
        this.HasError = True
    ElseIf TypeOf this.Target Is MSForms.ComboBox Then
        this.Target.BackColor = ErrorBackColor
        this.Target.ForeColor = ErrorForeColor
        this.Target.ControlTipText = Message
        this.HasError = True
    End If
    
    On Error GoTo 0
End Sub

' ��������������� ����������� ����� ��������
Private Sub IValidationErrorFormatter_Restore()
    On Error Resume Next
    
    If this.HasError Then
        If TypeOf this.Target Is MSForms.TextBox Then
            this.Target.BackColor = this.DefaultBackColor
            this.Target.ForeColor = this.DefaultForeColor
            this.Target.ControlTipText = ""
        ElseIf TypeOf this.Target Is MSForms.ComboBox Then
            this.Target.BackColor = this.DefaultBackColor
            this.Target.ForeColor = this.DefaultForeColor
            this.Target.ControlTipText = ""
        End If
        
        this.HasError = False
    End If
    
    On Error GoTo 0
End Sub
