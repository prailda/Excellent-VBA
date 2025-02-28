VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "ValidationManager"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
'@Folder "MVVM.Validation"
Option Explicit
Implements IHandleValidationError

Public Event PropertyChanged(ByVal Source As Object, ByVal PropertyName As String)

Private Type TValidationState
    Errors As Collection
    PropertyErrors As Object  ' Dictionary
    Message As String
End Type

Private this As TValidationState

' Создает новый экземпляр ValidationManager
Public Function Create() As IHandleValidationError
    GuardClauses.GuardNonDefaultInstance Me, ValidationManager
    
    Dim result As New ValidationManager
    Set result.Errors = New Collection
    Set result.PropertyErrors = CreateObject("Scripting.Dictionary")
    
    Set Create = result
End Function

' Свойства доступа
Public Property Get Errors() As Collection
    Set Errors = this.Errors
End Property

Public Property Set Errors(ByVal value As Collection)
    Set this.Errors = value
End Property

Public Property Get PropertyErrors() As Object
    Set PropertyErrors = this.PropertyErrors
End Property

Public Property Set PropertyErrors(ByVal value As Object)
    Set this.PropertyErrors = value
End Property

Public Property Get Message() As String
    If this.Errors.Count > 0 Then
        Message = this.Errors(1)
    Else
        Message = ""
    End If
End Property

' Инициализация объекта
Private Sub Class_Initialize()
    Set this.Errors = New Collection
    Set this.PropertyErrors = CreateObject("Scripting.Dictionary")
End Sub

' Очистка ошибок валидации для свойства
Private Sub IHandleValidationError_ClearValidationErrors(ByVal PropertyName As String)
    If this.PropertyErrors.Exists(PropertyName) Then
        this.PropertyErrors.Remove PropertyName
        RebuildErrorsList
        RaiseEvent PropertyChanged(Me, PropertyName)
    End If
End Sub

' Вызывается когда привязка обнаруживает ошибку валидации
Private Sub IHandleValidationError_OnValidationError(ByVal Source As IPropertyBinding, ByVal Message As String)
    Dim PropertyName As String
    PropertyName = Source.SourcePropertyPath
    
    ' Добавление ошибки в словарь
    If Not this.PropertyErrors.Exists(PropertyName) Then
        this.PropertyErrors.Add PropertyName, Message
    Else
        this.PropertyErrors(PropertyName) = Message
    End If
    
    RebuildErrorsList
    RaiseEvent PropertyChanged(Me, PropertyName)
End Sub

' Перестраивает список ошибок на основе словаря
Private Sub RebuildErrorsList()
    Set this.Errors = New Collection
    
    Dim key As Variant
    For Each key In this.PropertyErrors.Keys
        this.Errors.Add this.PropertyErrors(key)
    Next key
End Sub

' Возвращает True если нет ошибок валидации
Private Property Get IHandleValidationError_IsValid() As Boolean
    IHandleValidationError_IsValid = (this.Errors.Count = 0)
End Property

' Возвращает первое сообщение об ошибке
Private Property Get IHandleValidationError_Message() As String
    If this.Errors.Count > 0 Then
        IHandleValidationError_Message = this.Errors(1)
    Else
        IHandleValidationError_Message = ""
    End If
End Property

' Возвращает все сообщения об ошибках как массив
Private Property Get IHandleValidationError_ValidationErrors() As Variant
    Dim result() As String
    Dim i As Long
    
    If this.Errors.Count = 0 Then
        ReDim result(0 To 0)
        result(0) = ""
    Else
        ReDim result(0 To this.Errors.Count - 1)
        
        For i = 1 To this.Errors.Count
            result(i - 1) = this.Errors(i)
        Next i
    End If
    
    IHandleValidationError_ValidationErrors = result
End Property
