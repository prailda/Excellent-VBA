VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GuardClauses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


' Константа для пользовательских ошибок
Private Const CustomError As Long = vbObjectError Or 32

' Вызывает ошибку времени выполнения, если указанное логическое выражение равно True
Public Sub GuardExpression(ByVal Throw As Boolean, _
Optional ByVal Source As String = vbNullString, _
Optional ByVal Message As String = "Invalid procedure call or argument.")
    If Throw Then Err.Raise CustomError, Source, Message
End Sub

' Вызывает ошибку времени выполнения, если указанный экземпляр не является экземпляром по умолчанию
Public Sub GuardNonDefaultInstance(ByVal Instance As Object, ByVal defaultInstance As Object, _
Optional ByVal Source As String = vbNullString, _
Optional ByVal Message As String = "Method should be invoked from the default/predeclared instance of this class.")
    Debug.Assert TypeName(Instance) = TypeName(defaultInstance)
    GuardExpression Not Instance Is defaultInstance, Source, Message
End Sub

' Вызывает ошибку времени выполнения, если указанная ссылка на объект уже установлена
Public Sub GuardDoubleInitialization(ByVal Instance As Object, _
Optional ByVal Source As String = vbNullString, _
Optional ByVal Message As String = "Object is already initialized.")
    GuardExpression Not Instance Is Nothing, Source, Message
End Sub

' Вызывает ошибку времени выполнения, если указанная ссылка на объект равна Nothing
Public Sub GuardNullReference(ByVal Instance As Object, _
Optional ByVal Source As String = vbNullString, _
Optional ByVal Message As String = "Object reference cannot be Nothing.")
    GuardExpression Instance Is Nothing, Source, Message
End Sub

' Вызывает ошибку времени выполнения, если указанная строка пуста
Public Sub GuardEmptyString(ByVal value As String, _
Optional ByVal Source As String = vbNullString, _
Optional ByVal Message As String = "String cannot be empty.")
    GuardExpression value = vbNullString, Source, Message
End Sub

