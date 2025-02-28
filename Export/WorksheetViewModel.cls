VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "WorksheetViewModel"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
'@Folder "MVVM.ViewModel"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

' Событие изменения свойства
Public Event PropertyChanged(ByVal Source As Object, ByVal PropertyName As String)

Private Type TViewModel
    Model As WorksheetModel
    Handlers As Collection
    
    ' Свойства ViewModel
    SelectedIndex As Long
    SelectedWorksheetName As String
    SelectedWorkbookName As String
    NewWorksheetName As String
    WorksheetsList As Collection
    ActiveWorkbooks As Collection
    StatusMessage As String
End Type

Private this As TViewModel
'Private ValidationHandler As Object  ' Это будет ValidationManager

'' Фабричный метод для создания экземпляра
'Public Function Create() As IViewModel
'    GuardClauses.GuardNonDefaultInstance Me, WorksheetViewModel
'
'    Dim result As New WorksheetViewModel
'    Set result.Model = New WorksheetModel
'
'    ' Инициализация обработчика валидации
'    Set result.ValidationHandler = ValidationManager.Create()
'
'    ' Инициализация коллекций
'    Set result.WorksheetsList = New Collection
'    Set result.ActiveWorkbooks = New Collection
'
'    Set Create = result
'End Function

' Свойство для доступа к модели
Public Property Get Model() As WorksheetModel
    Set Model = this.Model
End Property

Public Property Set Model(ByVal RHS As WorksheetModel)
    GuardClauses.GuardDoubleInitialization this.Model, TypeName(Me)
    Set this.Model = RHS
End Property

' Свойство для доступа к ValidationHandler
Public Property Get ValidationHandler() As Object
    Set ValidationHandler = this.ValidationHandler
End Property

Public Property Set ValidationHandler(ByVal RHS As Object)
    Set this.ValidationHandler = RHS
End Property

' Получить коллекцию листов
Public Property Get WorksheetsList() As Collection
    Set WorksheetsList = this.WorksheetsList
End Property

' Обновление данных из модели
Public Sub RefreshData()
    On Error Resume Next
    
    ' Очищаем текущие данные
    Set this.WorksheetsList = New Collection
    Set this.ActiveWorkbooks = New Collection
    
    ' Загружаем данные из модели
    Dim worksheets As Collection
    Set worksheets = this.Model.GetWorksheets()
    
    ' Заполняем список листов
    Dim i As Long
    For i = 1 To worksheets.Count
        this.WorksheetsList.Add worksheets(i).name
    Next i
    
    ' Сбросить выбранный индекс
    this.SelectedIndex = 0
    this.SelectedWorksheetName = ""
    this.SelectedWorkbookName = ""
    
    ' Устанавливаем сообщение о статусе
    this.StatusMessage = "Данные о листах обновлены"
    
    ' Уведомляем об изменениях
    OnPropertyChanged Me, "WorksheetsList"
    OnPropertyChanged Me, "SelectedIndex"
    OnPropertyChanged Me, "SelectedWorksheetName"
    OnPropertyChanged Me, "SelectedWorkbookName"
    OnPropertyChanged Me, "StatusMessage"
    
    On Error GoTo 0
End Sub

' Выбранный индекс листа
Public Property Get SelectedIndex() As Long
    SelectedIndex = this.SelectedIndex
End Property

Public Property Let SelectedIndex(ByVal value As Long)
    ' Меняем только если индекс изменился
    If this.SelectedIndex <> value Then
        this.SelectedIndex = value
        
        ' Обновляем связанные свойства
        If value > 0 And value <= this.WorksheetsList.Count Then
            this.SelectedWorksheetName = this.WorksheetsList(value)
            
            ' Получаем имя книги для выбранного листа
            Dim wsObj As Object
            Set wsObj = this.Model.GetWorksheetByIndex(value)
            If Not wsObj Is Nothing Then
                this.SelectedWorkbookName = wsObj.Parent.name
            End If
        Else
            this.SelectedWorksheetName = ""
            this.SelectedWorkbookName = ""
        End If
        
        ' Уведомляем об изменениях
        OnPropertyChanged Me, "SelectedIndex"
        OnPropertyChanged Me, "SelectedWorksheetName"
        OnPropertyChanged Me, "SelectedWorkbookName"
    End If
End Property

' Имя выбранного листа
Public Property Get SelectedWorksheetName() As String
    SelectedWorksheetName = this.SelectedWorksheetName
End Property

Public Property Let SelectedWorksheetName(ByVal value As String)
    If this.SelectedWorksheetName <> value Then
        this.SelectedWorksheetName = value
        OnPropertyChanged Me, "SelectedWorksheetName"
    End If
End Property

' Имя книги выбранного листа
Public Property Get SelectedWorkbookName() As String
    SelectedWorkbookName = this.SelectedWorkbookName
End Property

Public Property Let SelectedWorkbookName(ByVal value As String)
    If this.SelectedWorkbookName <> value Then
        this.SelectedWorkbookName = value
        OnPropertyChanged Me, "SelectedWorkbookName"
    End If
End Property

' Имя нового листа
Public Property Get NewWorksheetName() As String
    NewWorksheetName = this.NewWorksheetName
End Property

Public Property Let NewWorksheetName(ByVal value As String)
    If this.NewWorksheetName <> value Then
        this.NewWorksheetName = value
        OnPropertyChanged Me, "NewWorksheetName"
    End If
End Property

' Статусное сообщение
Public Property Get StatusMessage() As String
    StatusMessage = this.StatusMessage
End Property

Public Property Let StatusMessage(ByVal value As String)
    If this.StatusMessage <> value Then
        this.StatusMessage = value
        OnPropertyChanged Me, "StatusMessage"
    End If
End Property

' Добавление нового листа
Public Function AddNewWorksheet(ByVal sheetName As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Len(Trim$(sheetName)) = 0 Then
        ' Генерируем имя по умолчанию
        sheetName = "Лист" & Format(Now, "yyyymmdd_hhnnss")
    End If
    
    ' Добавляем лист через модель
    If this.Model.AddWorksheet(sheetName) Then
        RefreshData
        
        ' Находим индекс нового листа
        Dim i As Long
        For i = 1 To this.WorksheetsList.Count
            If this.WorksheetsList(i) = sheetName Then
                this.SelectedIndex = i
                Exit For
            End If
        Next i
        
        this.StatusMessage = "Лист '" & sheetName & "' успешно добавлен"
        OnPropertyChanged Me, "StatusMessage"
        AddNewWorksheet = True
    End If
    
    Exit Function
    
ErrorHandler:
    this.StatusMessage = "Ошибка при добавлении листа: " & Err.Description
    OnPropertyChanged Me, "StatusMessage"
    AddNewWorksheet = False
End Function

' Удаление выбранного листа
Public Function DeleteSelectedWorksheet() As Boolean
    On Error GoTo ErrorHandler
    
    If this.SelectedIndex <= 0 Then
        this.StatusMessage = "Не выбран лист для удаления"
        OnPropertyChanged Me, "StatusMessage"
        DeleteSelectedWorksheet = False
        Exit Function
    End If
    
    Dim wsName As String
    wsName = this.SelectedWorksheetName
    
    ' Удаляем лист через модель
    If this.Model.DeleteWorksheet(this.SelectedIndex) Then
        this.StatusMessage = "Лист '" & wsName & "' успешно удален"
        OnPropertyChanged Me, "StatusMessage"
        RefreshData
        DeleteSelectedWorksheet = True
    End If
    
    Exit Function
    
ErrorHandler:
    this.StatusMessage = "Ошибка при удалении листа: " & Err.Description
    OnPropertyChanged Me, "StatusMessage"
    DeleteSelectedWorksheet = False
End Function

' Переименование выбранного листа
Public Function RenameSelectedWorksheet(ByVal newName As String) As Boolean
    On Error GoTo ErrorHandler
    
    If this.SelectedIndex <= 0 Then
        this.StatusMessage = "Не выбран лист для переименования"
        OnPropertyChanged Me, "StatusMessage"
        RenameSelectedWorksheet = False
        Exit Function
    End If
    
    If Len(Trim$(newName)) = 0 Then
        this.StatusMessage = "Новое имя не может быть пустым"
        OnPropertyChanged Me, "StatusMessage"
        RenameSelectedWorksheet = False
        Exit Function
    End If
    
    Dim oldName As String
    oldName = this.SelectedWorksheetName
    
    ' Переименовываем лист через модель
    If this.Model.RenameWorksheet(this.SelectedIndex, newName) Then
        this.StatusMessage = "Лист '" & oldName & "' переименован в '" & newName & "'"
        OnPropertyChanged Me, "StatusMessage"
        RefreshData
        
        ' Находим индекс переименованного листа
        Dim i As Long
        For i = 1 To this.WorksheetsList.Count
            If this.WorksheetsList(i) = newName Then
                this.SelectedIndex = i
                Exit For
            End If
        Next i
        
        RenameSelectedWorksheet = True
    End If
    
    Exit Function
    
ErrorHandler:
    this.StatusMessage = "Ошибка при переименовании листа: " & Err.Description
    OnPropertyChanged Me, "StatusMessage"
    RenameSelectedWorksheet = False
End Function

' Уведомление об изменении свойства
Public Sub OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    RaiseEvent PropertyChanged(Source, PropertyName)
End Sub

' Реализация интерфейса INotifyPropertyChanged
Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    OnPropertyChanged Source, PropertyName
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    If this.Handlers Is Nothing Then Set this.Handlers = New Collection
    this.Handlers.Add Handler
End Sub

' Реализация интерфейса IViewModel
Private Property Get IViewModel_Validation() As IHandleValidationError
    Set IViewModel_Validation = this.ValidationHandler
End Property

' Обработка изменений свойств ValidationHandler
Private Sub ValidationHandler_PropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    If TypeOf Source Is ValidationManager Then
        this.StatusMessage = "Ошибка валидации: " & Source.Message
        OnPropertyChanged Me, "StatusMessage"
    End If
End Sub
    
    ' Добавьте эти свойства и методы в класс WorksheetViewModel

' Свойства для команд
Private Type TCommands
    AddCommand As ICommand
    DeleteCommand As ICommand
    RenameCommand As ICommand
    RefreshCommand As ICommand
End Type

Private cmdSet As TCommands

' Инициализация команд (добавьте этот код в метод Create)
Public Function Create() As IViewModel
    GuardClauses.GuardNonDefaultInstance Me, WorksheetViewModel
    
    Dim result As New WorksheetViewModel
    Set result.Model = New WorksheetModel
    
    ' Инициализация обработчика валидации
    Set result.ValidationHandler = ValidationManager.Create()
    
    ' Инициализация коллекций
    Set result.WorksheetsList = New Collection
    Set result.ActiveWorkbooks = New Collection
    
    ' Инициализация команд
    Set result.AddCommand = AddWorksheetCommand.Create(result)
    Set result.DeleteCommand = DeleteWorksheetCommand.Create(result)
    Set result.RenameCommand = RenameWorksheetCommand.Create(result)
    Set result.RefreshCommand = RefreshWorksheetCommand.Create(result)
    
    Set Create = result
End Function

' Свойства для доступа к командам
Public Property Get AddCommand() As ICommand
    Set AddCommand = cmdSet.AddCommand
End Property

Public Property Set AddCommand(ByVal RHS As ICommand)
    Set cmdSet.AddCommand = RHS
End Property

Public Property Get DeleteCommand() As ICommand
    Set DeleteCommand = cmdSet.DeleteCommand
End Property

Public Property Set DeleteCommand(ByVal RHS As ICommand)
    Set cmdSet.DeleteCommand = RHS
End Property

Public Property Get RenameCommand() As ICommand
    Set RenameCommand = cmdSet.RenameCommand
End Property

Public Property Set RenameCommand(ByVal RHS As ICommand)
    Set cmdSet.RenameCommand = RHS
End Property

Public Property Get RefreshCommand() As ICommand
    Set RefreshCommand = cmdSet.RefreshCommand
End Property

Public Property Set RefreshCommand(ByVal RHS As ICommand)
    Set cmdSet.RefreshCommand = RHS
End Property
