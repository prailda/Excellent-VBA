VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSOGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsFSOGeneral - класс для общих операций с файловой системой
Option Explicit

' Общее свойство для работы с диском
Private pDriveLetter As String
Private pFSOBase As clsFSOBase

Private Sub Class_Initialize()
    Set pFSOBase = New clsFSOBase
End Sub

' Устанавливаем букву диска (например, "C:\")
Public Property Let DriveLetter(ByVal Value As String)
    ' Очистка: удаляем пробелы и убеждаемся, что строка заканчивается знаком "\"
    Dim tmp As String
    tmp = Trim(Value)
    If Right(tmp, 1) <> "\" Then
        tmp = tmp & "\"
    End If
    pDriveLetter = tmp
End Property

Public Property Get DriveLetter() As String
    DriveLetter = pDriveLetter
End Property

' Получение объекта Drive через FSO
Private Function GetDriveObject() As Object
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set GetDriveObject = fso.GetDrive(pDriveLetter)
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOGeneral.GetDriveObject", Err.Number, Err.Description, _
                   "Невозможно получить объект диска: " & pDriveLetter
    Set GetDriveObject = Nothing
End Function

' Метод получения свободного места на диске (в байтах)
Public Function GetAvailableSpace() As Currency
    On Error GoTo ErrorHandler
    
    Dim drv As Object
    Set drv = GetDriveObject()
    
    If drv Is Nothing Then
        GetAvailableSpace = -1
    Else
        GetAvailableSpace = drv.availableSpace
    End If
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOGeneral.GetAvailableSpace", Err.Number, Err.Description, _
                   "Ошибка при получении свободного места для " & pDriveLetter
    GetAvailableSpace = -1
End Function

' Метод получения общего размера диска (в байтах)
Public Function GetTotalSize() As Currency
    On Error GoTo ErrorHandler
    
    Dim drv As Object
    Set drv = GetDriveObject()
    
    If drv Is Nothing Then
        GetTotalSize = -1
    Else
        GetTotalSize = drv.totalSize
    End If
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOGeneral.GetTotalSize", Err.Number, Err.Description, _
                   "Ошибка при получении общего размера для " & pDriveLetter
    GetTotalSize = -1
End Function

' Метод получения занятого пространства (рассчитывается как TotalSize - AvailableSpace)
Public Function GetUsedSpace() As Currency
    Dim tot As Currency, avail As Currency
    
    tot = GetTotalSize()
    avail = GetAvailableSpace()
    
    If tot = -1 Or avail = -1 Then
        GetUsedSpace = -1
    Else
        GetUsedSpace = tot - avail
    End If
End Function

' Проверка, достаточно ли свободного места для требуемого объёма (в байтах)
Public Function IsEnoughFreeSpace(ByVal requiredSpace As Currency) As Boolean
    Dim avail As Currency
    
    avail = GetAvailableSpace()
    
    If avail = -1 Then
        IsEnoughFreeSpace = False
    Else
        IsEnoughFreeSpace = (avail >= requiredSpace)
    End If
End Function

' Метод форматирования байтов в человекочитаемый вид
Public Function FormatBytes(ByVal bytes As Currency) As String
    FormatBytes = pFSOBase.FormatSize(bytes)
End Function

' Метод получения типа диска
Public Function GetDriveType() As FSODriveType
    On Error GoTo ErrorHandler
    
    Dim drv As Object
    Set drv = GetDriveObject()
    
    If drv Is Nothing Then
        GetDriveType = DRIVE_TYPE_UNKNOWN
    Else
        GetDriveType = drv.driveType
    End If
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOGeneral.GetDriveType", Err.Number, Err.Description, _
                   "Ошибка при получении типа диска для " & pDriveLetter
    GetDriveType = DRIVE_TYPE_UNKNOWN
End Function

' Метод определения, готов ли диск
Public Function IsDriveReady() As Boolean
    On Error GoTo ErrorHandler
    
    Dim drv As Object
    Set drv = GetDriveObject()
    
    If drv Is Nothing Then
        IsDriveReady = False
    Else
        IsDriveReady = drv.IsReady
    End If
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOGeneral.IsDriveReady", Err.Number, Err.Description, _
                   "Ошибка при проверке готовности диска " & pDriveLetter
    IsDriveReady = False
End Function

' Метод получения текстового представления типа диска
Public Function GetDriveTypeText() As String
    Dim driveType As FSODriveType
    driveType = GetDriveType()
    
    Select Case driveType
        Case DRIVE_TYPE_UNKNOWN
            GetDriveTypeText = "Неизвестный"
        Case DRIVE_TYPE_REMOVABLE
            GetDriveTypeText = "Съемный"
        Case DRIVE_TYPE_FIXED
            GetDriveTypeText = "Фиксированный"
        Case DRIVE_TYPE_NETWORK
            GetDriveTypeText = "Сетевой"
        Case DRIVE_TYPE_CDROM
            GetDriveTypeText = "CD-ROM"
        Case DRIVE_TYPE_RAMDISK
            GetDriveTypeText = "RAM-диск"
        Case Else
            GetDriveTypeText = "Неизвестный тип: " & driveType
    End Select
End Function

' Метод получения серийного номера диска
Public Function GetDriveSerialNumber() As String
    On Error GoTo ErrorHandler
    
    Dim drv As Object
    Set drv = GetDriveObject()
    
    If drv Is Nothing Then
        GetDriveSerialNumber = ""
    Else
        GetDriveSerialNumber = drv.SerialNumber
    End If
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOGeneral.GetDriveSerialNumber", Err.Number, Err.Description, _
                   "Ошибка при получении серийного номера диска " & pDriveLetter
    GetDriveSerialNumber = ""
End Function

' Метод получения процента свободного места
Public Function GetFreeSpacePercent() As Double
    Dim totalSize As Currency
    Dim availableSpace As Currency
    
    totalSize = GetTotalSize()
    availableSpace = GetAvailableSpace()
    
    If totalSize <= 0 Or availableSpace < 0 Then
        GetFreeSpacePercent = -1
    Else
        GetFreeSpacePercent = Round((availableSpace / totalSize) * 100, 2)
    End If
End Function

' Метод для получения списка всех доступных дисков
Public Function GetAllDrives() As Collection
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim result As Collection
    Set result = New Collection
    
    Dim drv As Object
    For Each drv In fso.drives
        result.Add drv.DriveLetter
    Next drv
    
    Set GetAllDrives = result
    
    Exit Function
    
ErrorHandler:
    Dim logger As clsFSOErrorLogger
    Set logger = GetFSOErrorLoggerInstance()
    logger.LogError "clsFSOGeneral.GetAllDrives", Err.Number, Err.Description, _
                   "Ошибка при получении списка дисков"
    Set GetAllDrives = New Collection
End Function

