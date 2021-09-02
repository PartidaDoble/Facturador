Attribute VB_Name = "Services"
Option Explicit

Public Function Post(Url As String, Body As String) As Boolean
    On Error GoTo HandleErrors
    Dim ClientHttp As New XMLHTTP60
    
    ClientHttp.Open "POST", Url, False
    ClientHttp.setRequestHeader "Content-Type", "application/json"
    ClientHttp.send Body
    
    DebugLog "Status " & ClientHttp.Status & " " & Url, "Post"
    Post = ClientHttp.Status = 201
    Exit Function
HandleErrors:
    ErrorLog Url, "Post"
End Function

Public Sub CreatePdf(Document As Object)
    If Prop.App.Premium Then
        If Prop.App.PrintMode = PrintSfs Then
            CreatePdfSfs Document.Id
        ElseIf Prop.App.PrintMode = PrintA4 Then
            Run "FillA4", Document
            Run "SavePdf", Document.Id, PrintA4
        ElseIf Prop.App.PrintMode = PrintTicket Then
            Run "FillTicket", Document
            Run "SavePdf", Document.Id, PrintTicket
        End If
    Else
        CreatePdfSfs Document.Id
    End If
End Sub

Public Sub CreatePdfSfs(DocumentNumber As String)
    Dim EndPoint As String
    Dim FileName As String
    
    EndPoint = "http://localhost:" & Prop.Sfs.Port & "/api/MostrarXml.htm"
    FileName = Prop.Company.Ruc & "-" & DocumentNumber
    Post EndPoint, "{""nomArch"": """ & FileName & """}"
End Sub

Public Sub OpenPdf(DocumentNumber As String)
    On Error GoTo HandleErrors
    Dim fs As New FileSystemObject
    Dim FileName As String
    Dim PdfPath As String
    
    FileName = Prop.Company.Ruc & "-" & DocumentNumber & ".pdf"
    PdfPath = fs.BuildPath(Prop.Sfs.REPOPath, FileName)
    Shell "cmd /k " & PdfPath & " & exit"
    Exit Sub
HandleErrors:
    ErrorLog "El archivo " & FileName & " no existe.", "OpenPdf"
End Sub

Public Function SendEmail(ReceiverEmail As String, Subject As String, HtmlBody As String, Attachments As Collection) As Boolean
    If Prop.Email.Provider = GmailProv Then
        SendEmail = Run("SendEmailGmail", ReceiverEmail, Subject, HtmlBody, Attachments)
    ElseIf Prop.Email.Provider = OutlookProv Then
        SendEmail = Run("SendEmailOutlook", ReceiverEmail, Subject, HtmlBody, Attachments)
    End If
End Function

Public Function GetDigestValue(DocumentNumber As String) As String
    On Error GoTo HandleErrors
    Dim XmlDocument As New DOMDocument60
    Dim FileName As String
    Dim ZipPath As String
    Dim XmlName As String
    Dim ZipName As String
    
    FileName = Prop.Company.Ruc & "-" & DocumentNumber
    ZipName = FileName & ".zip"
    ZipPath = PathJoin(Prop.Sfs.ENVIOPath, ZipName)
    XmlName = FileName & ".xml"
    XmlDocument.LoadXML GetXmlContentFromZip(ZipPath, XmlName)
    XmlDocument.SetProperty "SelectionNamespaces", "xmlns:ds='http://www.w3.org/2000/09/xmldsig#'"
    
    GetDigestValue = XmlDocument.SelectSingleNode("//ds:DigestValue").text
    DebugLog "Se obtiene el DigestValue del documento: " & FileName, "GetDigestValue"
    Exit Function
HandleErrors:
    ErrorLog "No se pudo obtener el DigestValue del documento: " & FileName, "GetDigestValue", Err.Number
End Function

Public Function ElectronicDocumentExists(ZipFileName As String) As Boolean
    Dim fs As New FileSystemObject
    ElectronicDocumentExists = fs.FileExists(PathJoin(Prop.Sfs.ENVIOPath, ZipFileName))
End Function

Sub RunSfs()
    Dim ScriptPath As String
    ScriptPath = PathJoin(Prop.Sfs.Path, "EjecutarSFS.bat")
    Shell "cmd /k " & Left(Prop.Sfs.Path, 2) & " & cd " & Prop.Sfs.Path & " & " & ScriptPath & " & exit"
    InfoLog "Ejecutando EjecutarSFS.bat"
End Sub

Public Function SfsPrepared() As Boolean
    SfsPrepared = True
    
    If SfsBatExists Then
        If Not SfsIsRunning Then
            RunSfs
            MsgBox "El Facturador SUNAT no se está ejecutando.", vbCritical, ""
            MsgBox "Ejecutando el Facturador SUNAT automáticamente..." & Chr(13) & Chr(13) & _
                   "Espere a que el Facturador SUNAT esté listo. Esto puede demorar alrededor de un minuto (depende de la velocidad de su computadora)", vbInformation, "Ejecutando el Facturador SUNAT"
            SfsPrepared = False
        End If
    Else
        MsgBox "Debe configurar la ruta del Facturador SUNAT.", vbCritical, "Subsane la observación"
        SfsPrepared = False
    End If
End Function

Public Function SfsBatExists() As Boolean
    Dim fs As New FileSystemObject
    SfsBatExists = fs.FileExists(PathJoin(Prop.Sfs.Path, "EjecutarSFS.bat"))
End Function

Public Function SfsIsRunning() As Boolean
    Dim EndPoint As String
    
    EndPoint = "http://localhost:" & Prop.Sfs.Port & "/api/ActualizarPantalla.htm"
    If Post(EndPoint, "{""txtSecuencia"": ""000""}") Then
        SfsIsRunning = True
    Else
        WarnLog "El facturadorApp.jar no se está ejecutando.", "SfsIsRunning"
        SfsIsRunning = False
    End If
End Function

' Carga en BDFacturador.db los archivos json que se encuentran en la carpeta DATA
Public Sub RefreshSfsScreen()
    Dim EndPoint As String
    EndPoint = "http://localhost:" & Prop.Sfs.Port & "/api/ActualizarPantalla.htm"
    Post EndPoint, "{""txtSecuencia"": ""000""}"
End Sub

Public Sub RemoveSfsTray()
    Dim EndPoint As String
    EndPoint = "http://localhost:" & Prop.Sfs.Port & "/api/EliminarBandeja.htm"
    Post EndPoint, "{""rutaCertificado"": """"}"
End Sub

Function GetTicketNumber(Observation As String)
    GetTicketNumber = IIf(Left(Observation, 12) = "Nro. Ticket:", Mid(Observation, InStrRev(Observation, " ") + 1), Empty)
End Function

Public Sub DbUpdateDocumentSituation(FileName As String, Situation As String)
    On Error GoTo HandleErrors
    Dim Data As New Dictionary
    
    Data("IND_SITU") = Situation
    DB.Table("DOCUMENTO").Update Data, "NOM_ARCH = '" & FileName & "'"
    Exit Sub
HandleErrors:
    ErrorLog "Error al consultar la base de datos BDFacturador " & FileName, "DbUpdateDocumentSituation", Err.Number
End Sub

Public Function DbGetDocumentSituation(FileName As String) As String
    On Error GoTo HandleErrors
    Dim Document As Collection
    
    Set Document = DB.Table("DOCUMENTO").Where("NOM_ARCH", "=", FileName).GetAll
    DbGetDocumentSituation = IIf(Document.Count > 0, Document(1)("IND_SITU"), Empty)
    Exit Function
HandleErrors:
    ErrorLog "Error al consultar la base de datos BDFacturador " & FileName, "DbGetDocumentSituation", Err.Number
End Function

Public Function DbGetDocumentObservation(FileName As String) As String
    On Error GoTo HandleErrors
    Dim Document As Collection
    Dim Observation As String
    
    Set Document = DB.Table("DOCUMENTO").Where("NOM_ARCH", "=", FileName).GetAll
    Observation = IIf(Document.Count > 0, Document(1)("DES_OBSE"), Empty)
    DbGetDocumentObservation = IIf(Observation = "-", Empty, Observation)
    Exit Function
HandleErrors:
    ErrorLog "Error al consultar la base de datos BDFacturador " & FileName, "DbGetDocumentObservation", Err.Number
End Function

Public Sub DbDeleteRow(FileName As String)
    On Error GoTo HandleErrors
    DB.Table("DOCUMENTO").Delete "NOM_ARCH = '" & FileName & "'"
    Exit Sub
HandleErrors:
    ErrorLog "Error al consultar la base de datos BDFacturador " & FileName, "DbGetDocumentObservation", Err.Number
End Sub

Public Sub TraceLog(Message As String, Optional From As String = Empty)
    SettingLog
    LogTrace Message, From
End Sub

Public Sub DebugLog(Message As String, Optional From As String = Empty)
    SettingLog
    LogDebug Message, From
End Sub

Public Sub InfoLog(Message As String, Optional From As String = Empty)
    SettingLog
    LogInfo Message, From
End Sub

Public Sub WarnLog(Message As String, Optional From As String = Empty)
    SettingLog
    LogWarn Message, From
End Sub

Public Sub ErrorLog(Message As String, Optional From As String = Empty, Optional ErrNumber As Long = 0)
    SettingLog
    LogError Message, From, ErrNumber
End Sub

Private Sub SettingLog()
    If Prop.App.Env = EnvProduction Then
        If Prop.App.LogLevel > 0 Then
            LogCallback = "LogFile"
            LogThreshold = Prop.App.LogLevel
        End If
    Else
        LogEnabled = True
    End If
End Sub

Public Sub LogFile(Level As Long, Message As String, From As String)
    On Error Resume Next
    Dim fs As New FileSystemObject
    Dim Stream As TextStream
    Dim LevelName As String
    Dim LogFilePath As String
    Dim DirPath As String
    
    Select Case Level
        Case 1
            LevelName = "Trace"
        Case 2
            LevelName = "Debug"
        Case 3
            LevelName = "Info "
        Case 4
            LevelName = "WARN "
        Case 5
            LevelName = "ERROR"
    End Select
    
    DirPath = PathJoin(ThisWorkbook.Path, "logs")
    If Not fs.FolderExists(DirPath) Then fs.CreateFolder DirPath

    LogFilePath = fs.BuildPath(DirPath, Format(Date, "yyyy-mm-dd") & ".txt")
    If fs.FileExists(LogFilePath) Then
        Set Stream = fs.OpenTextFile(LogFilePath, ForAppending)
    Else
        Set Stream = fs.CreateTextFile(LogFilePath)
    End If
    
    Stream.WriteLine Format(Now, "dd/mm/yyyy hh:mm:ss") & " " & LevelName & " - " & IIf(From <> "", From & ": ", "") & Message
    Stream.Close
End Sub

Public Function GetSituationFromEnum(Situation As SituationEnum) As String
    Select Case Situation
        Case CdpPorGenerarXml
            GetSituationFromEnum = "01 - por generar xml"
        Case CdpXmlGenerado
            GetSituationFromEnum = "02 - xml generado"
        Case CdpEnviadoAceptado
            GetSituationFromEnum = "03 - enviado y aceptado sunat"
        Case CdpEnviadoAceptadoConObs
            GetSituationFromEnum = "04 - enviado y aceptado sunat con obs"
        Case CdpRechazado
            GetSituationFromEnum = "05 - rechazado por sunat"
        Case CdpConErrores
            GetSituationFromEnum = "06 - con errores"
        Case CdpPorValidarXml
            GetSituationFromEnum = "07 - por validar xml"
        Case CdpEnviadoPorProcesar
            GetSituationFromEnum = "08 - enviado a sunat por procesar"
        Case CdpEnviadoProcesando
            GetSituationFromEnum = "09 - enviado a sunat procesando"
        Case CdpRechazado10
            GetSituationFromEnum = "10 - rechazado por sunat"
        Case CdpEnviadoAceptado11
            GetSituationFromEnum = "11 - enviado y aceptado sunat"
        Case CdpEnviadoAceptadoConObs12
            GetSituationFromEnum = "12 - enviado y aceptado sunat"
    End Select
End Function

Public Function GetSituationFromCode(Code As String) As SituationEnum
    Select Case Code
        Case "01"
            GetSituationFromCode = CdpPorGenerarXml
        Case "02"
            GetSituationFromCode = CdpXmlGenerado
        Case "03"
            GetSituationFromCode = CdpEnviadoAceptado
        Case "04"
            GetSituationFromCode = CdpEnviadoAceptadoConObs
        Case "05"
            GetSituationFromCode = CdpRechazado
        Case "06"
            GetSituationFromCode = CdpConErrores
        Case "07"
            GetSituationFromCode = CdpPorValidarXml
        Case "08"
            GetSituationFromCode = CdpEnviadoPorProcesar
        Case "09"
            GetSituationFromCode = CdpEnviadoProcesando
        Case "10"
            GetSituationFromCode = CdpRechazado10
        Case "11"
            GetSituationFromCode = CdpEnviadoAceptado11
        Case "12"
            GetSituationFromCode = CdpEnviadoAceptadoConObs12
    End Select
End Function
