Attribute VB_Name = "App"
Option Explicit

Public Sub UpdateDocumentSituation(FileName As String, Situation As String)
    Dim NewData As New Dictionary
    NewData("IND_SITU") = Situation
    DB.Table("DOCUMENTO").Update NewData, "NOM_ARCH = '" & FileName & "'"
End Sub

Public Function GetDocumentSituation(FileName As String) As String
    Dim Document As Collection
    Set Document = DB.Table("DOCUMENTO").Where("NOM_ARCH", "=", FileName).GetAll
    GetDocumentSituation = IIf(Document.Count > 0, Document(1)("IND_SITU"), Empty)
End Function

Public Function GetDocumentObservation(FileName As String) As String
    Dim Document As Collection
    Dim Observation As String
    Set Document = DB.Table("DOCUMENTO").Where("NOM_ARCH", "=", FileName).GetAll
    Observation = IIf(Document.Count > 0, Document(1)("DES_OBSE"), Empty)
    GetDocumentObservation = IIf(Observation = "-", Empty, Observation)
End Function

Public Function CreateDocumentEntity(Invoice As InvoiceEntity) As DocumentEntity
    Dim Document As New DocumentEntity
    Dim FileName As String
    Dim Situation As String
    Dim Observation As String
    
    FileName = CreateFileName(Invoice.DocType, Invoice.DocSerie, Invoice.DocNumber)
    Situation = GetSituationFromCode(GetDocumentSituation(FileName))
    Observation = GetDocumentObservation(FileName)
    
    Document.Emission = Invoice.EmissionDate
    Document.DocType = Invoice.DocType
    Document.DocSerie = Invoice.DocSerie
    Document.DocNumber = Invoice.DocNumber
    Document.CustomerDocType = Invoice.Customer.DocType
    Document.CustomerDocNumber = Invoice.Customer.DocNumber
    Document.CustomerName = Invoice.Customer.Name
    Document.TypeCurrency = Invoice.TypeCurrency
    Document.SubTotal = Invoice.SubTotal
    Document.Igv = Invoice.Igv
    Document.Total = Invoice.Total
    Document.Situation = Situation
    Document.Observation = Observation
    
    Set CreateDocumentEntity = Document
End Function

Public Function CreateCanceledDocumentEntity(Document As DocumentEntity, Motivo As String, SendingDate As Date) As CanceledDocumentEntity
    Dim CanceledDocument As New CanceledDocumentEntity
    
    CanceledDocument.GenerationDate = Document.Emission
    CanceledDocument.CommunicationDate = SendingDate
    CanceledDocument.DocType = Document.DocType
    CanceledDocument.DocNumber = Document.DocSerie & "-" & Format(Document.DocNumber, "00000000")
    CanceledDocument.Motivo = Motivo
    
    Set CreateCanceledDocumentEntity = CanceledDocument
End Function

Public Function GetDigestValue(DocType As String, DocSerie As String, DocNumber As Long) As String
    On Error GoTo HandleErrors
    Dim fs As New FileSystemObject
    Dim XmlDocument As New DOMDocument60
    Dim FileName As String
    Dim ZipPath As String
    Dim XmlName As String
    Dim ZipName As String
    
    FileName = CreateFileName(DocType, DocSerie, DocNumber)
    ZipName = FileName & ".zip"
    ZipPath = fs.BuildPath(Prop.Sfs.ENVIOPath, ZipName)
    XmlName = FileName & ".xml"
    XmlDocument.LoadXML GetXmlContentFromZip(ZipPath, XmlName)
    XmlDocument.SetProperty "SelectionNamespaces", "xmlns:ds='http://www.w3.org/2000/09/xmldsig#'"
    
    GetDigestValue = XmlDocument.SelectSingleNode("//ds:DigestValue").text
    DebugLog "Se obtiene el DigestValue del documento: " & FileName, "GetDigestValue"
    Exit Function
HandleErrors:
    ErrorLog "No se pudo obtener el DigestValue del documento: " & FileName, "GetDigestValue", Err.Number
End Function

Public Function GetResponseSunat(DocType As String, DocSerie As String, DocNumber As Long) As Dictionary
    On Error GoTo HandleErrors
    Dim fs As New FileSystemObject
    Dim XmlDocument As New DOMDocument60
    Dim FileName As String
    Dim ZipName As String
    Dim ZipPath As String
    Dim XmlName As String
    Dim Response As New Dictionary
    
    FileName = CreateFileName(DocType, DocSerie, DocNumber)
    ZipName = "R" & FileName & ".zip"
    ZipPath = fs.BuildPath(Prop.Sfs.RPTAPath, ZipName)
    XmlName = "R-" & FileName & ".xml"
    XmlDocument.LoadXML GetXmlContentFromZip(ZipPath, XmlName)
    XmlDocument.SetProperty "SelectionNamespaces", "xmlns:cbc='urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2'"
    
    Response.Add "ReferenceID", XmlDocument.SelectSingleNode("//cbc:ReferenceID").text
    Response.Add "ResponseCode", XmlDocument.SelectSingleNode("//cbc:ResponseCode").text
    Response.Add "Description", XmlDocument.SelectSingleNode("//cbc:Description").text
    
    Set GetResponseSunat = Response
    DebugLog XmlDocument.SelectSingleNode("//cbc:Description").text, "GetResponseSunat"
    Exit Function
HandleErrors:
    ErrorLog "Error al consultar la respuesta de la SUNAT del documento: " & FileName, "GetResponseSunat", Err.Number
End Function

Public Function ElectronicDocumentExists(DocumentNumber As String) As Boolean
    Dim fs As New FileSystemObject
    Dim ZipDocumentPath As String
    Dim FileName As String
    
    FileName = Prop.Company.Ruc & "-" & DocumentNumber
    ZipDocumentPath = fs.BuildPath(Prop.Sfs.ENVIOPath, FileName & ".zip")
    ElectronicDocumentExists = fs.FileExists(ZipDocumentPath)
End Function

Public Function ResponseFileExists(DocType As String, DocSerie As String, DocNumber As Long) As Boolean
    Dim fs As New FileSystemObject
    Dim ZipDocumentPath As String
    Dim FileName As String
    
    FileName = CreateFileName(DocType, DocSerie, DocNumber)
    ZipDocumentPath = fs.BuildPath(Prop.Sfs.RPTAPath, "R" & FileName & ".zip")
    ResponseFileExists = fs.FileExists(ZipDocumentPath)
End Function

Sub RunSfs()
    Dim fs As New FileSystemObject
    Dim ScriptPath As String

    ScriptPath = fs.BuildPath(Prop.Sfs.Path, "EjecutarSFS.bat")
    
    If Not fs.FolderExists(Prop.Sfs.Path) Or Not fs.FileExists(ScriptPath) Then
        MsgBox "No se pudo ejecutar el Facturador SUNAT. Verifique que la ruta del Facturador SUNAT esté configurado.", vbInformation, "Subsane la observación"
        Exit Sub
    End If
    
    Shell "cmd /k " & Left(Prop.Sfs.Path, 2) & " & cd " & Prop.Sfs.Path & " & " & ScriptPath & " & exit"
    InfoLog "Ejecutando EjecutarSFS.bat"
End Sub

Public Function CreateFileName(DocType As String, DocSerie As String, DocNumber As Long) As String
    CreateFileName = Prop.Company.Ruc & "-" & DocType & "-" & DocSerie & "-" & Format(DocNumber, "00000000")
End Function

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

Public Function SfsIsRunning() As Boolean
    Dim EndPoint As String
    EndPoint = "http://localhost:" & Prop.Sfs.Port & "/api/ActualizarPantalla.htm"
    If Post(EndPoint, "{""txtSecuencia"": ""000""}") Then
        SfsIsRunning = True
    Else
        WarnLog "Al parecer facturadorApp.jar no se está ejecutando.", "SfsIsRunning"
        SfsIsRunning = False
    End If
End Function

' Registra en BDFacturador.db el documento JSON ubicado en DATA
Public Sub RefreshSfsScreen()
    Dim EndPoint As String
    EndPoint = "http://localhost:" & Prop.Sfs.Port & "/api/ActualizarPantalla.htm"
    Post EndPoint, "{""txtSecuencia"": ""000""}"
End Sub

' Elimina todas las filas de la tabla DOCUMENTO en BDFacturador.db
Public Sub RemoveSfsTray()
    Dim EndPoint As String
    EndPoint = "http://localhost:" & Prop.Sfs.Port & "/api/EliminarBandeja.htm"
    Post EndPoint, "{""rutaCertificado"": """"}"
End Sub

Public Sub CreatePdf(DocumentNumber As String)
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
    ThisWorkbook.FollowHyperlink PdfPath
    Exit Sub
HandleErrors:
    ErrorLog "El archivo " & FileName & " no existe.", "OpenPdf"
End Sub
