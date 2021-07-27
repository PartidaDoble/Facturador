Attribute VB_Name = "Core"
Option Explicit

Public Sub CheckTicketStatus()
    Dim CanceledDocument As Object
    Dim FileName As String
    Dim Situation As String
    Dim CanceledDocumentRepo As New CanceledDocumentRepository
    
    For Each CanceledDocument In CanceledDocumentRepo.GetAll
        If CanceledDocument.Situation = CdpEnviadoPorProcesar Then
            FileName = Prop.Company.Ruc & "-" & CanceledDocument.Id
            Situation = GetDocumentSituation(FileName)
            If Situation <> "08" Then
                CanceledDocument.Situation = GetSituationFromCode(Situation)
                CanceledDocument.Observation = GetDocumentObservation(FileName)
                CanceledDocumentRepo.Update CanceledDocument
            End If
        End If
    Next CanceledDocument
End Sub

Public Sub CancelDocument(DocumentNumber As String, Motivo As String, SendingDate As Date)
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    Dim CanceledDocuments As New Collection
    Dim CanceledDocument As CanceledDocumentEntity
    Dim Correlative As Long
    Dim CanceledDocumentNumber As String
    Dim FileName As String
    Dim CanceledDocumentRepo As New CanceledDocumentRepository

    Set Document = DocumentRepo.GetItem(DocumentNumber)
    If Not Document Is Nothing Then
        If Document.Situation = CdpEnviadoAceptado Or Document.Situation = CdpEnviadoAceptadoConObs Then
            Set CanceledDocument = CreateCanceledDocumentEntity(Document, Motivo, SendingDate)
            CanceledDocuments.Add CanceledDocument
            Correlative = WorksheetFunction.CountIf(sheetCanceledDocuments.Columns(3), SendingDate) + 1
            
            CanceledDocumentNumber = "RA-" & Format(SendingDate, "yyyymmdd") & "-" & Format(Correlative, "000")
            CreateFileJsonCancelDocument CanceledDocuments, CanceledDocumentNumber
            RefreshSfsScreen
            GenerateElectronicDocument "RA", CanceledDocumentNumber
            If ElectronicDocumentExists(CanceledDocumentNumber) Then
                SendElectronicDocument "RA", CanceledDocumentNumber
                
                FileName = Prop.Company.Ruc & "-" & CanceledDocumentNumber
                CanceledDocument.Correlative = Correlative
                CanceledDocument.Situation = GetSituationFromCode(GetDocumentSituation(FileName))
                CanceledDocument.Observation = GetDocumentObservation(FileName)
                
                CanceledDocumentRepo.Add CanceledDocument
                MsgBox "La solicitud de anulación del comprobante " & DocumentNumber & " fue enviado a la SUNAT correctamente.", vbInformation, ""
                InfoLog "La solicitud de anulación del comprobante " & DocumentNumber & " fue enviado a la SUNAT correctamente.", "CancelDocument"
            Else
                MsgBox "Error al generar el xml para anular el comprobante " & DocumentNumber, vbCritical, ""
                ErrorLog "Error al generar el xml para anular el comprobante " & DocumentNumber, "CancelDocument"
            End If
        Else
            MsgBox "Para anular el comprobante " & DocumentNumber & " debe tener la situación ""03 - enviado y aceptado sunat"" o ""04 - enviado y aceptado sunat con obs.""", vbExclamation, ""
            WarnLog "Para anular el comprobante " & DocumentNumber & " debe tener la situación 03 o 04", "CancelDocument"
        End If
    Else
        MsgBox "Para anular el comprobante " & DocumentNumber & " debe estar registrado en la hoja Comprobantes.", vbExclamation, ""
        WarnLog "Para anular el comprobante " & DocumentNumber & " debe estar registrado en la hoja Comprobantes.", "CancelDocument"
    End If
End Sub

Public Sub CreateFileJsonCancelDocument(Documents As Collection, DocumentNumber As String)
    On Error GoTo HandleErrors
    Dim fs As New FileSystemObject
    Dim Stream As TextStream
    Dim FileName As String
    Dim DocumentPath As String

    FileName = Prop.Company.Ruc & "-" & DocumentNumber & ".json"
    DocumentPath = fs.BuildPath(Prop.Sfs.DATAPath, FileName)
    Set Stream = fs.CreateTextFile(DocumentPath)
    Stream.Write CanceledDocumentsToJson(Documents)
    Stream.Close
    DebugLog "El archivo JSON " & FileName & " para anular comprobantes fue creado correctamente.", "CreateFileJsonCancelDocument"
    Exit Sub
HandleErrors:
    ErrorLog "Error al crear el archivo JSON " & FileName & " para anular comprobantes.", "CreateFileJsonCancelDocument", Err.Number
End Sub

Public Function CanceledDocumentsToJson(CanceledDocuments As Collection, Optional Pretty As Boolean = True) As String
    Dim Data As New Dictionary
    Dim Document As Dictionary
    Dim Documents As New Collection
    Dim CanceledDocument As Variant
    
    For Each CanceledDocument In CanceledDocuments
        Set Document = New Dictionary
        Document.Add "fecGeneracion", Format(CanceledDocument.GenerationDate, "yyyy-mm-dd")
        Document.Add "fecComunicacion", Format(CanceledDocument.CommunicationDate, "yyyy-mm-dd")
        Document.Add "tipDocBaja", CanceledDocument.DocType
        Document.Add "numDocBaja", CanceledDocument.DocNumber
        Document.Add "desMotivoBaja", CanceledDocument.Motivo
        Documents.Add Document
    Next CanceledDocument
    
    Data.Add "resumenBajas", Documents
    
    If Pretty Then
        CanceledDocumentsToJson = ConvertToJson(Data, 2)
    Else
        CanceledDocumentsToJson = ConvertToJson(Data)
    End If
End Function

Public Sub CreateInvoiceJsonFile(Invoice As InvoiceEntity)
    On Error GoTo HandleErrors
    Dim fs As New FileSystemObject
    Dim Stream As TextStream
    Dim InvoiceFileName As String
    Dim InvoicePath As String

    InvoiceFileName = CreateFileName(Invoice.DocType, Invoice.DocSerie, Invoice.DocNumber) & ".json"
    InvoicePath = fs.BuildPath(Prop.Sfs.DATAPath, InvoiceFileName)
    Set Stream = fs.CreateTextFile(InvoicePath)
    Stream.Write InvoiceToJson(Invoice)
    Stream.Close
    DebugLog "El archivo JSON " & InvoiceFileName & " fue creado correctamente.", "CreateInvoiceJsonFile"
    Exit Sub
HandleErrors:
    ErrorLog "Error al crear el archivo JSON " & InvoiceFileName, "CreateInvoiceJsonFile", Err.Number
End Sub

Public Function InvoiceToJson(Invoice As InvoiceEntity, Optional Pretty As Boolean = True) As String
    Dim Data As New Dictionary
    Dim Cabecera As New Dictionary
    Dim Detalle As New Collection
    Dim DetalleItem As Dictionary
    Dim Tributos As New Collection
    Dim Leyendas As New Collection
    Dim Leyenda As New Dictionary
    Dim Igv As New Dictionary
    Dim Item As Variant

    Cabecera.Add "tipOperacion", "0101"
    Cabecera.Add "fecEmision", Format(Invoice.EmissionDate, "yyyy-mm-dd")
    Cabecera.Add "horEmision", Format(Invoice.EmissionTime, "HH:mm:ss")
    Cabecera.Add "fecVencimiento", IIf(CInt(Invoice.DueDate) = 0, "-", Format(Invoice.DueDate, "yyyy-mm-dd"))
    Cabecera.Add "codLocalEmisor", Prop.Company.LocalCodeEmission
    Cabecera.Add "tipDocUsuario", Invoice.Customer.DocType
    Cabecera.Add "numDocUsuario", IIf(Invoice.Customer.DocNumber = Empty, "00000000", Invoice.Customer.DocNumber)
    Cabecera.Add "rznSocialUsuario", IIf(Invoice.Customer.Name = Empty, "varios", Invoice.Customer.Name)
    Cabecera.Add "tipMoneda", Invoice.TypeCurrency
    Cabecera.Add "sumTotTributos", Format(Invoice.Igv, "0.00")
    Cabecera.Add "sumTotValVenta", Format(Invoice.SubTotal, "0.00")
    Cabecera.Add "sumPrecioVenta", Format(Invoice.Total, "0.00")
    Cabecera.Add "sumDescTotal", "0.00"
    Cabecera.Add "sumOtrosCargos", "0.00"
    Cabecera.Add "sumTotalAnticipos", "0.00"
    Cabecera.Add "sumImpVenta", Format(Invoice.Total, "0.00")
    Cabecera.Add "ublVersionId", "2.1"
    Cabecera.Add "customizationId", "2.0"

    For Each Item In Invoice.Items
        Set DetalleItem = New Dictionary
        DetalleItem.Add "codUnidadMedida", Item.UnitMeasure
        DetalleItem.Add "ctdUnidadItem", Format(Item.Quantity, "0.00")
        DetalleItem.Add "codProducto", Item.ProductCode
        DetalleItem.Add "codProductoSUNAT", "-" ' catálogo 25
        DetalleItem.Add "desItem", Item.Description
        DetalleItem.Add "mtoValorUnitario", Format(Item.UnitValue, "0.00000000")
        DetalleItem.Add "sumTotTributosItem", Format(Item.Igv, "0.00") ' IGV + ISC ____
        DetalleItem.Add "codTriIGV", "1000" ' catálogo 5
        DetalleItem.Add "mtoIgvItem", Format(Item.Igv, "0.00")
        DetalleItem.Add "mtoBaseIgvItem", Format(Item.SaleValue, "0.00")
        DetalleItem.Add "nomTributoIgvItem", "IGV" ' catálogo 5 name
        DetalleItem.Add "codTipTributoIgvItem", "VAT" ' catálogo 5
        DetalleItem.Add "tipAfeIGV", "10" ' catálogo 7
        DetalleItem.Add "porIgvItem", "18.00" ' tasa IGV ___
        DetalleItem.Add "mtoPrecioVentaUnitario", Format(Item.UnitValue + (Item.UnitValue * Prop.Rate.Igv), "0.00") ' mtoValorUnitario + mtoIgvUnitario
        DetalleItem.Add "mtoValorVentaItem", Format(Item.Quantity * Item.UnitValue, "0.00") ' ctdUnidadItem * mtoValorUnitario
        Detalle.Add DetalleItem
    Next Item

    Igv.Add "ideTributo", "1000" ' catálogo 5
    Igv.Add "nomTributo", "IGV"
    Igv.Add "codTipTributo", "VAT"
    Igv.Add "mtoBaseImponible", Format(Invoice.SubTotal, "0.00")
    Igv.Add "mtoTributo", Format(Invoice.Igv, "0.00")
    Tributos.Add Igv
    
    Leyenda.Add "codLeyenda", "1000"
    Leyenda.Add "desLeyenda", AmountInLetters(Invoice.Total, Invoice.TypeCurrency)
    Leyendas.Add Leyenda

    Data.Add "cabecera", Cabecera
    Data.Add "detalle", Detalle
    Data.Add "tributos", Tributos
    Data.Add "leyendas", Leyendas

    If Pretty Then
        InvoiceToJson = ConvertToJson(Data, 2)
    Else
        InvoiceToJson = ConvertToJson(Data)
    End If
End Function

Public Sub SendElectronicDocument(DocType As String, DocNumber As String)
    Dim Body As New Dictionary
    Dim EndPoint As String

    Body.Add "num_ruc", Prop.Company.Ruc
    Body.Add "tip_docu", DocType
    Body.Add "num_docu", DocNumber
    EndPoint = "http://localhost:" & Prop.Sfs.Port & "/api/enviarXML.htm"
    
    If Post(EndPoint, ConvertToJson(Body)) Then
        DebugLog "El comprobante electrónico " & DocNumber & " se envió a la SUNAT correctamente.", "SendElectronicDocument"
    Else
        ErrorLog "Error al enviar el documento electrónico " & DocNumber & " a la SUNAT.", "SendElectronicDocument"
    End If
End Sub

Public Sub GenerateElectronicDocument(DocType As String, DocNumber As String)
    Dim Body As New Dictionary
    Dim EndPoint As String
    
    Body.Add "num_ruc", Prop.Company.Ruc
    Body.Add "tip_docu", DocType
    Body.Add "num_docu", DocNumber
    EndPoint = "http://localhost:" & Prop.Sfs.Port & "/api/GenerarComprobante.htm"
    
    If Post(EndPoint, ConvertToJson(Body)) Then
        DebugLog "El comprobante electrónico " & DocNumber & " se generó correctamente.", "GenerateElectronicDocument"
    Else
        ErrorLog "Error al generar el documento electrónico " & DocNumber, "GenerateElectronicDocument"
    End If
End Sub
