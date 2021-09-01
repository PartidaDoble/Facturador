Attribute VB_Name = "Core"
Option Explicit

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

Public Sub SendElectronicDocument(DocType As String, DocNumber As String)
    Dim Body As New Dictionary
    Dim EndPoint As String

    Body.Add "num_ruc", Prop.Company.Ruc
    Body.Add "tip_docu", DocType
    Body.Add "num_docu", DocNumber
    EndPoint = "http://localhost:" & Prop.Sfs.Port & "/api/enviarXML.htm"
    
    If Not Post(EndPoint, ConvertToJson(Body)) Then
        ErrorLog "Error al enviar el documento electrónico " & DocNumber & " a la SUNAT.", "SendElectronicDocument"
    End If
End Sub

Public Sub CreateInvoiceJsonFile(Invoice As DocumentEntity)
    On Error GoTo HandleErrors
    Dim JsonFileName As String

    JsonFileName = Prop.Company.Ruc & "-" & Invoice.Id & ".json"
    WriteFile PathJoin(Prop.Sfs.DATAPath, JsonFileName), DocumentToJson(Invoice)
    DebugLog "El archivo JSON " & JsonFileName & " fue creado correctamente.", "CreateInvoiceJsonFile"
    Exit Sub
HandleErrors:
    ErrorLog "Error al crear el archivo JSON " & JsonFileName, "CreateInvoiceJsonFile", Err.Number
End Sub

Public Sub CreateNoteJsonFile(Note As DocumentEntity)
    On Error GoTo HandleErrors
    Dim JsonFileName As String

    JsonFileName = Prop.Company.Ruc & "-" & Note.Id & ".json"
    WriteFile PathJoin(Prop.Sfs.DATAPath, JsonFileName), DocumentToJson(Note)
    DebugLog "El archivo JSON " & JsonFileName & " fue creado correctamente.", "CreateNoteJsonFile"
    Exit Sub
HandleErrors:
    ErrorLog "Error al crear el archivo JSON " & JsonFileName, "CreateNoteJsonFile", Err.Number
End Sub

Public Sub CreateDailySummaryJsonFile(JsonFileName As String, Documents As Collection, SummaryDate As Date)
    On Error GoTo HandleErrors
    WriteFile PathJoin(Prop.Sfs.DATAPath, JsonFileName), DailySummaryToJson(Documents, SummaryDate)
    DebugLog "El archivo JSON " & JsonFileName & " de resumen diario de boletas fue creado correctamente.", "CreateDailySummaryJsonFile"
    Exit Sub
HandleErrors:
    ErrorLog "Error al crear el archivo JSON " & JsonFileName & " de resumen diario de boletas.", "CreateDailySummaryJsonFile", Err.Number
End Sub

Public Sub CreateCanceledDocumentJsonFile(JsonFileName As String, Documents As Collection, SummaryDate As Date)
    On Error GoTo HandleErrors
    WriteFile PathJoin(Prop.Sfs.DATAPath, JsonFileName), CanceledDocumentsToJson(Documents, SummaryDate)
    DebugLog "El archivo JSON " & JsonFileName & " para anular comprobantes fue creado correctamente.", "CreateCanceledDocumentJsonFile"
    Exit Sub
HandleErrors:
    ErrorLog "Error al crear el archivo JSON " & JsonFileName & " para anular comprobantes.", "CreateCanceledDocumentJsonFile", Err.Number
End Sub

Public Function DocumentToJson(Document As DocumentEntity, Optional Pretty As Boolean = True) As String
    Dim Item As ItemEntity
    Dim Data As New Dictionary
    Dim Cabecera As New Dictionary
    Dim AdicionalCabecera As New Dictionary
    Dim Detalle As New Collection
    Dim DetalleItem As Dictionary
    Dim Tributos As New Collection
    Dim Leyendas As New Collection
    Dim Leyenda As Dictionary
    Dim Igv As New Dictionary
    
    Cabecera.Add "tipOperacion", Document.OperationCode
    Cabecera.Add "fecEmision", Format(Document.Emission, "yyyy-mm-dd")
    Cabecera.Add "horEmision", Format(Document.EmissionTime, "HH:mm:ss")
    Cabecera.Add "fecVencimiento", "-"
    Cabecera.Add "codLocalEmisor", Prop.Company.LocalCodeEmission
    
    If Document.Customer Is Nothing Then
        Cabecera.Add "tipDocUsuario", "1"
        Cabecera.Add "numDocUsuario", "00000000"
        Cabecera.Add "rznSocialUsuario", "VARIOS"
    Else
        Cabecera.Add "tipDocUsuario", Document.Customer.DocType
        Cabecera.Add "numDocUsuario", Document.Customer.DocNumber
        Cabecera.Add "rznSocialUsuario", Document.Customer.Name
    End If
    
    Cabecera.Add "tipMoneda", Document.TypeCurrency
    
    If Not Document.NoteInfo Is Nothing Then
        Cabecera.Add "codMotivo", Document.NoteInfo.MotiveCode
        Cabecera.Add "desMotivo", Document.NoteInfo.Motive
        Cabecera.Add "tipDocAfectado", Document.NoteInfo.RefDocType
        Cabecera.Add "numDocAfectado", Document.NoteInfo.RefDocSerie & "-" & Format(Document.NoteInfo.RefDocNumber, "00000000")
    End If
    
    Cabecera.Add "sumTotTributos", Format(Document.Igv, "0.00")
    Cabecera.Add "sumTotValVenta", Format(Document.SubTotal, "0.00")
    Cabecera.Add "sumPrecioVenta", Format(Document.Total, "0.00")
    Cabecera.Add "sumDescTotal", "0.00"
    Cabecera.Add "sumOtrosCargos", "0.00"
    Cabecera.Add "sumTotalAnticipos", "0.00"
    Cabecera.Add "sumImpVenta", Format(Document.Total, "0.00")
    Cabecera.Add "ublVersionId", "2.1"
    Cabecera.Add "customizationId", "2.0"
    
    If Not Document.Detraction Is Nothing Then
        AdicionalCabecera.Add "ctaBancoNacionDetraccion", Prop.Company.NroCtaDetraction
        AdicionalCabecera.Add "codBienDetraccion", Document.Detraction.Code
        AdicionalCabecera.Add "porDetraccion", CStr(Document.Detraction.Percentage)
        AdicionalCabecera.Add "mtoDetraccion", Format(Document.Detraction.Amount, "0.00")
        AdicionalCabecera.Add "codMedioPago", Document.Detraction.PaymentMethod
        
        Set Leyenda = New Dictionary
        Leyenda.Add "codLeyenda", "2006"
        Leyenda.Add "desLeyenda", "OPERACION SUJETA AL SPOT"
        Leyendas.Add Leyenda
    End If
    
    If Not Document.Customer Is Nothing Then
        If Document.Customer.HasValidAddress Then
            AdicionalCabecera.Add "codPaisCliente", "PE"
            AdicionalCabecera.Add "codUbigeoCliente", Document.Customer.Ubigeo
            AdicionalCabecera.Add "desDireccionCliente", Document.Customer.Address
        End If
    End If
    
    If AdicionalCabecera.Count > 0 Then
        Cabecera.Add "adicionalCabecera", AdicionalCabecera
    End If
    
    For Each Item In Document.Items
        Set DetalleItem = New Dictionary
        DetalleItem.Add "codUnidadMedida", Item.UnitMeasure
        DetalleItem.Add "ctdUnidadItem", Format(Item.Quantity, "0.00")
        DetalleItem.Add "codProducto", Item.ProductCode
        DetalleItem.Add "codProductoSUNAT", "-"
        DetalleItem.Add "desItem", Item.Description
        DetalleItem.Add "mtoValorUnitario", Format(Item.UnitValue, "0.00000000")
        DetalleItem.Add "sumTotTributosItem", Format(Item.Igv, "0.00")
        DetalleItem.Add "codTriIGV", "1000"
        DetalleItem.Add "mtoIgvItem", Format(Item.Igv, "0.00")
        DetalleItem.Add "mtoBaseIgvItem", Format(Item.SaleValue, "0.00")
        DetalleItem.Add "nomTributoIgvItem", "IGV"
        DetalleItem.Add "codTipTributoIgvItem", "VAT"
        DetalleItem.Add "tipAfeIGV", "10"
        DetalleItem.Add "porIgvItem", Format(Prop.Rate.Igv * 100, "0.00")
        DetalleItem.Add "mtoPrecioVentaUnitario", Format(Item.UnitPrice, "0.00")
        DetalleItem.Add "mtoValorVentaItem", Format(Item.SaleValue, "0.00")
        Detalle.Add DetalleItem
    Next Item
    
    Igv.Add "ideTributo", "1000"
    Igv.Add "nomTributo", "IGV"
    Igv.Add "codTipTributo", "VAT"
    Igv.Add "mtoBaseImponible", Format(Document.SubTotal, "0.00")
    Igv.Add "mtoTributo", Format(Document.Igv, "0.00")
    Tributos.Add Igv
    
    Set Leyenda = New Dictionary
    Leyenda.Add "codLeyenda", "1000"
    Leyenda.Add "desLeyenda", AmountInLetters(Document.Total, Document.TypeCurrency)
    Leyendas.Add Leyenda
    
    Data.Add "cabecera", Cabecera
    Data.Add "detalle", Detalle
    Data.Add "tributos", Tributos
    Data.Add "leyendas", Leyendas
    
    If Pretty Then
        DocumentToJson = ConvertToJson(Data, 2)
    Else
        DocumentToJson = ConvertToJson(Data)
    End If
End Function

Public Function DailySummaryToJson(Documents As Collection, SummaryDate As Date, Optional Pretty As Boolean = True) As String
    Dim Data As New Dictionary
    Dim ResumenDiario As New Collection
    Dim SummaryDocument As Dictionary
    Dim TributosDocResumen As Collection
    Dim Tributos As Dictionary
    Dim Document As DocumentEntity
    Dim DocumentCounter As Integer
    Dim State As String
    
    For Each Document In Documents
        Set SummaryDocument = New Dictionary
        SummaryDocument.Add "fecEmision", Format(Document.Emission, "yyyy-mm-dd")
        SummaryDocument.Add "fecResumen", Format(SummaryDate, "yyyy-mm-dd")
        SummaryDocument.Add "tipDocResumen", Document.DocType
        SummaryDocument.Add "idDocResumen", Document.DocSerie & "-" & Format(Document.DocNumber, "00000000")
        
        If Document.Customer Is Nothing Then
            SummaryDocument.Add "tipDocUsuario", "1"
            SummaryDocument.Add "numDocUsuario", "00000000"
        Else
            SummaryDocument.Add "tipDocUsuario", Document.Customer.DocType
            SummaryDocument.Add "numDocUsuario", Document.Customer.DocNumber
        End If
        
        SummaryDocument.Add "tipMoneda", Document.TypeCurrency
        SummaryDocument.Add "totValGrabado", Format(Document.SubTotal, "0.00")
        SummaryDocument.Add "totValExoneado", "0.00"
        SummaryDocument.Add "totValInafecto", "0.00"
        SummaryDocument.Add "totValExportado", "0.00"
        SummaryDocument.Add "totValGratuito", "0.00"
        SummaryDocument.Add "totOtroCargo", "0.00"
        SummaryDocument.Add "totImpCpe", Format(Document.Total, "0.00")
        
        If Not Document.NoteInfo Is Nothing Then
            If Document.IsBoletaNote Then
                SummaryDocument.Add "tipDocModifico", Document.NoteInfo.RefDocType
                SummaryDocument.Add "serDocModifico", Document.NoteInfo.RefDocSerie
                SummaryDocument.Add "numDocModifico", Format(Document.NoteInfo.RefDocNumber, "00000000")
            End If
        End If
        
        SummaryDocument.Add "tipEstado", Document.GetState
        
        DocumentCounter = DocumentCounter + 1
        Set Tributos = New Dictionary
        Tributos.Add "idLineaRd", CStr(DocumentCounter)
        Tributos.Add "ideTributoRd", "1000"
        Tributos.Add "nomTributoRd", "IGV"
        Tributos.Add "codTipTributoRd", "VAT"
        Tributos.Add "mtoBaseImponibleRd", Format(Document.SubTotal, "0.00")
        Tributos.Add "mtoTributoRd", Format(Document.Igv, "0.00")
        
        Set TributosDocResumen = New Collection
        TributosDocResumen.Add Tributos
        
        SummaryDocument.Add "tributosDocResumen", TributosDocResumen
        
        ResumenDiario.Add SummaryDocument
    Next Document
    
    Data.Add "resumenDiario", ResumenDiario
    
    If Pretty Then
        DailySummaryToJson = ConvertToJson(Data, 2)
    Else
        DailySummaryToJson = ConvertToJson(Data)
    End If
End Function

Public Function CanceledDocumentsToJson(Documents As Collection, SummaryDate As Date, Optional Pretty As Boolean = True) As String
    Dim Document As DocumentEntity
    Dim Data As New Dictionary
    Dim CanceledDocument As Dictionary
    Dim CanceledDocuments As New Collection
    
    For Each Document In Documents
        Set CanceledDocument = New Dictionary
        
        CanceledDocument.Add "fecGeneracion", Format(Document.Emission, "yyyy-mm-dd")
        CanceledDocument.Add "fecComunicacion", Format(SummaryDate, "yyyy-mm-dd")
        CanceledDocument.Add "tipDocBaja", Document.DocType
        CanceledDocument.Add "numDocBaja", Document.DocSerie & "-" & Format(Document.DocNumber, "00000000")
        CanceledDocument.Add "desMotivoBaja", Split(Document.CancelInfo, "|")(1)
        
        CanceledDocuments.Add CanceledDocument
    Next Document
    
    Data.Add "resumenBajas", CanceledDocuments
    
    If Pretty Then
        CanceledDocumentsToJson = ConvertToJson(Data, 2)
    Else
        CanceledDocumentsToJson = ConvertToJson(Data)
    End If
End Function
