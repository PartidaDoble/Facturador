Attribute VB_Name = "CoreTests"
Option Explicit

Private Sub RunAllModuleTests()
    DocumentToJson_Document_ConUnItem
    DocumentToJson_Document_SinCliente
    DocumentToJson_Document_ConDosItems
    DocumentToJson_Document_ConDireccion
    DocumentToJson_Document_ConDetraccion
    DocumentToJson_NotaDeCredito
    DailySummaryToJson_UnaBoleta
    DailySummaryToJson_NotaYBoletaAnulada
    CanceledDocumentsToJson_Factura
    CanceledDocumentsToJson_FacturaYNota
    DocumentToJson_Invoice_PagoContado
    DocumentToJson_Invoice_PagoCredito
End Sub

Private Sub DocumentToJson_Document_ConUnItem()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim Document As New DocumentEntity
    Dim Customer As New CustomerEntity
    Dim Item As New ItemEntity
    
    Document.OperationCode = "0101"
    Document.Emission = CDate("31/08/2021")
    Document.EmissionTime = TimeValue("10:20:14")
    Document.TypeCurrency = "PEN"
    
    Customer.DocType = "6"
    Customer.DocNumber = "20131380951"
    Customer.Name = "CLIENTE SAC"
    Set Document.Customer = Customer
    
    Item.ProductCode = "10000"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    
    Document.AddItem Item
    
    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-08-31"",""horEmision"":""10:20:14"",""fecVencimiento"":""-"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""6"",""numDocUsuario"":""20131380951"",""rznSocialUsuario"":""CLIENTE SAC"",""tipMoneda"":""PEN"",""sumTotTributos"":""18.00"",""sumTotValVenta"":""100.00"",""sumPrecioVenta"":""118.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""118.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""10000"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""100.00"",""mtoTributo"":""18.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO DIECIOCHO CON 00/100 SOLES""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"
    
    With Test.It("Document_ConUnItem")
        .AssertEquals Expected, DocumentToJson(Document, False)
    End With
End Sub

Private Sub DocumentToJson_Document_SinCliente()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim Document As New DocumentEntity
    Dim Item As New ItemEntity
    
    Document.OperationCode = "0101"
    Document.Emission = CDate("31/08/2021")
    Document.EmissionTime = TimeValue("10:20:14")
    Document.TypeCurrency = "PEN"
    
    Item.ProductCode = "10000"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    
    Document.AddItem Item
    
    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-08-31"",""horEmision"":""10:20:14"",""fecVencimiento"":""-"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""rznSocialUsuario"":""VARIOS"",""tipMoneda"":""PEN"",""sumTotTributos"":""18.00"",""sumTotValVenta"":""100.00"",""sumPrecioVenta"":""118.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""118.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""10000"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""100.00"",""mtoTributo"":""18.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO DIECIOCHO CON 00/100 SOLES""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"
    
    With Test.It("Document_SinCliente")
        .AssertEquals Expected, DocumentToJson(Document, False)
    End With
End Sub

Private Sub DocumentToJson_Document_ConDosItems()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim Document As New DocumentEntity
    Dim Customer As New CustomerEntity
    Dim Item As ItemEntity
    
    Document.OperationCode = "0101"
    Document.Emission = CDate("31/08/2021")
    Document.EmissionTime = TimeValue("10:20:14")
    Document.TypeCurrency = "PEN"
        
    Customer.DocType = "6"
    Customer.DocNumber = "20131380951"
    Customer.Name = "CLIENTE SAC"
    Set Document.Customer = Customer
    
    Set Item = New ItemEntity
    Item.ProductCode = "10000"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    Document.AddItem Item
    
    Set Item = New ItemEntity
    Item.ProductCode = "10001"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 2"
    Item.Quantity = 5
    Item.UnitValue = 10
    Item.IgvRate = 0.18
    Document.AddItem Item
    
    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-08-31"",""horEmision"":""10:20:14"",""fecVencimiento"":""-"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""6"",""numDocUsuario"":""20131380951"",""rznSocialUsuario"":""CLIENTE SAC"",""tipMoneda"":""PEN"",""sumTotTributos"":""27.00"",""sumTotValVenta"":""150.00"",""sumPrecioVenta"":""177.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""177.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""10000"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""},{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""5.00"",""codProducto"":""10001"",""codProductoSUNAT"":""-"",""desItem"":""Producto 2"",""mtoValorUnitario"":""10.00000000"",""sumTotTributosItem"":""9.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""9.00"",""mtoBaseIgvItem"":""50.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""11.80"",""mtoValorVentaItem"":""50.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""150.00"",""mtoTributo"":""27.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO SETENTA Y SIETE CON 00/100 SOLES""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"
    
    With Test.It("Document_ConDosItems")
        .AssertEquals Expected, DocumentToJson(Document, False)
    End With
End Sub

Private Sub DocumentToJson_Document_ConDireccion()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim Document As New DocumentEntity
    Dim Customer As New CustomerEntity
    Dim Item As New ItemEntity
    
    Document.OperationCode = "0101"
    Document.Emission = CDate("31/08/2021")
    Document.EmissionTime = TimeValue("10:20:14")
    Document.TypeCurrency = "PEN"
    
    Customer.DocType = "6"
    Customer.DocNumber = "20131380951"
    Customer.Name = "CLIENTE SAC"
    Customer.Ubigeo = "210801"
    Customer.Address = "JR. TACNA NRO. 562"
    Set Document.Customer = Customer
    
    Item.ProductCode = "10000"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    
    Document.AddItem Item
    
    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-08-31"",""horEmision"":""10:20:14"",""fecVencimiento"":""-"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""6"",""numDocUsuario"":""20131380951"",""rznSocialUsuario"":""CLIENTE SAC"",""tipMoneda"":""PEN"",""sumTotTributos"":""18.00"",""sumTotValVenta"":""100.00"",""sumPrecioVenta"":""118.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""118.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0"",""adicionalCabecera"":{""codPaisCliente"":""PE"",""codUbigeoCliente"":""210801"",""desDireccionCliente"":""JR. TACNA NRO. 562""}}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""10000"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""100.00"",""mtoTributo"":""18.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO DIECIOCHO CON 00/100 SOLES""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"
    
    With Test.It("Document_ConDireccion")
        .AssertEquals Expected, DocumentToJson(Document, False)
    End With
End Sub

Private Sub DocumentToJson_Document_ConDetraccion()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim Document As New DocumentEntity
    Dim Customer As New CustomerEntity
    Dim Detraction As New DetractionEntity
    Dim Item As New ItemEntity
    
    Document.OperationCode = "1001"
    Document.Emission = CDate("31/08/2021")
    Document.EmissionTime = TimeValue("10:20:14")
    Document.TypeCurrency = "PEN"
    
    Customer.DocType = "6"
    Customer.DocNumber = "20131380951"
    Customer.Name = "CLIENTE SAC"
    Set Document.Customer = Customer
    
    Item.ProductCode = "10000"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    
    Document.AddItem Item
    
    Detraction.Code = "020"
    Detraction.Percentage = "10"
    Detraction.Amount = "18.00"
    Detraction.PaymentMethod = "001"
    Set Document.Detraction = Detraction
    
    Cabecera = """cabecera"":{""tipOperacion"":""1001"",""fecEmision"":""2021-08-31"",""horEmision"":""10:20:14"",""fecVencimiento"":""-"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""6"",""numDocUsuario"":""20131380951"",""rznSocialUsuario"":""CLIENTE SAC"",""tipMoneda"":""PEN"",""sumTotTributos"":""18.00"",""sumTotValVenta"":""100.00"",""sumPrecioVenta"":""118.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""118.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0"",""adicionalCabecera"":{""ctaBancoNacionDetraccion"":""00-711-099445"",""codBienDetraccion"":""020"",""porDetraccion"":""10"",""mtoDetraccion"":""18.00"",""codMedioPago"":""001""}}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""10000"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""100.00"",""mtoTributo"":""18.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""2006"",""desLeyenda"":""OPERACION SUJETA AL SPOT""},{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO DIECIOCHO CON 00/100 SOLES""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"
    
    With Test.It("Document_ConUnItem")
        .AssertEquals Expected, DocumentToJson(Document, False)
    End With
End Sub

Private Sub DocumentToJson_NotaDeCredito()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim Document As New DocumentEntity
    Dim Customer As New CustomerEntity
    Dim NoteInfo As New NoteInfoEntity
    Dim Item As New ItemEntity
    
    Document.OperationCode = "0101"
    Document.Emission = CDate("31/08/2021")
    Document.EmissionTime = TimeValue("10:20:14")
    Document.TypeCurrency = "PEN"
    
    Customer.DocType = "6"
    Customer.DocNumber = "20131380951"
    Customer.Name = "CLIENTE SAC"
    Set Document.Customer = Customer
    
    NoteInfo.RefDocType = "01"
    NoteInfo.RefDocSerie = "F001"
    NoteInfo.RefDocNumber = "00000001"
    NoteInfo.MotiveCode = "01"
    NoteInfo.Motive = "ANULACION DE LA OPERACION"
    Set Document.NoteInfo = NoteInfo
    
    Item.ProductCode = "10000"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    
    Document.AddItem Item
    
    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-08-31"",""horEmision"":""10:20:14"",""fecVencimiento"":""-"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""6"",""numDocUsuario"":""20131380951"",""rznSocialUsuario"":""CLIENTE SAC"",""tipMoneda"":""PEN"",""codMotivo"":""01"",""desMotivo"":""ANULACION DE LA OPERACION"",""tipDocAfectado"":""01"",""numDocAfectado"":""F001-00000001"",""sumTotTributos"":""18.00"",""sumTotValVenta"":""100.00"",""sumPrecioVenta"":""118.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""118.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""10000"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""100.00"",""mtoTributo"":""18.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO DIECIOCHO CON 00/100 SOLES""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"
    
    With Test.It("Document_ConUnItem")
        .AssertEquals Expected, DocumentToJson(Document, False)
    End With
End Sub

Private Sub DailySummaryToJson_UnaBoleta()
    Dim Document As New DocumentEntity
    Dim Item As ItemEntity
    Dim Documents As New Collection
    Dim Expected As String
    
    Document.Emission = CDate("30/08/2021")
    Document.DocType = "03"
    Document.DocSerie = "B001"
    Document.DocNumber = 1
    Document.TypeCurrency = "PEN"
    
    Set Item = New ItemEntity
    Item.Quantity = 1
    Item.UnitValue = 100
    Item.IgvRate = 0.18
    Document.AddItem Item
    
    Documents.Add Document
    
    Expected = "{""resumenDiario"":[{""fecEmision"":""2021-08-30"",""fecResumen"":""2021-08-31"",""tipDocResumen"":""03"",""idDocResumen"":""B001-00000001"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""tipMoneda"":""PEN"",""totValGrabado"":""100.00"",""totValExoneado"":""0.00"",""totValInafecto"":""0.00"",""totValExportado"":""0.00"",""totValGratuito"":""0.00"",""totOtroCargo"":""0.00"",""totImpCpe"":""118.00"",""tipEstado"":""1"",""tributosDocResumen"":[{""idLineaRd"":""1"",""ideTributoRd"":""1000"",""nomTributoRd"":""IGV"",""codTipTributoRd"":""VAT"",""mtoBaseImponibleRd"":""100.00"",""mtoTributoRd"":""18.00""}]}]}"
    
    With Test.It("DailySummaryToJson_UnaBoleta")
        .AssertEquals Expected, DailySummaryToJson(Documents, CDate("31/08/2021"), False)
    End With
End Sub

Private Sub DailySummaryToJson_NotaYBoletaAnulada()
    Dim Document As DocumentEntity
    Dim Item As ItemEntity
    Dim NoteInfo As New NoteInfoEntity
    Dim Documents As New Collection
    Dim CreditNote As String
    Dim Boleta As String
    Dim Expected As String
    
    Set Document = New DocumentEntity
    Document.Emission = CDate("30/08/2021")
    Document.DocType = "07"
    Document.DocSerie = "BC01"
    Document.DocNumber = 1
    Document.TypeCurrency = "PEN"
    
    NoteInfo.RefDocType = "03"
    NoteInfo.RefDocSerie = "B001"
    NoteInfo.RefDocNumber = 1
    Set Document.NoteInfo = NoteInfo
    
    Set Item = New ItemEntity
    Item.Quantity = 1
    Item.UnitValue = 100
    Item.IgvRate = 0.18
    Document.AddItem Item
    
    Documents.Add Document
    
    
    Set Document = New DocumentEntity
    Document.Emission = CDate("30/08/2021")
    Document.DocType = "03"
    Document.DocSerie = "B001"
    Document.DocNumber = 1
    Document.CancelInfo = "X|foo bar"
    Document.TypeCurrency = "PEN"
    
    Set Item = New ItemEntity
    Item.Quantity = 1
    Item.UnitValue = 100
    Item.IgvRate = 0.18
    Document.AddItem Item
    
    Documents.Add Document
    
    CreditNote = "{""fecEmision"":""2021-08-30"",""fecResumen"":""2021-08-31"",""tipDocResumen"":""07"",""idDocResumen"":""BC01-00000001"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""tipMoneda"":""PEN"",""totValGrabado"":""100.00"",""totValExoneado"":""0.00"",""totValInafecto"":""0.00"",""totValExportado"":""0.00"",""totValGratuito"":""0.00"",""totOtroCargo"":""0.00"",""totImpCpe"":""118.00"",""tipDocModifico"":""03"",""serDocModifico"":""B001"",""numDocModifico"":""00000001"",""tipEstado"":""2"",""tributosDocResumen"":[{""idLineaRd"":""1"",""ideTributoRd"":""1000"",""nomTributoRd"":""IGV"",""codTipTributoRd"":""VAT"",""mtoBaseImponibleRd"":""100.00"",""mtoTributoRd"":""18.00""}]}"
    Boleta = "{""fecEmision"":""2021-08-30"",""fecResumen"":""2021-08-31"",""tipDocResumen"":""03"",""idDocResumen"":""B001-00000001"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""tipMoneda"":""PEN"",""totValGrabado"":""100.00"",""totValExoneado"":""0.00"",""totValInafecto"":""0.00"",""totValExportado"":""0.00"",""totValGratuito"":""0.00"",""totOtroCargo"":""0.00"",""totImpCpe"":""118.00"",""tipEstado"":""3"",""tributosDocResumen"":[{""idLineaRd"":""2"",""ideTributoRd"":""1000"",""nomTributoRd"":""IGV"",""codTipTributoRd"":""VAT"",""mtoBaseImponibleRd"":""100.00"",""mtoTributoRd"":""18.00""}]}"
    Expected = "{""resumenDiario"":[" & CreditNote & "," & Boleta & "]}"
    
    With Test.It("DailySummaryToJson_NotaYBoletaAnulada")
        .AssertEquals Expected, DailySummaryToJson(Documents, CDate("31/08/2021"), False)
    End With
End Sub

Sub CanceledDocumentsToJson_Factura()
    Dim Document As New DocumentEntity
    Dim Documents As New Collection
    Dim Expected As String
    
    Document.Emission = CDate("30/08/2021")
    Document.DocType = "01"
    Document.DocSerie = "F001"
    Document.DocNumber = 1
    Document.CancelInfo = "X|ERROR EN EL RUC"
    
    Documents.Add Document
    
    Expected = "{""resumenBajas"":[{""fecGeneracion"":""2021-08-30"",""fecComunicacion"":""2021-08-31"",""tipDocBaja"":""01"",""numDocBaja"":""F001-00000001"",""desMotivoBaja"":""ERROR EN EL RUC""}]}"
    
    With Test.It("CanceledDocumentsToJson_Factura")
        .AssertEquals Expected, CanceledDocumentsToJson(Documents, CDate("31/08/2021"), False)
    End With
End Sub

Sub CanceledDocumentsToJson_FacturaYNota()
    Dim Document As DocumentEntity
    Dim Documents As New Collection
    Dim Expected As String
    
    Set Document = New DocumentEntity
    Document.Emission = CDate("30/08/2021")
    Document.DocType = "01"
    Document.DocSerie = "F001"
    Document.DocNumber = 1
    Document.CancelInfo = "X|ERROR EN EL RUC"
    
    Documents.Add Document
    
    Set Document = New DocumentEntity
    Document.Emission = CDate("30/08/2021")
    Document.DocType = "07"
    Document.DocSerie = "FC01"
    Document.DocNumber = 1
    Document.CancelInfo = "X|ERROR EN EL RUC"
    
    Documents.Add Document
    
    Expected = "{""resumenBajas"":[{""fecGeneracion"":""2021-08-30"",""fecComunicacion"":""2021-08-31"",""tipDocBaja"":""01"",""numDocBaja"":""F001-00000001"",""desMotivoBaja"":""ERROR EN EL RUC""},{""fecGeneracion"":""2021-08-30"",""fecComunicacion"":""2021-08-31"",""tipDocBaja"":""07"",""numDocBaja"":""FC01-00000001"",""desMotivoBaja"":""ERROR EN EL RUC""}]}"
    
    With Test.It("CanceledDocumentsToJson_FacturaYNota")
        .AssertEquals Expected, CanceledDocumentsToJson(Documents, CDate("31/08/2021"), False)
    End With
End Sub

Private Sub DocumentToJson_Invoice_PagoContado()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim DatoPago As String
    Dim Document As New DocumentEntity
    Dim Customer As New CustomerEntity
    Dim Item As New ItemEntity
    Dim WayPay As New WayPayEntity
    
    Document.OperationCode = "0101"
    Document.Emission = CDate("31/08/2021")
    Document.EmissionTime = TimeValue("10:20:14")
    Document.TypeCurrency = "PEN"
    
    Customer.DocType = "6"
    Customer.DocNumber = "20131380951"
    Customer.Name = "CLIENTE SAC"
    Set Document.Customer = Customer
    
    Item.ProductCode = "10000"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    Document.AddItem Item
    
    WayPay.Way = "Contado"
    WayPay.NetAmountPending = "0.00"
    WayPay.TypeCurrency = "PEN"
    
    Set Document.WayPay = WayPay
    
    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-08-31"",""horEmision"":""10:20:14"",""fecVencimiento"":""-"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""6"",""numDocUsuario"":""20131380951"",""rznSocialUsuario"":""CLIENTE SAC"",""tipMoneda"":""PEN"",""sumTotTributos"":""18.00"",""sumTotValVenta"":""100.00"",""sumPrecioVenta"":""118.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""118.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""10000"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""100.00"",""mtoTributo"":""18.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO DIECIOCHO CON 00/100 SOLES""}]"
    DatoPago = """datoPago"":{""formaPago"":""Contado"",""mtoNetoPendientePago"":""0.00"",""tipMonedaMtoNetoPendientePago"":""PEN""}"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "," & DatoPago & "}"
    
    With Test.It("Document_ConUnItem")
        .AssertEquals Expected, DocumentToJson(Document, False)
    End With
End Sub

Private Sub DocumentToJson_Invoice_PagoCredito()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim DatoPago As String
    Dim DetallePago As String
    Dim Document As New DocumentEntity
    Dim Customer As New CustomerEntity
    Dim Item As New ItemEntity
    Dim WayPay As New WayPayEntity
    Dim Installment1 As New InstallmentEntity
    Dim Installment2 As New InstallmentEntity
    
    Document.OperationCode = "0101"
    Document.Emission = CDate("31/08/2021")
    Document.EmissionTime = TimeValue("10:20:14")
    Document.TypeCurrency = "PEN"
    
    Customer.DocType = "6"
    Customer.DocNumber = "20131380951"
    Customer.Name = "CLIENTE SAC"
    Set Document.Customer = Customer
    
    Item.ProductCode = "10000"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    Document.AddItem Item
    
    WayPay.Way = "Credito"
    WayPay.NetAmountPending = "118.00"
    WayPay.TypeCurrency = "PEN"
    
    Installment1.PaymentDate = CDate("10/09/2021")
    Installment1.Amount = 50
    Installment1.TypeCurrency = "PEN"
    
    Installment2.PaymentDate = CDate("20/09/2021")
    Installment2.Amount = 68
    Installment2.TypeCurrency = "PEN"
    
    WayPay.Installments.Add Installment1
    WayPay.Installments.Add Installment2
    
    Set Document.WayPay = WayPay
    
    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-08-31"",""horEmision"":""10:20:14"",""fecVencimiento"":""-"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""6"",""numDocUsuario"":""20131380951"",""rznSocialUsuario"":""CLIENTE SAC"",""tipMoneda"":""PEN"",""sumTotTributos"":""18.00"",""sumTotValVenta"":""100.00"",""sumPrecioVenta"":""118.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""118.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""10000"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""100.00"",""mtoTributo"":""18.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO DIECIOCHO CON 00/100 SOLES""}]"
    DatoPago = """datoPago"":{""formaPago"":""Credito"",""mtoNetoPendientePago"":""118.00"",""tipMonedaMtoNetoPendientePago"":""PEN""}"
    DetallePago = """detallePago"":[{""mtoCuotaPago"":""50.00"",""fecCuotaPago"":""2021-09-10"",""tipMonedaCuotaPago"":""PEN""},{""mtoCuotaPago"":""68.00"",""fecCuotaPago"":""2021-09-20"",""tipMonedaCuotaPago"":""PEN""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "," & DatoPago & "," & DetallePago & "}"
    
    With Test.It("Document_ConUnItem")
        .AssertEquals Expected, DocumentToJson(Document, False)
    End With
End Sub
