Attribute VB_Name = "CoreTests"
Option Explicit

Public Function Test() As VBAUnit
    Dim UnitTest As New VBAUnit
    Set Test = UnitTest
End Function

Private Sub RunAllModuleTests()
    generar_archivo_json_de_boleta_de_venta_con_un_item
    generar_archivo_json_de_cuando_el_cp_esta_en_dolares
    generar_archivo_json_de_boleta_de_venta_con_dos_items
End Sub

Private Sub generar_archivo_json_de_boleta_de_venta_con_un_item()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim Invoice As New InvoiceEntity
    Dim Item As New ItemEntity

    Item.ProductCode = "CD0001"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50

    Invoice.EmissionDate = DateValue("30/06/2021")
    Invoice.EmissionTime = TimeValue("10:20:14")
    Invoice.TypeCurrency = "PEN"
    Invoice.Customer.DocType = "1"
    
    Invoice.AddItem Item

    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-06-30"",""horEmision"":""10:20:14"",""fecVencimiento"":""-"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""rznSocialUsuario"":""varios"",""tipMoneda"":""PEN"",""sumTotTributos"":""18.00"",""sumTotValVenta"":""100.00"",""sumPrecioVenta"":""118.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""118.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""CD0001"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""100.00"",""mtoTributo"":""18.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO DIECIOCHO CON 00/100 SOLES""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"

    With Test.It("BV con un item")
        .AssertEquals Expected, InvoiceToJson(Invoice, False)
    End With
End Sub

Private Sub generar_archivo_json_de_cuando_el_cp_esta_en_dolares()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim Invoice As New InvoiceEntity
    Dim Item As New ItemEntity

    Item.ProductCode = "CD0001"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50

    Invoice.EmissionDate = DateValue("30/06/2021")
    Invoice.EmissionTime = TimeValue("10:20:14")
    Invoice.TypeCurrency = "USD"
    Invoice.Customer.DocType = "1"
    
    Invoice.AddItem Item

    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-06-30"",""horEmision"":""10:20:14"",""fecVencimiento"":""-"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""rznSocialUsuario"":""varios"",""tipMoneda"":""USD"",""sumTotTributos"":""18.00"",""sumTotValVenta"":""100.00"",""sumPrecioVenta"":""118.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""118.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""CD0001"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""100.00"",""mtoTributo"":""18.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO DIECIOCHO CON 00/100 D\u00D3LARES AMERICANOS""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"

    With Test.It("BV con un item")
        .AssertEquals Expected, InvoiceToJson(Invoice, False)
    End With
End Sub

Private Sub generar_archivo_json_de_boleta_de_venta_con_dos_items()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim Invoice As New InvoiceEntity
    Dim Item As ItemEntity

    Set Item = New ItemEntity
    Item.ProductCode = "CD0001"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Invoice.AddItem Item

    Set Item = New ItemEntity
    Item.ProductCode = "CD0002"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 2"
    Item.Quantity = 5
    Item.UnitValue = 10
    Invoice.AddItem Item

    Invoice.EmissionDate = DateValue("30/06/2021")
    Invoice.EmissionTime = TimeValue("10:20:14")
    Invoice.TypeCurrency = "PEN"
    Invoice.Customer.DocType = "1"

    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-06-30"",""horEmision"":""10:20:14"",""fecVencimiento"":""-"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""rznSocialUsuario"":""varios"",""tipMoneda"":""PEN"",""sumTotTributos"":""27.00"",""sumTotValVenta"":""150.00"",""sumPrecioVenta"":""177.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""177.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""CD0001"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""},{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""5.00"",""codProducto"":""CD0002"",""codProductoSUNAT"":""-"",""desItem"":""Producto 2"",""mtoValorUnitario"":""10.00000000"",""sumTotTributosItem"":""9.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""9.00"",""mtoBaseIgvItem"":""50.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""11.80"",""mtoValorVentaItem"":""50.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""150.00"",""mtoTributo"":""27.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO SETENTA Y SIETE CON 00/100 SOLES""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"

    With Test.It("BV con dos items")
        .AssertEquals Expected, InvoiceToJson(Invoice, False)
    End With
End Sub

Private Sub baja_de_dos_documentos()
    Dim Doc As New CanceledDocumentEntity
    Dim Doc1 As New CanceledDocumentEntity
    Dim Docs As New Collection
    Dim Expected As String
    
    Doc.GenerationDate = DateValue("18/07/2021")
    Doc.CommunicationDate = DateValue("18/07/2021")
    Doc.DocType = "01"
    Doc.DocNumber = "F001-00000007"
    Doc.Motivo = "ERROR EN EL NOMBRE DEL CLIENTE"
    Docs.Add Doc
    
    Doc1.GenerationDate = DateValue("18/07/2021")
    Doc1.CommunicationDate = DateValue("18/07/2021")
    Doc1.DocType = "01"
    Doc1.DocNumber = "F001-00000008"
    Doc1.Motivo = "ERROR EN EL NUMERO DE RUC"
    Docs.Add Doc1
    
    Expected = "{""resumenBajas"":[{""fecGeneracion"":""2021-07-18"",""fecComunicacion"":""2021-07-18"",""tipDocBaja"":""01"",""numDocBaja"":""F001-00000007"",""desMotivoBaja"":""ERROR EN EL NOMBRE DEL CLIENTE""},{""fecGeneracion"":""2021-07-18"",""fecComunicacion"":""2021-07-18"",""tipDocBaja"":""01"",""numDocBaja"":""F001-00000008"",""desMotivoBaja"":""ERROR EN EL NUMERO DE RUC""}]}"
    
    With Test.It("baja de dos facturas")
        .AssertEquals Expected, CanceledDocumentsToJson(Docs, False)
    End With
End Sub

Private Sub ResumenDiarioBoletas_UnaBoleta()
    Dim Expected As String
    Dim Documents As New Collection
    Dim Document As New DocumentEntity
    
    Document.Emission = DateValue("25/07/2021")
    Document.DocType = "03"
    Document.DocSerie = "B001"
    Document.DocNumber = 1
    Document.CustomerDocType = "1"
    Document.CustomerDocNumber = "00000000"
    Document.TypeCurrency = "PEN"
    Document.SubTotal = 100
    Document.Igv = 18
    Document.Total = 118

    Documents.Add Document
    
    Expected = "{""resumenDiario"":[{""fecEmision"":""2021-07-25"",""fecResumen"":""2021-07-26"",""tipDocResumen"":""03"",""idDocResumen"":""B001-00000001"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""tipMoneda"":""PEN"",""totValGrabado"":""100.00"",""totValExoneado"":""0.00"",""totValInafecto"":""0.00"",""totValExportado"":""0.00"",""totValGratuito"":""0.00"",""totOtroCargo"":""0.00"",""totImpCpe"":""118.00"",""tipEstado"":""1"",""tributosDocResumen"":[{""idLineaRd"":""1"",""ideTributoRd"":""1000"",""nomTributoRd"":""IGV"",""codTipTributoRd"":""VAT"",""mtoBaseImponibleRd"":""100.00"",""mtoTributoRd"":""18.00""}]}]}"
    
    With Test.It("Con estado 1 (adicionar)")
        .AssertEquals Expected, DailySummaryToJson(Documents, DateValue("26/07/2021"), "1", False)
    End With
End Sub

Private Sub ResumenDiarioBoletas_DosBoleta()
    Dim Expected As String
    Dim Documents As New Collection
    Dim Document1 As New DocumentEntity
    Dim Document2 As New DocumentEntity
    
    Document1.Emission = DateValue("25/07/2021")
    Document1.DocType = "03"
    Document1.DocSerie = "B001"
    Document1.DocNumber = 1
    Document1.CustomerDocType = "1"
    Document1.CustomerDocNumber = "00000000"
    Document1.TypeCurrency = "PEN"
    Document1.SubTotal = 100
    Document1.Igv = 18
    Document1.Total = 118

    Documents.Add Document1
    
    Document2.Emission = DateValue("25/07/2021")
    Document2.DocType = "03"
    Document2.DocSerie = "B001"
    Document2.DocNumber = 2
    Document2.CustomerDocType = "1"
    Document2.CustomerDocNumber = "00000000"
    Document2.TypeCurrency = "PEN"
    Document2.SubTotal = 200
    Document2.Igv = 36
    Document2.Total = 236

    Documents.Add Document2
    
    Expected = "{""resumenDiario"":[{""fecEmision"":""2021-07-25"",""fecResumen"":""2021-07-26"",""tipDocResumen"":""03"",""idDocResumen"":""B001-00000001"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""tipMoneda"":""PEN"",""totValGrabado"":""100.00"",""totValExoneado"":""0.00"",""totValInafecto"":""0.00"",""totValExportado"":""0.00"",""totValGratuito"":""0.00"",""totOtroCargo"":""0.00"",""totImpCpe"":""118.00"",""tipEstado"":""1"",""tributosDocResumen"":[{""idLineaRd"":""1"",""ideTributoRd"":""1000"",""nomTributoRd"":""IGV"",""codTipTributoRd"":""VAT"",""mtoBaseImponibleRd"":""100.00"",""mtoTributoRd"":""18.00""}]},"
    Expected = Expected & "{""fecEmision"":""2021-07-25"",""fecResumen"":""2021-07-26"",""tipDocResumen"":""03"",""idDocResumen"":""B001-00000002"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""tipMoneda"":""PEN"",""totValGrabado"":""200.00"",""totValExoneado"":""0.00"",""totValInafecto"":""0.00"",""totValExportado"":""0.00"",""totValGratuito"":""0.00"",""totOtroCargo"":""0.00"",""totImpCpe"":""236.00"",""tipEstado"":""1"",""tributosDocResumen"":[{""idLineaRd"":""2"",""ideTributoRd"":""1000"",""nomTributoRd"":""IGV"",""codTipTributoRd"":""VAT"",""mtoBaseImponibleRd"":""200.00"",""mtoTributoRd"":""36.00""}]}]}"
    
    With Test.It("Con estado 1 (adicionar)")
        .AssertEquals Expected, DailySummaryToJson(Documents, DateValue("26/07/2021"), "1", False)
    End With
End Sub
