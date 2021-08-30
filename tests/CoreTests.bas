Attribute VB_Name = "CoreTests"
Option Explicit

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
    Dim Invoice As New DocumentEntity
    Dim Item As New ItemEntity
    
    Item.ProductCode = "CD0001"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    
    Invoice.OperationCode = "0101"
    Invoice.Emission = CDate("30/06/2021")
    Invoice.EmissionTime = TimeValue("10:20:14")
    Invoice.TypeCurrency = "PEN"
    
    Invoice.AddItem Item
    
    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-06-30"",""horEmision"":""10:20:14"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""rznSocialUsuario"":""VARIOS"",""tipMoneda"":""PEN"",""sumTotTributos"":""18.00"",""sumTotValVenta"":""100.00"",""sumPrecioVenta"":""118.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""118.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""CD0001"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""100.00"",""mtoTributo"":""18.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO DIECIOCHO CON 00/100 SOLES""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"
    
    With Test.It("BV con un item")
        .AssertEquals Expected, DocumentToJson(Invoice, False)
    End With
End Sub

Private Sub generar_archivo_json_de_cuando_el_cp_esta_en_dolares()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim Invoice As New DocumentEntity
    Dim Item As New ItemEntity
    
    Item.ProductCode = "CD0001"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    
    Invoice.OperationCode = "0101"
    Invoice.Emission = CDate("30/06/2021")
    Invoice.EmissionTime = TimeValue("10:20:14")
    Invoice.TypeCurrency = "USD"
    
    Invoice.AddItem Item
    
    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-06-30"",""horEmision"":""10:20:14"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""rznSocialUsuario"":""VARIOS"",""tipMoneda"":""USD"",""sumTotTributos"":""18.00"",""sumTotValVenta"":""100.00"",""sumPrecioVenta"":""118.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""118.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""CD0001"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""100.00"",""mtoTributo"":""18.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO DIECIOCHO CON 00/100 D\u00D3LARES AMERICANOS""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"
    
    With Test.It("BV con un item")
        .AssertEquals Expected, DocumentToJson(Invoice, False)
    End With
End Sub

Private Sub generar_archivo_json_de_boleta_de_venta_con_dos_items()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim Invoice As New DocumentEntity
    Dim Item As ItemEntity
    
    Invoice.OperationCode = "0101"
    Invoice.Emission = CDate("30/06/2021")
    Invoice.EmissionTime = TimeValue("10:20:14")
    Invoice.TypeCurrency = "PEN"
    
    Set Item = New ItemEntity
    Item.ProductCode = "CD0001"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    Invoice.AddItem Item
    
    Set Item = New ItemEntity
    Item.ProductCode = "CD0002"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 2"
    Item.Quantity = 5
    Item.UnitValue = 10
    Item.IgvRate = 0.18
    Invoice.AddItem Item
    
    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-06-30"",""horEmision"":""10:20:14"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""rznSocialUsuario"":""VARIOS"",""tipMoneda"":""PEN"",""sumTotTributos"":""27.00"",""sumTotValVenta"":""150.00"",""sumPrecioVenta"":""177.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""177.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""CD0001"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""},{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""5.00"",""codProducto"":""CD0002"",""codProductoSUNAT"":""-"",""desItem"":""Producto 2"",""mtoValorUnitario"":""10.00000000"",""sumTotTributosItem"":""9.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""9.00"",""mtoBaseIgvItem"":""50.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""11.80"",""mtoValorVentaItem"":""50.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""150.00"",""mtoTributo"":""27.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO SETENTA Y SIETE CON 00/100 SOLES""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"
    
    With Test.It("BV con dos items")
        .AssertEquals Expected, DocumentToJson(Invoice, False)
    End With
End Sub

Private Sub NotaTest()
    Dim Note As New DocumentEntity
    Dim Item As ItemEntity
    Dim Customer As New CustomerEntity
    Dim NoteInfo As New NoteInfoEntity
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    
    Note.OperationCode = "0101"
    Note.Emission = CDate("03/08/2021")
    Note.EmissionTime = TimeValue("13:01:30")
    Note.TypeCurrency = "PEN"
    
    Customer.DocType = "6"
    Customer.DocNumber = "20448177484"
    Customer.Name = "TEST SAC"
    Set Note.Customer = Customer
    
    NoteInfo.MotiveCode = "07"
    NoteInfo.Motive = "DEVOLUCION DE DOS PRODUCTOS"
    NoteInfo.RefDocType = "01"
    NoteInfo.RefDocSerie = "F001"
    NoteInfo.RefDocNumber = 1
    Set Note.NoteInfo = NoteInfo
    
    Set Item = New ItemEntity
    Item.ProductCode = "CD0001"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    Note.AddItem Item
    
    Set Item = New ItemEntity
    Item.ProductCode = "CD0002"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 2"
    Item.Quantity = 5
    Item.UnitValue = 10
    Item.IgvRate = 0.18
    Note.AddItem Item
    
    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-08-03"",""horEmision"":""13:01:30"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""6"",""numDocUsuario"":""20448177484"",""rznSocialUsuario"":""TEST SAC"",""tipMoneda"":""PEN"",""codMotivo"":""07"",""desMotivo"":""DEVOLUCION DE DOS PRODUCTOS"",""tipDocAfectado"":""01"",""numDocAfectado"":""F001-00000001"",""sumTotTributos"":""27.00"",""sumTotValVenta"":""150.00"",""sumPrecioVenta"":""177.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""177.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""CD0001"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""},{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""5.00"",""codProducto"":""CD0002"",""codProductoSUNAT"":""-"",""desItem"":""Producto 2"",""mtoValorUnitario"":""10.00000000"",""sumTotTributosItem"":""9.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""9.00"",""mtoBaseIgvItem"":""50.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""11.80"",""mtoValorVentaItem"":""50.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""150.00"",""mtoTributo"":""27.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO SETENTA Y SIETE CON 00/100 SOLES""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"
    
    With Test.It("Nota de crédito con dos productos devueltos")
        .AssertEquals Expected, DocumentToJson(Note, False)
    End With
End Sub

Private Sub BoletaTest()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Leyendas As String
    Dim Invoice As New DocumentEntity
    Dim Item As ItemEntity
    
    Invoice.OperationCode = "0101"
    Invoice.Emission = CDate("30/06/2021")
    Invoice.EmissionTime = TimeValue("10:20:14")
    Invoice.TypeCurrency = "PEN"

    Set Item = New ItemEntity
    Item.ProductCode = "CD0001"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 1"
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    Invoice.AddItem Item

    Set Item = New ItemEntity
    Item.ProductCode = "CD0002"
    Item.UnitMeasure = "NIU"
    Item.Description = "Producto 2"
    Item.Quantity = 5
    Item.UnitValue = 10
    Item.IgvRate = 0.18
    Invoice.AddItem Item

    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-06-30"",""horEmision"":""10:20:14"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""rznSocialUsuario"":""VARIOS"",""tipMoneda"":""PEN"",""sumTotTributos"":""27.00"",""sumTotValVenta"":""150.00"",""sumPrecioVenta"":""177.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""177.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""CD0001"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00000000"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""},{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""5.00"",""codProducto"":""CD0002"",""codProductoSUNAT"":""-"",""desItem"":""Producto 2"",""mtoValorUnitario"":""10.00000000"",""sumTotTributosItem"":""9.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""9.00"",""mtoBaseIgvItem"":""50.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""11.80"",""mtoValorVentaItem"":""50.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""150.00"",""mtoTributo"":""27.00""}]"
    Leyendas = """leyendas"":[{""codLeyenda"":""1000"",""desLeyenda"":""CIENTO SETENTA Y SIETE CON 00/100 SOLES""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "," & Leyendas & "}"

    With Test.It("BV con dos items")
        .AssertEquals Expected, DocumentToJson(Invoice, False)
    End With
End Sub

Private Sub ResumenDiarioBoletas_UnaBoleta()
    Dim Expected As String
    Dim Documents As New Collection
    Dim Document As New DocumentEntity
    Dim Item As ItemEntity
    
    Document.Emission = CDate("25/07/2021")
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
    
    Expected = "{""resumenDiario"":[{""fecEmision"":""2021-07-25"",""fecResumen"":""2021-07-26"",""tipDocResumen"":""03"",""idDocResumen"":""B001-00000001"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""tipMoneda"":""PEN"",""totValGrabado"":""100.00"",""totValExoneado"":""0.00"",""totValInafecto"":""0.00"",""totValExportado"":""0.00"",""totValGratuito"":""0.00"",""totOtroCargo"":""0.00"",""totImpCpe"":""118.00"",""tipEstado"":""1"",""tributosDocResumen"":[{""idLineaRd"":""1"",""ideTributoRd"":""1000"",""nomTributoRd"":""IGV"",""codTipTributoRd"":""VAT"",""mtoBaseImponibleRd"":""100.00"",""mtoTributoRd"":""18.00""}]}]}"
    
    With Test.It("Con estado 1 (adicionar)")
        .AssertEquals Expected, DailySummaryToJson(Documents, CDate("26/07/2021"), False)
    End With
End Sub

Private Sub ResumenDiarioBoletas_DosBoletas()
    Dim Expected As String
    Dim Documents As New Collection
    Dim Document1 As New DocumentEntity
    Dim Document2 As New DocumentEntity
    Dim Item As ItemEntity
    
    Document1.Emission = CDate("25/07/2021")
    Document1.DocType = "03"
    Document1.DocSerie = "B001"
    Document1.DocNumber = 1
    Document1.TypeCurrency = "PEN"
    
    Set Item = New ItemEntity
    Item.Quantity = 1
    Item.UnitValue = 100
    Item.IgvRate = 0.18
    Document1.AddItem Item
    
    Documents.Add Document1
    
    Document2.Emission = CDate("25/07/2021")
    Document2.DocType = "03"
    Document2.DocSerie = "B001"
    Document2.DocNumber = 2
    Document2.TypeCurrency = "PEN"
    
    Set Item = New ItemEntity
    Item.Quantity = 1
    Item.UnitValue = 200
    Item.IgvRate = 0.18
    Document2.AddItem Item
    
    Documents.Add Document2
    
    Expected = "{""resumenDiario"":[{""fecEmision"":""2021-07-25"",""fecResumen"":""2021-07-26"",""tipDocResumen"":""03"",""idDocResumen"":""B001-00000001"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""tipMoneda"":""PEN"",""totValGrabado"":""100.00"",""totValExoneado"":""0.00"",""totValInafecto"":""0.00"",""totValExportado"":""0.00"",""totValGratuito"":""0.00"",""totOtroCargo"":""0.00"",""totImpCpe"":""118.00"",""tipEstado"":""1"",""tributosDocResumen"":[{""idLineaRd"":""1"",""ideTributoRd"":""1000"",""nomTributoRd"":""IGV"",""codTipTributoRd"":""VAT"",""mtoBaseImponibleRd"":""100.00"",""mtoTributoRd"":""18.00""}]},"
    Expected = Expected & "{""fecEmision"":""2021-07-25"",""fecResumen"":""2021-07-26"",""tipDocResumen"":""03"",""idDocResumen"":""B001-00000002"",""tipDocUsuario"":""1"",""numDocUsuario"":""00000000"",""tipMoneda"":""PEN"",""totValGrabado"":""200.00"",""totValExoneado"":""0.00"",""totValInafecto"":""0.00"",""totValExportado"":""0.00"",""totValGratuito"":""0.00"",""totOtroCargo"":""0.00"",""totImpCpe"":""236.00"",""tipEstado"":""1"",""tributosDocResumen"":[{""idLineaRd"":""2"",""ideTributoRd"":""1000"",""nomTributoRd"":""IGV"",""codTipTributoRd"":""VAT"",""mtoBaseImponibleRd"":""200.00"",""mtoTributoRd"":""36.00""}]}]}"
    
    With Test.It("Con estado 1 (adicionar)")
        .AssertEquals Expected, DailySummaryToJson(Documents, CDate("26/07/2021"), False)
    End With
End Sub

Public Function Test() As VBAUnit
    Dim UnitTest As New VBAUnit
    Set Test = UnitTest
End Function
