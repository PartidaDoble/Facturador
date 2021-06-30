Attribute VB_Name = "UseCasesTests"
Option Explicit

Sub generar_archivo_json_de_boleta_de_venta()
    Dim Expected As String
    Dim Cabecera As String
    Dim Detalle As String
    Dim Tributos As String
    Dim Invoice As New InvoiceEntity
    
    Cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-06-30"",""horEmision"":""10:20:14"",""fecVencimiento"":""-"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""0"",""numDocUsuario"":""00000000"",""rznSocialUsuario"":""varios"",""tipMoneda"":""PEN"",""sumTotTributos"":""18.00"",""sumTotValVenta"":""100.00"",""sumPrecioVenta"":""118.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""118.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    Detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""CD0001"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.00"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""}]"
    Tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""100.00"",""mtoTributo"":""18.00""}]"
    Expected = "{" & Cabecera & "," & Detalle & "," & Tributos & "}"
    
    With Test.It("con un item (un producto vendido)")
        .AssertEquals Expected, InvoiceToJson(Invoice)
    End With
End Sub
