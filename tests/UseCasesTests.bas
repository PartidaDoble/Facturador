Attribute VB_Name = "UseCasesTests"
Option Explicit

Sub generar_archivo_json()
    Dim expected As String
    Dim cabecera As String
    Dim detalle As String
    Dim tributos As String
    Dim invoice As New InvoiceEntity
    
    cabecera = """cabecera"":{""tipOperacion"":""0101"",""fecEmision"":""2021-06-28"",""horEmision"":""10:20:14"",""fecVencimiento"":""-"",""codLocalEmisor"":""0000"",""tipDocUsuario"":""0"",""numDocUsuario"":""00000000"",""rznSocialUsuario"":""varios"",""tipMoneda"":""PEN"",""sumTotTributos"":""18.00"",""sumTotValVenta"":""100.00"",""sumPrecioVenta"":""118.00"",""sumDescTotal"":""0.00"",""sumOtrosCargos"":""0.00"",""sumTotalAnticipos"":""0.00"",""sumImpVenta"":""118.00"",""ublVersionId"":""2.1"",""customizationId"":""2.0""}"
    detalle = """detalle"":[{""codUnidadMedida"":""NIU"",""ctdUnidadItem"":""2.00"",""codProducto"":""CD0001"",""codProductoSUNAT"":""-"",""desItem"":""Producto 1"",""mtoValorUnitario"":""50.00"",""sumTotTributosItem"":""18.00"",""codTriIGV"":""1000"",""mtoIgvItem"":""18.00"",""mtoBaseIgvItem"":""100.00"",""nomTributoIgvItem"":""IGV"",""codTipTributoIgvItem"":""VAT"",""tipAfeIGV"":""10"",""porIgvItem"":""18.0"",""mtoPrecioVentaUnitario"":""59.00"",""mtoValorVentaItem"":""100.00""}]"
    tributos = """tributos"":[{""ideTributo"":""1000"",""nomTributo"":""IGV"",""codTipTributo"":""VAT"",""mtoBaseImponible"":""100.00"",""mtoTributo"":""18.00""}]"
    expected = "{" & cabecera & detalle & tributos & "}"
    
    
    With Test.It("de una boleta con un producto y un item")
        .AssertEquals expected, InvoiceToJson(invoice)
    End With
End Sub

Public Sub FunJson()
    Dim Data As New Dictionary
    Dim Detail As New Collection
    
    Data.Add "foo", 10
    Data.Add "bar", True
    Data.Add "baz", "qwerty"
    
    Detail.Add 10
    Detail.Add "jj"
    
    Data.Add "detalle", Detail
    
    Debug.Print ConvertToJson(Data, 2)
End Sub
