Attribute VB_Name = "UseCases"
Option Explicit

Public Function InvoiceToJson(Invoice As InvoiceEntity) As String
    Dim Data As New Dictionary
    Dim Cabecera As New Dictionary
    Dim Detalle As New Collection
    Dim Item As New Dictionary
    Dim Tributos As New Collection
    Dim Igv As New Dictionary

    Cabecera.Add "tipOperacion", "0101"
    Cabecera.Add "fecEmision", "2021-06-30"
    Cabecera.Add "horEmision", "10:20:14"
    Cabecera.Add "fecVencimiento", "-"
    Cabecera.Add "codLocalEmisor", "0000"
    Cabecera.Add "tipDocUsuario", "0"
    Cabecera.Add "numDocUsuario", "00000000"
    Cabecera.Add "rznSocialUsuario", "varios"
    Cabecera.Add "tipMoneda", "PEN" ' USD EUR
    Cabecera.Add "sumTotTributos", Format(Invoice.Igv, "0.00") ' sumatoria Tributos
    Cabecera.Add "sumTotValVenta", Format(Invoice.SubTotal, "0.00") ' suma valor de venta de items
    Cabecera.Add "sumPrecioVenta", Format(Invoice.Total, "0.00") ' suma precio de venta de items
    Cabecera.Add "sumDescTotal", "0.00"
    Cabecera.Add "sumOtrosCargos", "0.00"
    Cabecera.Add "sumTotalAnticipos", "0.00"
    Cabecera.Add "sumImpVenta", Format(Invoice.Total, "0.00") ' importe total de la venta
    Cabecera.Add "ublVersionId", "2.1"
    Cabecera.Add "customizationId", "2.0" ' revisar

    Item.Add "codUnidadMedida", "NIU" ' cat�logo 3
    Item.Add "ctdUnidadItem", Format(Invoice.Item.Quantity, "0.00")
    Item.Add "codProducto", "CD0001"
    Item.Add "codProductoSUNAT", "-" ' cat�logo 25
    Item.Add "desItem", "Producto 1" ' descripci�n del item
    Item.Add "mtoValorUnitario", Format(Invoice.Item.UnitValue, "0.0000")
    Item.Add "sumTotTributosItem", Format(Invoice.Item.Igv, "0.00") ' IGV + ISC ____
    Item.Add "codTriIGV", "1000" ' cat�logo 5
    Item.Add "mtoIgvItem", Format(Invoice.Item.Igv, "0.00")
    Item.Add "mtoBaseIgvItem", Format(Invoice.Item.SaleValue, "0.00")
    Item.Add "nomTributoIgvItem", "IGV" ' cat�logo 5 name
    Item.Add "codTipTributoIgvItem", "VAT" ' cat�logo 5
    Item.Add "tipAfeIGV", "10" ' cat�logo 7
    Item.Add "porIgvItem", "18.00" ' tasa IGV ___
    Item.Add "mtoPrecioVentaUnitario", "59.00" ' (mtoValorVentaItem + mtoIgvItem ) / ctdUnidadItem
    Item.Add "mtoValorVentaItem", "100.00" ' mtoValorUnitario * ctdUnidadItem
    Detalle.Add Item

    Igv.Add "ideTributo", "1000" ' cat�logo 5
    Igv.Add "nomTributo", "IGV"
    Igv.Add "codTipTributo", "VAT"
    Igv.Add "mtoBaseImponible", Format(Invoice.Item.SaleValue, "0.00")
    Igv.Add "mtoTributo", Format(Invoice.Item.Igv, "0.00")
    Tributos.Add Igv

    Data.Add "cabecera", Cabecera
    Data.Add "detalle", Detalle
    Data.Add "tributos", Tributos

    InvoiceToJson = ConvertToJson(Data)
End Function
