Attribute VB_Name = "UseCases"
Option Explicit

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
    Cabecera.Add "codLocalEmisor", Prop.CodLocalEmisor
    If Invoice.Customer.DocType = AppTypeDocIdentyDNI Then Cabecera.Add "tipDocUsuario", "1"
    If Invoice.Customer.DocType = AppTypeDocIdentyRUC Then Cabecera.Add "tipDocUsuario", "6"
    Cabecera.Add "numDocUsuario", IIf(Invoice.Customer.DocNumber = Empty, "00000000", Invoice.Customer.DocNumber)
    Cabecera.Add "rznSocialUsuario", IIf(Invoice.Customer.Name = Empty, "varios", Invoice.Customer.Name)
    If Invoice.TypeCurrency = AppTypeCurrencyPEN Then Cabecera.Add "tipMoneda", "PEN"
    If Invoice.TypeCurrency = AppTypeCurrencyUSD Then Cabecera.Add "tipMoneda", "USD"
    Cabecera.Add "sumTotTributos", Format(Invoice.Igv, "0.00") ' sumatoria Tributos
    Cabecera.Add "sumTotValVenta", Format(Invoice.SubTotal, "0.00") ' suma valor de venta de items
    Cabecera.Add "sumPrecioVenta", Format(Invoice.Total, "0.00") ' suma precio de venta de items
    Cabecera.Add "sumDescTotal", "0.00"
    Cabecera.Add "sumOtrosCargos", "0.00"
    Cabecera.Add "sumTotalAnticipos", "0.00"
    Cabecera.Add "sumImpVenta", Format(Invoice.Total, "0.00") ' importe total de la venta
    Cabecera.Add "ublVersionId", "2.1"
    Cabecera.Add "customizationId", "2.0" ' revisar

    For Each Item In Invoice.Items
        Set DetalleItem = New Dictionary
        DetalleItem.Add "codUnidadMedida", "NIU" ' catálogo 3
        DetalleItem.Add "ctdUnidadItem", Format(Item.Quantity, "0.00")
        DetalleItem.Add "codProducto", Item.Code
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
