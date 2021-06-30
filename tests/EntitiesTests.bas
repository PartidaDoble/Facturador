Attribute VB_Name = "EntitiesTests"
Option Explicit

Sub un_item_deberia()
    Dim item As New ItemEntity
    
    item.Quantity = 2
    item.UnitValue = 50
    
    With Test.It("calcular su igv")
        .AssertEquals 18, item.Igv
    End With
    
    With Test.It("calcular su valor de venta de venta (cantidad * valor unitario)")
        .AssertEquals 100, item.SaleValue
    End With
    
    With Test.It("calcular su precio de de venta (cantidad * valor unitario + igv")
        .AssertEquals 118, item.SalePrice
    End With
End Sub

Sub una_factura_deberia()
    Dim invoice As New InvoiceEntity
    Dim item1 As New ItemEntity
    Dim item2 As New ItemEntity

    item1.Quantity = 2
    item1.UnitValue = 50
    
    item2.Quantity = 4
    item2.UnitValue = 50
    
    invoice.AddItem item1
    invoice.AddItem item2
    
    With Test.It("tener un valor de venta total igual a la suma del valor de venta de cada item")
        .AssertEquals 300, invoice.SubTotal
    End With
    
    With Test.It("tener un IGV igual a la suma del IGV de cada item")
        .AssertEquals 54, invoice.Igv
    End With

    With Test.It("tener un precio de venta total igual a la suma del precio de venta de cada item")
        .AssertEquals 354, invoice.Total
    End With
End Sub
