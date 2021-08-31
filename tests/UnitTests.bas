Attribute VB_Name = "UnitTests"
Option Explicit

Private Sub RunAllTests()
    ItemEntityTest
    InvoiceEntityTest
End Sub

Private Sub ItemEntityTest()
    Dim Item As New ItemEntity

    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    
    With Test.It("IGV unitario")
        .AssertEquals 9, Item.UnitIgv
    End With
    
    With Test.It("Precio de venta unitario")
        .AssertEquals 59, Item.UnitPrice
    End With

    With Test.It("Valor de venta (cantidad * valor unitario)")
        .AssertEquals 100, Item.SaleValue
    End With
    
    With Test.It("IGV")
        .AssertEquals 18, Item.Igv
    End With

    With Test.It("Precio de de venta (cantidad * valor unitario + IGV")
        .AssertEquals 118, Item.SalePrice
    End With
End Sub

Private Sub InvoiceEntityTest()
    Dim Invoice As New InvoiceEntity
    Dim Item1 As New ItemEntity
    Dim Item2 As New ItemEntity

    Item1.Quantity = 2
    Item1.UnitValue = 50
    Item1.IgvRate = 0.18
    

    Item2.Quantity = 4
    Item2.UnitValue = 50
    Item2.IgvRate = 0.18
    
    Invoice.AddItem Item1
    Invoice.AddItem Item2

    With Test.It("Sub total = suma del valor de venta de cata item")
        .AssertEquals 300, Invoice.SubTotal
    End With

    With Test.It("IGV = suma del IGV de cada item")
        .AssertEquals 54, Invoice.Igv
    End With

    With Test.It("Total = suma del precio de venta de cada item")
        .AssertEquals 354, Invoice.Total
    End With
End Sub
