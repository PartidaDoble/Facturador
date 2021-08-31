Attribute VB_Name = "UnitTests"
Option Explicit

Private Sub RunTests()
    ItemEntityTests
    DocumentEntityTest
    DocumentEntity_MethodsTests
End Sub

Private Sub ItemEntityTests()
    Dim Item As New ItemEntity
    
    Item.Quantity = 2
    Item.UnitValue = 50
    Item.IgvRate = 0.18
    
    With Test.It("Item.UnitIgv")
        .AssertEquals 9, Item.UnitIgv
    End With
    
    With Test.It("Item.UnitPrice")
        .AssertEquals 59, Item.UnitPrice
    End With
    
    With Test.It("Item.SaleValue (cantidad * valor unitario)")
        .AssertEquals 100, Item.SaleValue
    End With
    
    With Test.It("Item.Igv")
        .AssertEquals 18, Item.Igv
    End With
    
    With Test.It("Item.SalePrice ((cantidad * valor unitario) + IGV")
        .AssertEquals 118, Item.SalePrice
    End With
    
    With Test.It("Item.SalePrice ((cantidad * valor unitario) + IGV")
        .AssertEquals 118, Item.SalePrice
    End With
End Sub

Private Sub DocumentEntityTest()
    Dim Document As New DocumentEntity
    Dim Item1 As New ItemEntity
    Dim Item2 As New ItemEntity

    Item1.Quantity = 2
    Item1.UnitValue = 50
    Item1.IgvRate = 0.18
    
    Item2.Quantity = 4
    Item2.UnitValue = 50
    Item2.IgvRate = 0.18
    
    Document.AddItem Item1
    Document.AddItem Item2

    With Test.It("Document.SubTotal")
        .AssertEquals 300, Document.SubTotal
    End With

    With Test.It("Document.Igv")
        .AssertEquals 54, Document.Igv
    End With

    With Test.It("Document.Total")
        .AssertEquals 354, Document.Total
    End With
    
    With Test.It("Document.items.Count")
        .AssertEquals 2, Document.Items.Count
    End With
End Sub

Private Sub DocumentEntity_MethodsTests()
    Dim Document As New DocumentEntity
    
    Document.Situation = CdpEnviadoAceptado
    Document.DocType = "01"
    Document.CancelInfo = "X|Foo Bar"
    
    With Test.It("Document.Methods Invoice")
        .AssertTrue Document.IsAccepted
        .AssertTrue Document.IsInvoice
        .AssertTrue Document.IsCanceledNotSent
        .AssertEquals "Factura Electrónica", Document.GetName
        .AssertFalse Document.IsBoleta
        .AssertFalse Document.IsNote
        .AssertFalse Document.IsCanceled
    End With
    
    Document.Situation = CdpXmlGenerado
    Document.DocType = "03"
    Document.CancelInfo = "Sí"
    Document.DailySummary = "RC-20210831-001"
    
    With Test.It("Document.Methods Boleta")
        .AssertFalse Document.IsAccepted
        .AssertTrue Document.IsBoleta
        .AssertTrue Document.IsCanceled
        .AssertTrue Document.SentSummary
        .AssertEquals "Boleta de Venta Electrónica", Document.GetName
        .AssertEquals "1", Document.GetState
        .AssertFalse Document.IsInvoice
        .AssertFalse Document.IsNote
        .AssertFalse Document.IsCanceledNotSent
    End With
    
    Document.DocType = "07"
    Document.DocSerie = "BC01"
    
    With Test.It("Document.Methods Note")
        .AssertTrue Document.IsNote
        .AssertTrue Document.IsBoletaNote
        .AssertEquals "2", Document.GetState
    End With
    
    Document.DocType = "03"
    Document.CancelInfo = "X|Foo Bar"
    With Test.It("Document.Methods Note")
        .AssertTrue Document.IsCanceledNotSent
        .AssertEquals "3", Document.GetState
    End With
End Sub
