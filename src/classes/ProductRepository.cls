VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetAll() As Collection
    Dim Product As ProductEntity
    Dim Products As New Collection
    Dim Row As Long
    
    For Row = 2 To GetLastRow(sheetProducts)
        Set Product = New ProductEntity
        
        With sheetProducts
            Product.Code = .Cells(Row, 1)
            Product.UnitMeasure = .Cells(Row, 2)
            Product.Description = .Cells(Row, 3)
            Product.UnitPrice = .Cells(Row, 4)
        End With
        
        Products.Add Product
    Next Row
    
    Set GetAll = Products
End Function

Public Function Add(Product As ProductEntity) As Boolean
    Dim Row As Long
    Row = GetLastRow(sheetProducts) + 1
    With sheetProducts
        .Cells(Row, 1) = Product.Code
        .Cells(Row, 2) = Product.UnitMeasure
        .Cells(Row, 3) = Product.Description
        .Cells(Row, 4) = Product.UnitPrice
    End With
End Function

Public Function Contains(Product As ProductEntity) As Boolean
    Contains = WorksheetFunction.CountIf(sheetProducts.Columns(1), Product.Code) > 0
End Function
