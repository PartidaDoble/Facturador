VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductRepositoryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetAll() As Collection
    Dim Product1 As New ProductEntity
    Dim Product2 As New ProductEntity
    Dim Product3 As New ProductEntity
    Dim Products As New Collection

    Product1.Description = "Producto foo 1Producto foo 1Producto foo 1Producto foo 1Producto foo 1Producto foo 1"
    Product1.UnitPrice = 2234025.2
    
    Product2.Description = "Producto bar 2"
    Product2.UnitPrice = 900000.22
    
    Product3.Description = "Producto baz 3"
    Product3.UnitPrice = 15.33
    
    Products.Add Product1
    Products.Add Product2
    Products.Add Product3
    
    Set GetAll = Products
End Function
