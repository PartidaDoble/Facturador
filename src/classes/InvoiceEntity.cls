VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InvoiceEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pItems As New Collection

Property Get Total() As Double
    Dim Item As Variant
    Dim priceSum As Double
    
    For Each Item In pItems
        priceSum = priceSum + Item.SalePrice
    Next Item
    
    Total = priceSum
End Property

Property Get SubTotal() As Double
    Dim Item As Variant
    Dim valueSum As Double
    
    For Each Item In pItems
        valueSum = valueSum + Item.SaleValue
    Next Item
    
    SubTotal = valueSum
End Property

Property Get Igv() As Double
    Dim Item As Variant
    Dim igvSum As Double
    
    For Each Item In pItems
        igvSum = igvSum + Item.Igv
    Next Item
    
    Igv = igvSum
End Property

Public Sub AddItem(Item As ItemEntity)
    pItems.Add Item
End Sub
