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
    Dim item As Variant
    Dim priceSum As Double
    
    For Each item In pItems
        priceSum = priceSum + item.SalePrice
    Next item
    
    Total = priceSum
End Property

Property Get SubTotal() As Double
    Dim item As Variant
    Dim valueSum As Double
    
    For Each item In pItems
        valueSum = valueSum + item.SaleValue
    Next item
    
    SubTotal = valueSum
End Property

Property Get Igv() As Double
    Dim item As Variant
    Dim igvSum As Double
    
    For Each item In pItems
        igvSum = igvSum + item.Igv
    Next item
    
    Igv = igvSum
End Property

Public Sub AddItem(item As ItemEntity)
    pItems.Add item
End Sub





