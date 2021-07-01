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
    Dim PriceSum As Double
    
    For Each Item In pItems
        PriceSum = PriceSum + Item.SalePrice
    Next Item
    
    Total = PriceSum
End Property

Property Get SubTotal() As Double
    Dim Item As Variant
    Dim ValueSum As Double
    
    For Each Item In pItems
        ValueSum = ValueSum + Item.SaleValue
    Next Item
    
    SubTotal = ValueSum
End Property

Property Get Igv() As Double
    Dim Item As Variant
    Dim IgvSum As Double
    
    For Each Item In pItems
        IgvSum = IgvSum + Item.Igv
    Next Item
    
    Igv = IgvSum
End Property

Public Sub AddItem(Item As ItemEntity)
    pItems.Add Item
End Sub

Property Get Items() As Collection
    Set Items = pItems
End Property

