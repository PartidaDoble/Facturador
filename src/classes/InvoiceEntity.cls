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

Public EmissionDate As Date
Public EmissionTime As Date
Public DueDate As Date
Public TypeCurrency As String
Public Customer As New CustomerEntity
Private pItems As New Collection

Public Property Get Total() As Double
    Dim Item As Variant
    Dim PriceSum As Double

    For Each Item In pItems
        PriceSum = PriceSum + Item.SalePrice
    Next Item

    Total = PriceSum
End Property

Public Property Get Subtotal() As Double
    Dim Item As Variant
    Dim ValueSum As Double

    For Each Item In pItems
        ValueSum = ValueSum + Item.SaleValue
    Next Item

    Subtotal = ValueSum
End Property

Public Property Get Igv() As Double
    Dim Item As Variant
    Dim IgvSum As Double
    
    For Each Item In pItems
        IgvSum = IgvSum + Item.Igv
    Next Item
    
    Igv = IgvSum
End Property

Public Property Get Items() As Collection
    Set Items = pItems
End Property

Public Sub AddItem(Item As ItemEntity)
    pItems.Add Item
End Sub
