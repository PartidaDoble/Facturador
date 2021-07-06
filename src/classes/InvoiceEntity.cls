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

Public TypeCurrency As AppTypeCurrency

Public DocType As AppDocType

Public Customer As New CustomerEntity

Private pItems As New Collection

Public Property Get Total() As Double
    Dim Item As Variant
    Dim PriceSum As Double

    For Each Item In pItems
        PriceSum = PriceSum + Item.SalePrice
    Next Item
    
    If DocType = AppDocTypeBoletaVenta And PriceSum > 700 And Customer.DocNumber = Empty And Customer.Name = Empty Then
        Err.Raise AppErrorBVMayor700Soles, , "Total de la venta mayor a 700 soles. Debe ingresar el DNI y los apellidos y nombres del cliente."
    End If

    Total = PriceSum
End Property

Public Property Get SubTotal() As Double
    Dim Item As Variant
    Dim ValueSum As Double

    For Each Item In pItems
        ValueSum = ValueSum + Item.SaleValue
    Next Item

    SubTotal = ValueSum
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
