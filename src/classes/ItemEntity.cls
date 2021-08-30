VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ProductCode As String

Public UnitMeasure As String

Public Description As String

Public Quantity As Long

Public UnitValue As Double

Public IgvRate As Double

Public Property Get UnitIgv() As Double
    UnitIgv = UnitValue * IgvRate
End Property

Public Property Get UnitPrice() As Double
    UnitPrice = UnitValue + UnitIgv
End Property

Public Property Get SaleValue() As Double
    SaleValue = Quantity * UnitValue
End Property

Public Property Get Igv() As Double
    Igv = Quantity * UnitValue * IgvRate
End Property

Public Property Get SalePrice() As Double
    SalePrice = Quantity * UnitValue + Igv
End Property


