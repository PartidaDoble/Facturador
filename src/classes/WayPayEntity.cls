VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WayPayEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Way As String

Public NetAmountPending As Double

Public TypeCurrency As String

Public Installments As New Collection

Property Get CurrencySymbol() As String
    CurrencySymbol = IIf(TypeCurrency = "PEN", "S/", "US$")
End Property
