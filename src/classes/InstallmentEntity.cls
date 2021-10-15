VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstallmentEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PaymentDate As Date

Public Amount As Double

Public TypeCurrency As String

Property Get CurrencySymbol() As String
    CurrencySymbol = IIf(TypeCurrency = "PEN", "S/", "US$")
End Property
