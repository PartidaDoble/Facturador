VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pRate As RateType

Property Get Igv() As Double
    Igv = 0.18
End Property

Property Get Rate() As RateType
    pRate.Igv = 0.18
    Rate = pRate
End Property
