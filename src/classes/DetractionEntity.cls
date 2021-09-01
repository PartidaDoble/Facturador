VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DetractionEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Code As String

Public Percentage As Integer

Public Amount As Double

Public PaymentMethod As String

Public Function ConcatData() As String
    ConcatData = Join(Array(Code, Percentage, Format(Amount, "#,##0.00"), PaymentMethod), "-")
End Function

Public Function Legend() As String
    Legend = "OPERACIÓN SUJETA AL SPOT"
End Function

Public Function LegendLarge() As String
    LegendLarge = "OPERACIÓN SUJETA AL SPOT S/ " & Format(Amount, "#,##0.00") & " (" & Percentage & "%)"
End Function
