Attribute VB_Name = "Common"
Option Explicit

Public Type RateType
    Igv As Double
End Type

Public Enum AppTypeCurrency
    AppTypeCurrencyPEN
    AppTypeCurrencyUSD
End Enum

Public Enum AppError
    AppErrorBVMayor700Soles = 65400
End Enum

Public Enum AppDocType
    AppDocTypeBoletaVenta
    AppDocTypeFactura
End Enum

Public Enum AppTypeDocIdenty
    AppTypeDocIdentyDNI
    AppTypeDocIdentyRUC
End Enum

Public Function Prop() As AppProperties
    Dim Properties As New AppProperties
    Set Prop = Properties
End Function

Public Function Test() As VBAUnit
    Dim UnitTest As New VBAUnit
    Set Test = UnitTest
End Function

Sub Fun()
    Dim i As Integer
    
    For i = 1 To 100
        Debug.Print i
    Next
End Sub
