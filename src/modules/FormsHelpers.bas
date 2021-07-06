Attribute VB_Name = "FormsHelpers"
Option Explicit

Public Function ToUppercase(InputKey) As Variant
    ToUppercase = Asc(UCase(Chr(InputKey)))
End Function

Public Function OnlyAlphanumeric(KeyAscii)
    OnlyAlphanumeric = IIf(IsAlphanumeric(KeyAscii), Asc(UCase(Chr(KeyAscii))), 0)
End Function

Public Function OnlyAmount(KeyAscci)
    Dim Keys As String
    Keys = "1234567890." & Chr(vbKeyBack)
    OnlyAmount = IIf(InStr(Keys, Chr(KeyAscci)) = 0, 0, KeyAscci)
End Function

Function OnlyInteger(InputKey) As Variant
    Dim Keys As String
    
    Keys = "1234567890" & Chr(vbKeyBack)
    
    
    If InStr(Keys, Chr(InputKey)) = 0 Then
        OnlyInteger = 0
    Else
        OnlyInteger = InputKey
    End If
End Function

Private Function IsAlphanumeric(KeyAscii As Variant) As Boolean
    Dim KeyIsNumber As Boolean
    Dim KeyIsLowercase As Boolean
    Dim KeyIsUppercase As Boolean
    
    KeyIsNumber = 48 <= KeyAscii And KeyAscii <= 57
    KeyIsLowercase = 97 <= KeyAscii And KeyAscii <= 122
    KeyIsUppercase = 65 <= KeyAscii And KeyAscii <= 90
    
    IsAlphanumeric = KeyIsNumber Or KeyIsLowercase Or KeyIsUppercase
End Function

Public Sub FrmInvoiceShowInformation()
    Dim TypeCurrency As AppTypeCurrency
    Dim TotalPrice As Double
    Dim SubTotal As Double
    Dim Igv As Double

    TotalPrice = FrmInvoiceSumTotalItems
    SubTotal = TotalPrice / (Prop.Rate.Igv + 1)
    Igv = TotalPrice - SubTotal

    frmInvoice.lblSubTotal = Format(SubTotal, "#,##0.00")
    frmInvoice.lblIGV = Format(Igv, "#,##0.00")
    frmInvoice.lblTotal = Format(TotalPrice, "#,##0.00")

    If frmInvoice.cboTypeCurrency = "Soles" Then TypeCurrency = AppTypeCurrencyPEN
    If frmInvoice.cboTypeCurrency = "Dólares" Then TypeCurrency = AppTypeCurrencyUSD
    frmInvoice.lblTotalInLetters.Caption = "SON: " & AmountInLetters(TotalPrice, TypeCurrency)
End Sub

Public Function FrmInvoiceSumTotalItems() As Double
    Dim Sum As Double
    Dim Index As Integer

    With frmInvoice.lstItems
        If .ListCount < 1 Then
            Sum = 0
        Else
            For Index = 0 To .ListCount - 1
                Sum = Sum + .List(Index, 3)
            Next Index
        End If
    End With

    FrmInvoiceSumTotalItems = Sum
End Function
