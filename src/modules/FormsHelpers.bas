Attribute VB_Name = "FormsHelpers"
Option Explicit

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
