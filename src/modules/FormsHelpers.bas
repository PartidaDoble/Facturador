Attribute VB_Name = "FormsHelpers"
Option Explicit

Public Sub FrmInvoiceShowInformation()
    Dim TypeCurrency As String
    Dim TotalPrice As Double
    Dim Subtotal As Double
    Dim Igv As Double

    TotalPrice = FrmInvoiceSumTotalItems
    Subtotal = TotalPrice / (Prop.Rate.Igv + 1)
    Igv = TotalPrice - Subtotal

    frmInvoice.lblSubTotal = Format(Subtotal, "#,##0.00")
    frmInvoice.lblIGV = Format(Igv, "#,##0.00")
    frmInvoice.lblTotal = Format(TotalPrice, "#,##0.00")

    If frmInvoice.cboTypeCurrency = "Soles" Then TypeCurrency = "PEN"
    If frmInvoice.cboTypeCurrency = "Dólares" Then TypeCurrency = "USD"
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
