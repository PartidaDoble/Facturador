Attribute VB_Name = "FormsHelpers"
Option Explicit

Public Function ToUppercase(InputKey) As Variant
    ToUppercase = Asc(UCase(Chr(InputKey)))
End Function

Public Function OnlyAmount(KeyAscci)
    Dim Keys As String
    Keys = "1234567890." & Chr(vbKeyBack)
    OnlyAmount = IIf(InStr(Keys, Chr(KeyAscci)) = 0, 0, KeyAscci)
End Function

Function OnlyInteger(InputKey) As Variant
    Dim Keys As String
    Keys = "1234567890" & Chr(vbKeyBack)
    OnlyInteger = IIf(InStr(Keys, Chr(InputKey)) = 0, 0, InputKey)
End Function

Public Function OnlyAlphanumeric(KeyAscii)
    OnlyAlphanumeric = IIf(IsAlphanumeric(KeyAscii), Asc(UCase(Chr(KeyAscii))), 0)
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

Public Sub FrmInvoiceCalculateTotals()
    Dim TypeCurrency As String
    Dim TotalPrice As Double
    Dim SubTotal As Double
    Dim Igv As Double

    TotalPrice = FrmInvoiceSumTotalItems
    SubTotal = TotalPrice / (Prop.Rate.Igv + 1)
    Igv = TotalPrice - SubTotal

    frmInvoice.lblSubTotal = Format(SubTotal, "#,##0.00")
    frmInvoice.lblIGV = Format(Igv, "#,##0.00")
    frmInvoice.lblTotal = Format(TotalPrice, "#,##0.00")

    TypeCurrency = IIf(frmInvoice.cboTypeCurrency = "Soles", "PEN", "USD")
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

Public Function GetCustomerName(Ruc As String) As String
    On Error GoTo HandleErrors
    Dim Doc As New Scraping
    Dim Response As String
    
    Doc.gotoPage "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp"
    Doc.Id("txtRuc").FieldValue Ruc
    Doc.Id("btnAceptar").click sleep:=3

    If Trim(Doc.css(".list-group .col-sm-5").Index(0).text) = "Número de RUC:" Then
        Response = Trim(Doc.css(".list-group .col-sm-7").Index(0).text)
    End If
    
    GetCustomerName = Trim(Mid(Response, InStr(Response, "-") + 1))
    DebugLog "Se obtiene el nombre del cliente desde la página web de la SUNAT | RUC: " & Ruc
    Exit Function
HandleErrors:
    WarnLog "Hubo problemas al obtener la información desde la página web de la SUNAT | RUC: " & Ruc, "GetCustomerName"
End Function
