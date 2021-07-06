VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInvoice 
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10170
   OleObjectBlob   =   "frmInvoice.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearchCustomer_Click()
    frmSearchCustomer.Show
End Sub

Private Sub txtDocNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub UserForm_Initialize()
    frmInvoice.Caption = "BOLETA DE VENTA"
    txtEmissionDate = Format(Date, "dd/mm/yyyy")
    cboTypeCurrency.List = Array("Soles", "Dólares")
    lblIGVTitle = "IGV " & Format(Prop.Rate.Igv * 100) & "%:"
End Sub

Private Sub cboTypeCurrency_Change()
    FrmInvoiceShowInformation
End Sub

Private Sub cmdAddProduct_Click()
    frmAddProduct.Show
End Sub

Private Sub lstItems_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    
    If lstItems.ListCount < 1 Then Exit Sub

    If KeyCode = 46 Then
        lstItems.RemoveItem lstItems.ListIndex
    End If

    FrmInvoiceShowInformation
End Sub

Private Sub cmdSave_Click()
    Dim Invoice As New InvoiceEntity
    Dim Item As ItemEntity
    Dim Index As Integer

    If Not ValidFields Then Exit Sub
    
    Invoice.EmissionDate = DateValue(txtEmissionDate)
    Invoice.EmissionTime = Time
    
    If cboTypeCurrency = "Soles" Then Invoice.TypeCurrency = AppTypeCurrencyPEN
    If cboTypeCurrency = "Dólares" Then Invoice.TypeCurrency = AppTypeCurrencyUSD
    
    If txtDocType = "01" Then Invoice.DocType = AppDocTypeFactura
    If txtDocType = "03" Then Invoice.DocType = AppDocTypeBoletaVenta
    
    If txtCustomerDocType = "1" Then Invoice.Customer.DocType = AppTypeDocIdentyDNI
    If txtCustomerDocType = "6" Then Invoice.Customer.DocType = AppTypeDocIdentyRUC
    Invoice.Customer.DocNumber = txtCustomerDocNumber
    Invoice.Customer.Name = txtCustomerName
    
    With lstItems
        For Index = 0 To .ListCount - 1
            Set Item = New ItemEntity
            Item.Code = .List(Index, 4)
            Item.UnitMeasurement = .List(Index, 5)
            Item.Description = Trim(.List(Index, 0))
            Item.Quantity = .List(Index, 1)
            Item.UnitValue = TaxLess(.List(Index, 2), Prop.Rate.Igv)
            
            Invoice.AddItem Item
        Next Index
    End With

    Debug.Print InvoiceToJson(Invoice)

    MsgBox "El documento electrónico se generó correctamente.", vbInformation, "Documento generado"
    Unload Me
    Exit Sub
End Sub

Private Function ValidFields() As Boolean
    If Trim(txtDocNumber) = Empty And Not IsNumeric(Trim(txtDocNumber)) Then
        MsgBox "Debe ingresar el número del comprobante.", vbExclamation, "Subsane la observación"
        txtDocNumber.SetFocus
        Exit Function
    End If
    If txtDocType = "03" And lblTotal > 700 And Trim(txtCustomerDocNumber) = Empty And Trim(txtCustomerName) = Empty Then
        MsgBox "El total de la venta es mayor a 700 soles. Debe ingresar el DNI y los apellidos y nombres del cliente.", vbExclamation, "Subsane la observación"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If txtDocType = "01" And Trim(txtCustomerDocNumber) = Empty And Trim(txtCustomerName) = Empty Then
        MsgBox "Debe ingresar el RUC y el nombre del cliente.", vbExclamation, "Subsane la observación"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If txtCustomerDocType = "1" And Len(txtCustomerDocNumber) <> 11 Then
        MsgBox "El número de RUC debe tener 11 dígitos.", vbExclamation, "Subsane la observación"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If txtCustomerDocType = "3" And txtCustomerDocNumber <> Empty And Len(txtCustomerDocNumber) <> 8 Then
        MsgBox "La boleta de venta debe tener 8 dígitos.", vbExclamation, "Subsane la observación"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If lstItems.ListCount < 1 Then
        MsgBox "Debe ingresar productos para vender.", vbExclamation, "Subsane la observación"
        Exit Function
    End If
    
    ValidFields = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub
