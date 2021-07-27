VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInvoice 
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10155
   OleObjectBlob   =   "frmInvoice.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    txtEmissionDate = Format(Date, "dd/mm/yyyy")
    cboTypeCurrency.List = Array("Soles", "Dólares")
    lblIGVTitle = "IGV " & Format(Prop.Rate.Igv * 100) & "%:"
    If Not SfsIsRunning Then
        RunSfs
        MsgBox "El Facturador SUNAT no se está ejecutando." & Chr(13) & "Ejecutando automáticamente..." & Chr(13) & Chr(13) & _
            "Espere a que el Facturador SUNAT esté listo, puede demorar hasta dos minutos (depende de la velocidad de su computadora).", vbExclamation, ""
    End If
End Sub

Private Sub txtDocNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub cboTypeCurrency_Change()
    FrmInvoiceCalculateTotals
End Sub

Private Sub cmdSearchCustomer_Click()
    frmSearchCustomer.Show
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

    FrmInvoiceCalculateTotals
End Sub

Private Sub cmdSave_Click()
    Dim Invoice As New InvoiceEntity
    Dim Item As ItemEntity
    Dim Index As Integer
    Dim DocumentNumber As String
    Dim ElectronicDocumentGenerated As Boolean
    Dim ElectronicDocumentSent As Boolean
    Dim DocumentRepo As New DocumentRepository
    Dim ResponseSunat As Dictionary
    
    If Not ValidFields Then Exit Sub
    
    Invoice.EmissionDate = DateValue(txtEmissionDate)
    Invoice.EmissionTime = Time
    
    Invoice.TypeCurrency = IIf(cboTypeCurrency = "Soles", "PEN", "USD")
    Invoice.DocType = txtDocType
    Invoice.DocSerie = cboDocSerie
    Invoice.DocNumber = Trim(txtDocNumber)
    
    Invoice.Customer.DocType = txtCustomerDocType
    Invoice.Customer.DocNumber = txtCustomerDocNumber
    Invoice.Customer.Name = txtCustomerName
    
    With lstItems
        For Index = 0 To .ListCount - 1
            Set Item = New ItemEntity
            Item.ProductCode = Trim(.List(Index, 4))
            Item.UnitMeasure = Trim(.List(Index, 5))
            Item.Description = Trim(.List(Index, 0))
            Item.Quantity = Trim(.List(Index, 1))
            Item.UnitValue = TaxLess(Trim(.List(Index, 2)), Prop.Rate.Igv)
            
            Invoice.AddItem Item
        Next Index
    End With
    
    DocumentNumber = Invoice.DocType & "-" & Invoice.DocSerie & "-" & Format(Invoice.DocNumber, "00000000")
    CreateInvoiceJsonFile Invoice
    RefreshSfsScreen
    GenerateElectronicDocument Invoice.DocType, Invoice.DocSerie & "-" & Format(Invoice.DocNumber, "00000000")
    ElectronicDocumentGenerated = ElectronicDocumentExists(DocumentNumber)
    If ElectronicDocumentGenerated Then
        If Prop.App.Internet Then
            SendElectronicDocument Invoice.DocType, Invoice.DocSerie & "-" & Format(Invoice.DocNumber, "00000000")
            ElectronicDocumentSent = ResponseFileExists(Invoice.DocType, Invoice.DocSerie, Invoice.DocNumber)
            If ElectronicDocumentSent Then
                Set ResponseSunat = GetResponseSunat(Invoice.DocType, Invoice.DocSerie, Invoice.DocNumber)
                If ResponseSunat("ResponseCode") = "0" Then
                    DocumentRepo.Add CreateDocumentEntity(Invoice)
                    MsgBox "El comprobante de pago electrónico se generó y se envió a la SUNAT correctamente.", vbInformation, "Comprobante enviado"
                    InfoLog "El comprobante de pago " & DocumentNumber & " se generó y se envió a la SUNAT correctamente.", "frmInvoice.cmdSave_Click"
                    CreatePdf DocumentNumber
                    OpenPdf DocumentNumber
                Else
                    DocumentRepo.Add CreateDocumentEntity(Invoice)
                    MsgBox "El comprobante de pago electrónico enviado a la SUNAT tiene observaciones.", vbExclamation, "Comprobante con observaciones"
                    WarnLog "El comprobante de pago " & DocumentNumber & " enviado a la SUNAT tiene observaciones.", "frmInvoice.cmdSave_Click"
                    CreatePdf DocumentNumber
                    OpenPdf DocumentNumber
                End If
            Else
                DocumentRepo.Add CreateDocumentEntity(Invoice)
                MsgBox "Debido a su conexión a internet o a problemas con los servidores de la SUNAT, el comprobante no pudo ser enviado." & _
                    "Sin embargo, el comprobante de pago electrónico se generó correctamente.", vbExclamation, "Comprobante generado"
                WarnLog "Error al enviar el comprobante de pago a la SUNAT. Sin embargo, se generó el comprobante de pago: " & DocumentNumber, "frmInvoice.cmdSave_Click"
                CreatePdf DocumentNumber
                OpenPdf DocumentNumber
            End If
        Else
            DocumentRepo.Add CreateDocumentEntity(Invoice)
            MsgBox "El comprobante de pago electrónico se generó correctamente.", vbInformation, "Comprobante generado"
            InfoLog "El comprobante de pago " & DocumentNumber & " se generó correctamente.", "frmInvoice.cmdSave_Click"
            CreatePdf DocumentNumber
            OpenPdf DocumentNumber
        End If
    Else
        MsgBox "Error al generar el comprobante de pago electrónico.", vbCritical, "ERROR"
        ErrorLog "Error al generar el comprobante de pago electrónico. Comprobante: " & DocumentNumber, "frmInvoice.cmdSave_Click"
    End If
    Unload Me
End Sub

Private Function ValidFields() As Boolean
    If Trim(cboDocSerie) = Empty Then
        MsgBox "Debe ingresar el número de seríe del comprobante de pago.", vbExclamation, "Subsane la observación"
        cboDocSerie.SetFocus
        Exit Function
    End If
    If Trim(txtDocNumber) = Empty Or Not IsNumeric(Trim(txtDocNumber)) Then
        MsgBox "Debe ingresar el número del comprobante de pago.", vbExclamation, "Subsane la observación"
        txtDocNumber.SetFocus
        Exit Function
    End If
    If txtDocType = "03" And lblTotal > 700 And (Trim(txtCustomerDocNumber) = Empty Or Trim(txtCustomerName) = Empty) Then
        MsgBox "El total de la venta es mayor a 700 soles. Debe ingresar el DNI y los apellidos y nombres del cliente.", vbExclamation, "Subsane la observación"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If txtDocType = "01" And (Trim(txtCustomerDocNumber) = Empty Or Trim(txtCustomerName) = Empty) Then
        MsgBox "Debe ingresar el RUC y el nombre del cliente.", vbExclamation, "Subsane la observación"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If txtCustomerDocType = "6" And Len(txtCustomerDocNumber) <> 11 Then
        MsgBox "El número de RUC debe tener 11 dígitos.", vbExclamation, "Subsane la observación"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If txtCustomerDocType = "3" And txtCustomerDocNumber <> Empty And Len(txtCustomerDocNumber) <> 8 Then
        MsgBox "El número de DNI debe tener 8 dígitos.", vbExclamation, "Subsane la observación"
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
