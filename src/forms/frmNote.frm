VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNote 
   ClientHeight    =   8415.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10155
   OleObjectBlob   =   "frmNote.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDocSerie_Change()
    On Error Resume Next
    
    If txtDocType = "07" Then
        txtDocNumber = NextCorrelativeNumber(cboDocSerie)
    ElseIf txtDocType = "08" Then
        txtDocNumber = NextCorrelativeNumber(cboDocSerie)
    End If
End Sub

Private Sub cboRefDocType_Change()
    On Error Resume Next
    
    If GetRefDocTypeCode = "01" Then
        cboRefDocSerie.List = CollectionToArray(GetInvoiceSeries)
    ElseIf GetRefDocTypeCode = "03" Then
        cboRefDocSerie.List = CollectionToArray(GetBoletaSeries)
    End If
End Sub

Private Sub UserForm_Initialize()
    txtEmissionDate = Format(Date, "dd/mm/yyyy")
    cboTypeCurrency.List = Array("Soles", "Dólares")
    lblIGVTitle = "IGV " & Format(Prop.Rate.Igv * 100) & "%:"
    
    cboRefDocType.List = Array("Factura", "Boleta de venta")
    
    If Prop.App.Env = EnvProduction Then
        txtCustomerAddress.Visible = False
        txtCustomerUbigeo.Visible = False
        txtDocType.Visible = False
        txtCustomerDocType.Visible = False
    End If
End Sub

Private Sub cmdShowDocument_Click()
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    Dim DocumentNumber As String
    
    If cboRefDocType = Empty Or cboRefDocSerie = Empty Or txtRefDocNumber = Empty Then
        MsgBox "Debe especificar el tipo, la serie y el número del documento que modifica.", vbInformation, "Subsane la observación"
        cboRefDocType.SetFocus
        Exit Sub
    End If
    
    DocumentNumber = GetRefDocTypeCode & "-" & cboRefDocSerie & "-" & Format(txtRefDocNumber, "00000000")
    Set Document = DocumentRepo.GetItem(DocumentNumber)
    
    If Document Is Nothing Then
        MsgBox "El comprobante " & DocumentNumber & " no está emitido y no se encuentra registrado en la hoja ""Comprobantes de Pago"".", vbInformation, "Comprobante no emitido"
    Else
        frmShowDocument.Caption = UCase(cboRefDocType)
        frmShowDocument.lblEmissionDate = Format(Document.Emission, "dd/mm/yyyy")
        frmShowDocument.lblDocument = Document.DocSerie & "-" & Format(Document.DocNumber, "00000000")
        frmShowDocument.lblCurrency = IIf(Document.TypeCurrency = "PEN", "Soles", "Dólares")
        frmShowDocument.lblCustomer = Document.Customer.Name
        frmShowDocument.lblTotal = Format(Document.Total, "#,##0.00")
        frmShowDocument.Show
    End If
End Sub

Private Function GetRefDocTypeCode() As String
    Select Case cboRefDocType
        Case "Factura"
            GetRefDocTypeCode = "01"
        Case "Boleta de venta"
            GetRefDocTypeCode = "03"
    End Select
End Function

Private Function GetMotiveCode() As String
    Select Case cboMotive
        Case "Anulación de la operación"
            GetMotiveCode = "01"
        Case "Anulación por error en el RUC"
            GetMotiveCode = "02"
        Case "Corrección por error en la descripción"
            GetMotiveCode = "03"
        Case "Descuento global"
            GetMotiveCode = "04"
        Case "Descuento por ítem"
            GetMotiveCode = "05"
        Case "Devolución total"
            GetMotiveCode = "06"
        Case "Devolución por ítem"
            GetMotiveCode = "07"
        Case "Bonificación"
            GetMotiveCode = "08"
        Case "Disminución en el valor"
            GetMotiveCode = "09"
        Case "Otros Conceptos"
            GetMotiveCode = "10"
        
        ' Nota de débito
        Case "Intereses por mora"
            GetMotiveCode = "01"
        Case "Aumento en el valor"
            GetMotiveCode = "02"
        Case "Penalidades / otros conceptos"
            GetMotiveCode = "03"
    End Select
End Function

Private Sub txtMotive_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ToUppercase(KeyAscii)
End Sub

Private Sub txtRefDocNumber_Change()
    On Error Resume Next
    txtRefDocNumber = CInt(txtRefDocNumber)
End Sub

Private Sub txtRefDocNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub txtDocNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub cboTypeCurrency_Change()
    FrmNoteCalculateTotals
End Sub

Private Sub cmdSearchCustomer_Click()
    frmSearchCustomer.Show
End Sub

Private Sub cmdAddProduct_Click()
    frmSearchProduct.Show
End Sub

Private Sub lstItems_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    
    If lstItems.ListCount < 1 Then Exit Sub

    If KeyCode = 46 Then
        lstItems.RemoveItem lstItems.ListIndex
        If lstItems.ListCount <= 8 Then lstItems.Width = 484
    End If

    FrmNoteCalculateTotals
End Sub

Private Sub cmdSave_Click()
    Dim DocumentRepo As New DocumentRepository
    Dim RefDocument As DocumentEntity
    Dim Note As New DocumentEntity
    Dim Customer As New CustomerEntity
    Dim NoteInfo As New NoteInfoEntity
    Dim Item As ItemEntity
    Dim Index As Integer
    Dim DocumentNumber As String
    Dim ElectronicDocumentGenerated As Boolean
    Dim DocumentExists As Boolean
    Dim RefDocNumber As String
    Dim RefDocId As String
    Dim Answer As Integer
    Dim FileName As String
    
    If Not ValidFields Then Exit Sub
    
    If CDate(txtEmissionDate) < Date Then
        Answer = MsgBox("La fecha de emisión debería ser el día de hoy (" & Format(Date, "dd/mm/yyyy") & "). " & _
               "Si de todas formas desea continuar con la emisión, se recomienda utilizar una serie especial " & _
               "para la emisión de comprobantes con fecha anterior, esto con el fin de mantener el correlativo de fecha y numeración." & Chr(13) & Chr(13) & _
               "¿Desea continuar con la emisión del comprobante?", vbYesNo + vbQuestion, "Fecha anterior")
        If Answer = vbNo Then Exit Sub
    End If
    
    RefDocNumber = Trim(cboRefDocSerie) & "-" & Format(Trim(txtRefDocNumber), "00000000")
    RefDocId = GetRefDocTypeCode & "-" & RefDocNumber
    
    Set RefDocument = DocumentRepo.GetItem(RefDocId)
    If RefDocument Is Nothing Then
        MsgBox "El comprobante " & RefDocId & " que quiere modificar no existe. Para modificar un comprobante, debe estar emitido y aceptado.", vbExclamation, "El comprobante no existe"
        Exit Sub
    End If
    
    If Not RefDocument.IsAccepted Then
        MsgBox "El comprobante " & RefDocId & " que quiere modificar debe tener la situación ""Enviado y Aceptado"".", vbExclamation, "No cumple condición"
        Exit Sub
    End If
    
    Note.Emission = CDate(txtEmissionDate)
    Note.EmissionTime = Time
    
    Note.TypeCurrency = IIf(Trim(cboTypeCurrency) = "Soles", "PEN", "USD")
    Note.DocType = txtDocType
    Note.DocSerie = Trim(cboDocSerie)
    Note.DocNumber = Trim(txtDocNumber)
    
    If txtCustomerDocNumber <> Empty Then
        Customer.DocType = txtCustomerDocType
        Customer.DocNumber = txtCustomerDocNumber
        Customer.Name = txtCustomerName
        Customer.Address = txtCustomerAddress
        Customer.Ubigeo = txtCustomerUbigeo
        
        Set Note.Customer = Customer
    End If
    
    NoteInfo.RefDocEmission = RefDocument.Emission
    NoteInfo.RefDocType = RefDocument.DocType
    NoteInfo.RefDocSerie = RefDocument.DocSerie
    NoteInfo.RefDocNumber = RefDocument.DocNumber
    NoteInfo.MotiveCode = GetMotiveCode
    NoteInfo.Motive = Trim(txtMotive)
    Set Note.NoteInfo = NoteInfo
    
    With lstItems
        For Index = 0 To .ListCount - 1
            Set Item = New ItemEntity
            Item.ProductCode = Trim(.List(Index, 4))
            Item.UnitMeasure = Trim(.List(Index, 5))
            Item.Description = Trim(.List(Index, 0))
            Item.Quantity = Trim(.List(Index, 1))
            Item.UnitValue = TaxLess(Trim(.List(Index, 2)), Prop.Rate.Igv)
            Item.IgvRate = Prop.Rate.Igv
            
            Note.AddItem Item
        Next Index
    End With
    
    DocumentNumber = Note.DocType & "-" & Note.DocSerie & "-" & Format(Note.DocNumber, "00000000")
    DocumentExists = GetMatchRow(sheetDocuments, 1, DocumentNumber) > 0
    If DocumentExists Then
        MsgBox "No puede emitir la " & Note.GetName & " " & DocumentNumber & " porque ya fue emitida anteriormente.", vbExclamation, "Duplicado"
        Exit Sub
    End If
    
    CreateNoteJsonFile Note
    RefreshSfsScreen
    GenerateElectronicDocument Note.DocType, Note.DocSerie & "-" & Format(Note.DocNumber, "00000000")
    
    ElectronicDocumentGenerated = ElectronicDocumentExists(Prop.Company.Ruc & "-" & DocumentNumber & ".zip")
    If ElectronicDocumentGenerated Then
        MsgBox "La " & Note.GetName & " electrónica se generó correctamente.", vbInformation, "Comprobante generado"
        InfoLog "La nota electrónica " & DocumentNumber & " se generó correctamente.", "frmNote.cmdSave_Click"
        
        FileName = Prop.Company.Ruc & "-" & Note.Id
        Note.Situation = GetSituationFromCode(DbGetDocumentSituation(FileName))
        Note.Observation = DbGetDocumentObservation(FileName)
        
        DocumentRepo.Add Note
        
        CreatePdf Note
        
        If Prop.App.Premium And Prop.Email.SendWhenEmit And Left(Note.DocSerie, 1) = "F" Then
            Run "ShowFrmSendEmail", Note.Id
        End If
        
        OpenPdf DocumentNumber
    Else
        MsgBox "Error al generar la " & Note.GetName & " " & DocumentNumber, vbCritical, "ERROR"
        ErrorLog "Error al generar la nota electrónica " & DocumentNumber, "frmNote.cmdSave_Click"
    End If
    
    ThisWorkbook.Save
    Unload Me
End Sub

Private Function ValidFields() As Boolean
    If Not IsValidDate(txtEmissionDate) Then
        MsgBox "Ingrese una fecha de emisión válida.", vbExclamation, "Subsane la observación"
        txtEmissionDate.SetFocus
        Exit Function
    End If
    If Date - CDate(txtEmissionDate) > 7 Then
        MsgBox "La fecha del comprobante no puede ser anterior a siete días.", vbExclamation, "Subsane la observación"
        txtEmissionDate.SetFocus
        Exit Function
    End If
    If CDate(txtEmissionDate) > Date Then
        MsgBox "La fecha del comprobante no puede ser una fecha posterior.", vbExclamation, "Subsane la observación"
        txtEmissionDate.SetFocus
        Exit Function
    End If
    If Trim(cboDocSerie) = Empty Then
        MsgBox "Debe ingresar el número de seríe del comprobante.", vbExclamation, "Subsane la observación"
        cboDocSerie.SetFocus
        Exit Function
    End If
    If Trim(txtDocNumber) = Empty Or Not IsNumeric(Trim(txtDocNumber)) Then
        MsgBox "Debe ingresar el número correlativo del comprobante.", vbExclamation, "Subsane la observación"
        txtDocNumber.SetFocus
        Exit Function
    End If
    If Trim(txtCustomerDocNumber) = Empty Or Trim(txtCustomerName) = Empty Then
        MsgBox "Debe ingresar los datos del cliente.", vbExclamation, "Subsane la observación"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If Trim(txtCustomerDocType) = Empty Then
        MsgBox "El tipo de documento del cliente no está registrado en la hoja ""Clientes"".", vbExclamation, "Subsane la observación"
        Exit Function
    End If
    If txtCustomerDocType = "6" And Len(txtCustomerDocNumber) <> 11 Then
        MsgBox "El número de RUC debe tener 11 dígitos.", vbExclamation, "Subsane la observación"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If txtCustomerDocType = "1" And txtCustomerDocNumber <> Empty And Len(txtCustomerDocNumber) <> 8 Then
        MsgBox "El número de DNI debe tener 8 dígitos.", vbExclamation, "Subsane la observación"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If Left(cboDocSerie, 1) = "F" And txtCustomerDocType = "1" Then
        MsgBox "El RUC del cliente debe tener 11 dígitos.", vbExclamation, "Subsane la observación"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If Left(cboDocSerie, 1) = "B" And txtCustomerDocType = "6" Then
        MsgBox "El DNI del cliente debe tener 8 dígitos.", vbExclamation, "Subsane la observación"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If Len(Trim(cboRefDocSerie)) <> 4 Then
        MsgBox "El número de serie del documento que desea modificar debe tener 4 dígitos.", vbInformation, "Subsane la observación"
        cboRefDocSerie.SetFocus
        Exit Function
    End If
    If Trim(cboRefDocType) = Empty Or Trim(cboRefDocSerie) = Empty Or Trim(txtRefDocNumber) = Empty Then
        MsgBox "Debe especificar el tipo, la serie y el número del comprobante que desea modificar.", vbInformation, "Subsane la observación"
        cboRefDocType.SetFocus
        Exit Function
    End If
    If Left(cboDocSerie, 1) = "F" And Left(cboRefDocSerie, 1) = "B" Then
        MsgBox "El tipo de documento modificado por la Nota debe ser una Factura y la serie debe comenzar con F.", vbInformation, "Subsane la observación"
        cboRefDocType.SetFocus
        Exit Function
    End If
    If Left(cboDocSerie, 1) = "B" And Left(cboRefDocSerie, 1) = "F" Then
        MsgBox "El tipo de documento modificado por la Nota debe ser una Boleta de Venta y la serie debe comenzar con B.", vbInformation, "Subsane la observación"
        cboRefDocType.SetFocus
        Exit Function
    End If
    If Trim(cboMotive) = Empty Then
        MsgBox "Debe elegir un tipo de motivo.", vbInformation, "Subsane la observación"
        cboMotive.SetFocus
        Exit Function
    End If
    If Trim(txtMotive) = Empty Then
        MsgBox "Debe especificar una descripción del motivo.", vbInformation, "Subsane la observación"
        txtMotive.SetFocus
        Exit Function
    End If
    If lstItems.ListCount < 1 Then
        MsgBox "Debe ingresar al menos un producto o servicio.", vbExclamation, "Subsane la observación"
        Exit Function
    End If
    
    ValidFields = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub
