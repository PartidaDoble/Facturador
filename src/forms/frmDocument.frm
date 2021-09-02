VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDocument
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10155
   OleObjectBlob   =   "frmDocument.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDocSerie_Change()
    On Error Resume Next

    If txtDocType = "01" Then
        txtDocNumber = NextCorrelativeNumber(cboDocSerie)
    ElseIf txtDocType = "03" Then
        txtDocNumber = NextCorrelativeNumber(cboDocSerie)
    ElseIf txtDocType = "07" Then
        txtDocNumber = NextCorrelativeNumber(cboDocSerie)
    ElseIf txtDocType = "08" Then
        txtDocNumber = NextCorrelativeNumber(cboDocSerie)
    End If

    If Left(cboDocSerie, 1) = "F" Then
        lblCustomerDocType = "RUC:"
    ElseIf Left(cboDocSerie, 1) = "B" Then
        lblCustomerDocType = "DNI:"
    End If
End Sub

Private Sub cmdReferenceDocument_Click()
    If txtDocType = "07" Then
        frmReferenceDocument.cboMotive.List = Array("Anulaci�n de la operaci�n", "Anulaci�n por error en el RUC", "Correcci�n por error en la descripci�n", "Descuento global", "Descuento por �tem", "Devoluci�n total", "Devoluci�n por �tem", "Bonificaci�n", "Disminuci�n en el valor", "Otros Conceptos")
    ElseIf txtDocType = "08" Then
        frmReferenceDocument.cboMotive.List = Array("Intereses por mora", "Aumento en el valor", "Penalidades / otros conceptos")
    End If

    frmReferenceDocument.Show
End Sub

Private Sub cmdShowDetraction_Click()
    If Prop.App.Premium Then
        frmDetraction.txtTotal = lblTotal
        frmDetraction.Show
    Else
        MsgBox "Esta funcionalidad no est� disponible en la versi�n libre. " & _
               "Para ocultar este bot�n borre el n�mero de cuenta de detracciones en la hoja Configuraci�n.", vbInformation, "No disponible"
    End If
End Sub

Private Sub UserForm_Initialize()
    txtEmissionDate = Format(Date, "dd/mm/yyyy")
    cboTypeCurrency.List = Array("Soles", "D�lares")
    lblIGVTitle = "IGV " & Format(Prop.Rate.Igv * 100) & "%:"

    If Prop.App.Env = EnvProduction Then
        txtCustomerAddress.Visible = False
        txtCustomerUbigeo.Visible = False
        txtDocType.Visible = False
        txtCustomerDocType.Visible = False
        txtDetractionData.Visible = False
    End If
End Sub

Private Sub txtDocNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub cboTypeCurrency_Change()
    FrmDocumentCalculateTotals
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

    FrmDocumentCalculateTotals
End Sub

Private Function GetReferenceDocument(Data As String) As Dictionary
    Dim ReferenceDocument As New Dictionary

    ReferenceDocument.Add "DocType", Split(Data, "-")(0)
    ReferenceDocument.Add "DocSerie", Split(Data, "-")(1)
    ReferenceDocument.Add "DocNumber", Split(Data, "-")(2)
    ReferenceDocument.Add "MotiveCode", Split(Data, "-")(3)
    ReferenceDocument.Add "Motive", Split(Data, "-")(4)

    Set GetReferenceDocument = ReferenceDocument
End Function

Private Sub cmdSave_Click()
    Dim DocumentRepo As New DocumentRepository
    Dim Document As New DocumentEntity
    Dim Customer As New CustomerEntity
    Dim Detraction As New DetractionEntity
    Dim Item As ItemEntity
    Dim Index As Integer
    Dim DocumentNumber As String
    Dim ElectronicDocumentGenerated As Boolean
    Dim DocumentExists As Boolean
    Dim Answer As Integer
    Dim OperationCode As String
    Dim FileName As String

    Dim NoteInfo As New NoteInfoEntity
    Dim RefDocument As DocumentEntity
    Dim ReferenceDocument As New Dictionary
    Dim RefDocId As String

    If Not ValidFields Then Exit Sub

    If CDate(txtEmissionDate) < Date Then
        Answer = MsgBox("La fecha de emisi�n deber�a ser el d�a de hoy (" & Format(Date, "dd/mm/yyyy") & "). " & _
               "Si de todas formas desea continuar con la emisi�n, se recomienda utilizar una serie especial " & _
               "para la emisi�n de comprobantes con fecha anterior, esto con el fin de mantener el correlativo de fecha y numeraci�n." & Chr(13) & Chr(13) & _
               "�Desea continuar con la emisi�n del comprobante?", vbYesNo + vbQuestion, "Fecha anterior")
        If Answer = vbNo Then Exit Sub
    End If

    If txtDocType = "07" Or txtDocType = "08" Then
        If Not ValidNoteFields Then Exit Sub

        Set ReferenceDocument = GetReferenceDocument(txtReferenceDocument)

        RefDocId = ReferenceDocument("DocType") & "-" & ReferenceDocument("DocSerie") & "-" & Format(ReferenceDocument("DocNumber"), "00000000")

        Set RefDocument = DocumentRepo.GetItem(RefDocId)
        If RefDocument Is Nothing Then
            MsgBox "El comprobante " & RefDocId & " que quiere modificar no existe. " & _
                   "Para modificar un comprobante, debe estar emitido y aceptado.", vbExclamation, "El comprobante no existe"
            Exit Sub
        End If

        If Not RefDocument.IsAccepted Then
            MsgBox "El comprobante " & RefDocId & " que pretende modificar debe tener la situaci�n " & _
                   """Enviado y Aceptado"".", vbExclamation, "No cumple condici�n"
            Exit Sub
        End If

        NoteInfo.RefDocEmission = RefDocument.Emission
        NoteInfo.RefDocType = RefDocument.DocType
        NoteInfo.RefDocSerie = RefDocument.DocSerie
        NoteInfo.RefDocNumber = RefDocument.DocNumber
        NoteInfo.MotiveCode = ReferenceDocument("MotiveCode")
        NoteInfo.Motive = ReferenceDocument("Motive")

        Set Document.NoteInfo = NoteInfo
    End If

    Document.Emission = CDate(txtEmissionDate)
    Document.EmissionTime = Time
    Document.TypeCurrency = IIf(Trim(cboTypeCurrency) = "Soles", "PEN", "USD")
    Document.DocType = txtDocType
    Document.DocSerie = Trim(cboDocSerie)
    Document.DocNumber = Trim(txtDocNumber)

    If txtCustomerDocNumber <> Empty Then
        Customer.DocType = txtCustomerDocType
        Customer.DocNumber = txtCustomerDocNumber
        Customer.Name = txtCustomerName
        Customer.Address = txtCustomerAddress
        Customer.Ubigeo = txtCustomerUbigeo

        Set Document.Customer = Customer
    End If

    OperationCode = "0101"

    If txtDetractionData <> Empty Then
        Detraction.Code = Split(txtDetractionData, "-")(0)
        Detraction.Percentage = Split(txtDetractionData, "-")(1)
        Detraction.Amount = Split(txtDetractionData, "-")(2)
        Detraction.PaymentMethod = Split(txtDetractionData, "-")(3)

        OperationCode = "1001"
        If Detraction.Code = "004" Then OperationCode = "1002"
        If Detraction.Code = "028" Then OperationCode = "1003"
        If Detraction.Code = "027" Then OperationCode = "1004"

        Set Document.Detraction = Detraction
    End If

    Document.OperationCode = OperationCode

    With lstItems
        For Index = 0 To .ListCount - 1
            Set Item = New ItemEntity
            Item.ProductCode = Trim(.List(Index, 4))
            Item.UnitMeasure = Trim(.List(Index, 5))
            Item.Description = Trim(.List(Index, 0))
            Item.Quantity = Trim(.List(Index, 1))
            Item.UnitValue = TaxLess(Trim(.List(Index, 2)), Prop.Rate.Igv)
            Item.IgvRate = Prop.Rate.Igv

            Document.AddItem Item
        Next Index
    End With

    DocumentNumber = Document.DocType & "-" & Document.DocSerie & "-" & Format(Document.DocNumber, "00000000")
    DocumentExists = GetMatchRow(sheetDocuments, 1, DocumentNumber) > 0
    If DocumentExists Then
        MsgBox "No puede emitir la " & Document.GetName & " " & DocumentNumber & " porque ya fue emitida anteriormente.", vbExclamation, "Duplicado"
        Exit Sub
    End If

    CreateDocumentJsonFile Document
    RefreshSfsScreen
    GenerateElectronicDocument Document.DocType, Document.DocSerie & "-" & Format(Document.DocNumber, "00000000")

    ElectronicDocumentGenerated = ElectronicDocumentExists(Prop.Company.Ruc & "-" & DocumentNumber & ".zip")
    If ElectronicDocumentGenerated Then
        MsgBox "La " & Document.GetName & " se gener� correctamente.", vbInformation, "Comprobante generado"
        InfoLog "El comprobante electr�nico " & DocumentNumber & " se gener� correctamente.", "frmDocument.cmdSave_Click"

        FileName = Prop.Company.Ruc & "-" & Document.Id
        Document.Situation = GetSituationFromCode(DbGetDocumentSituation(FileName))
        Document.Observation = DbGetDocumentObservation(FileName)

        DocumentRepo.Add Document

        CreatePdf Document

        If Prop.App.Premium And Prop.Email.SendWhenEmit And Document.DocType = "01" Then
            Run "ShowFrmSendEmail", Document.Id
        End If

        OpenPdf DocumentNumber
    Else
        MsgBox "Error al generar la " & Document.GetName & " " & DocumentNumber, vbCritical, "ERROR"
        ErrorLog "Error al generar el comprobante electr�nico " & DocumentNumber, "frmDocument.cmdSave_Click"
    End If

    SaveLastSerialNumber

    ThisWorkbook.Save
    Unload Me
End Sub

Private Function ValidFields() As Boolean
    If Not IsValidDate(txtEmissionDate) Then
        MsgBox "Ingrese una fecha de emisi�n v�lida. El formato de fecha es dd/mm/yyyy.", vbExclamation, "Subsane la observaci�n"
        txtEmissionDate.SetFocus
        Exit Function
    End If
    If Date - CDate(txtEmissionDate) > 7 Then
        MsgBox "La fecha del comprobante no puede ser anterior a siete d�as.", vbExclamation, "Subsane la observaci�n"
        txtEmissionDate.SetFocus
        Exit Function
    End If
    If CDate(txtEmissionDate) > Date Then
        MsgBox "La fecha del comprobante no puede ser una fecha posterior.", vbExclamation, "Subsane la observaci�n"
        txtEmissionDate.SetFocus
        Exit Function
    End If
    If Trim(cboDocSerie) = Empty Then
        MsgBox "Debe ingresar el n�mero de ser�e del comprobante.", vbExclamation, "Subsane la observaci�n"
        cboDocSerie.SetFocus
        Exit Function
    End If
    If Trim(txtDocNumber) = Empty Or Not IsNumeric(Trim(txtDocNumber)) Then
        MsgBox "Debe ingresar el n�mero correlativo del comprobante.", vbExclamation, "Subsane la observaci�n"
        txtDocNumber.SetFocus
        Exit Function
    End If

    If txtDocType = "03" And lblTotal > 700 And (Trim(txtCustomerDocNumber) = Empty Or Trim(txtCustomerName) = Empty) Then
        MsgBox "El total de la venta es mayor a 700 soles. Debe ingresar el DNI y los apellidos y nombres del cliente.", vbExclamation, "Subsane la observaci�n"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If Left(cboDocSerie, 1) = "F" And txtCustomerDocType = "1" Then
        MsgBox "El documento de identificaci�n del cliente debe ser un RUC, no un DNI", vbExclamation, "Subsane la observaci�n"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If Left(cboDocSerie, 1) = "B" And txtCustomerDocType = "6" Then
        MsgBox "El documento de identificaci�n del cliente debe ser un RUC, no un DNI", vbExclamation, "Subsane la observaci�n"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If txtDocType = "01" And (Trim(txtCustomerDocNumber) = Empty Or Trim(txtCustomerName) = Empty) Then
        MsgBox "Debe ingresar el RUC y el nombre del cliente.", vbExclamation, "Subsane la observaci�n"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If txtCustomerDocType = "6" And Len(txtCustomerDocNumber) <> 11 Then
        MsgBox "El n�mero de RUC debe tener 11 d�gitos.", vbExclamation, "Subsane la observaci�n"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If
    If txtCustomerDocType = "1" And txtCustomerDocNumber <> Empty And Len(txtCustomerDocNumber) <> 8 Then
        MsgBox "El n�mero de DNI debe tener 8 d�gitos.", vbExclamation, "Subsane la observaci�n"
        cmdSearchCustomer.SetFocus
        Exit Function
    End If

    If lstItems.ListCount < 1 Then
        MsgBox "Debe ingresar al menos un producto o servicio.", vbExclamation, "Subsane la observaci�n"
        Exit Function
    End If
    If (txtDocType = "01" Or txtDocType = "03") And lblTotal = 0 Then
        MsgBox "El total debe ser mayor a cero.", vbExclamation, "Subsane la observaci�n"
        Exit Function
    End If

    ValidFields = True
End Function

Private Function ValidNoteFields() As Boolean
    If CharCount(txtReferenceDocument, "-") <> 4 Then
        MsgBox "Debe ingresar los datos del documento que modifica.", vbExclamation, "Subsane la observaci�n"
        cmdReferenceDocument.SetFocus
        Exit Function
    End If
    If Left(cboDocSerie, 1) = "F" And Mid(txtReferenceDocument, 4, 1) = "B" Then
        MsgBox "El tipo de documento modificado por la Nota debe ser una Factura y la serie debe comenzar con F.", vbInformation, "Subsane la observaci�n"
        cmdReferenceDocument.SetFocus
        Exit Function
    End If
    If Left(cboDocSerie, 1) = "B" And Mid(txtReferenceDocument, 4, 1) = "F" Then
        MsgBox "El tipo de documento modificado por la Nota debe ser una Boleta de Venta y la serie debe comenzar con B.", vbInformation, "Subsane la observaci�n"
        cmdReferenceDocument.SetFocus
        Exit Function
    End If

    ValidNoteFields = True
End Function

Private Sub SaveLastSerialNumber()
    If txtDocType = "01" Then
        sheetSetting.Range("O1") = cboDocSerie
    ElseIf txtDocType = "03" Then
        sheetSetting.Range("O2") = cboDocSerie
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
