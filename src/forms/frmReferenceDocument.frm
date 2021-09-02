VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReferenceDocument 
   Caption         =   "INFORMACI�N DEL DOCUMENTO QUE MODIFICA"
   ClientHeight    =   3405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5910
   OleObjectBlob   =   "frmReferenceDocument.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmReferenceDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Dim ReferenceDocumentData As String
    
    If Not ValidFields Then Exit Sub
    
    ReferenceDocumentData = Join(Array(GetDocTypeCode, cboDocSerie, txtDocNumber, GetMotiveCode, txtMotive), "-")
    
    frmDocument.txtReferenceDocument = ReferenceDocumentData
    Unload Me
End Sub

Private Function ValidFields() As Boolean
    If Trim(cboDocType) = Empty Then
        MsgBox "Debe especificar el tipo del comprobante que desea modificar.", vbInformation, "Subsane la observaci�n"
        cboDocType.SetFocus
        Exit Function
    End If
    If Len(Trim(cboDocSerie)) <> 4 Then
        MsgBox "Debe especificar la serie del comprobante que desea modificar.", vbInformation, "Subsane la observaci�n"
        cboDocSerie.SetFocus
        Exit Function
    End If
    If Trim(txtDocNumber) = Empty Then
        MsgBox "Debe especificar el n�mero del comprobante que desea modificar.", vbInformation, "Subsane la observaci�n"
        txtDocNumber.SetFocus
        Exit Function
    End If
    If Trim(cboMotive) = Empty Then
        MsgBox "Debe elegir un tipo de motivo.", vbInformation, "Subsane la observaci�n"
        cboMotive.SetFocus
        Exit Function
    End If
    If Trim(txtMotive) = Empty Then
        MsgBox "Debe especificar una descripci�n del motivo.", vbInformation, "Subsane la observaci�n"
        txtMotive.SetFocus
        Exit Function
    End If
    
    ValidFields = True
End Function

Private Sub cmdShowDocument_Click()
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    Dim DocumentNumber As String
    
    If cboDocType = Empty Or cboDocSerie = Empty Or txtDocNumber = Empty Then
        MsgBox "Debe especificar el tipo, la serie y el n�mero del documento que modifica.", vbInformation, "Subsane la observaci�n"
        cboDocType.SetFocus
        Exit Sub
    End If
    
    DocumentNumber = GetDocTypeCode & "-" & cboDocSerie & "-" & Format(txtDocNumber, "00000000")
    Set Document = DocumentRepo.GetItem(DocumentNumber)
    
    If Document Is Nothing Then
        MsgBox "El comprobante " & DocumentNumber & " no est� emitido y no se encuentra registrado en la hoja ""Comprobantes de Pago"".", vbInformation, "Comprobante no emitido"
    Else
        frmShowDocument.Caption = UCase(cboDocType)
        frmShowDocument.lblEmissionDate = Format(Document.Emission, "dd/mm/yyyy")
        frmShowDocument.lblDocument = Document.DocSerie & "-" & Format(Document.DocNumber, "00000000")
        frmShowDocument.lblCurrency = IIf(Document.TypeCurrency = "PEN", "Soles", "D�lares")
        frmShowDocument.lblCustomer = Document.Customer.Name
        frmShowDocument.lblTotal = Format(Document.Total, "#,##0.00")
        frmShowDocument.Show
    End If
End Sub

Private Function GetDocTypeCode() As String
    Select Case cboDocType
        Case "Factura"
            GetDocTypeCode = "01"
        Case "Boleta de venta"
            GetDocTypeCode = "03"
    End Select
End Function

Private Function GetMotiveCode() As String
    Select Case cboMotive
        Case "Anulaci�n de la operaci�n"
            GetMotiveCode = "01"
        Case "Anulaci�n por error en el RUC"
            GetMotiveCode = "02"
        Case "Correcci�n por error en la descripci�n"
            GetMotiveCode = "03"
        Case "Descuento global"
            GetMotiveCode = "04"
        Case "Descuento por �tem"
            GetMotiveCode = "05"
        Case "Devoluci�n total"
            GetMotiveCode = "06"
        Case "Devoluci�n por �tem"
            GetMotiveCode = "07"
        Case "Bonificaci�n"
            GetMotiveCode = "08"
        Case "Disminuci�n en el valor"
            GetMotiveCode = "09"
        Case "Otros Conceptos"
            GetMotiveCode = "10"
        
        ' Nota de d�bito
        Case "Intereses por mora"
            GetMotiveCode = "01"
        Case "Aumento en el valor"
            GetMotiveCode = "02"
        Case "Penalidades / otros conceptos"
            GetMotiveCode = "03"
    End Select
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cboDocType_Change()
    On Error Resume Next
    
    If GetDocTypeCode = "01" Then
        cboDocSerie.List = CollectionToArray(GetInvoiceSeries)
    ElseIf GetDocTypeCode = "03" Then
        cboDocSerie.List = CollectionToArray(GetBoletaSeries)
    End If
End Sub

Private Sub txtMotive_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ToUppercase(KeyAscii)
End Sub

Private Sub txtDocNumber_Change()
    On Error Resume Next
    txtDocNumber = CInt(txtDocNumber)
End Sub

Private Sub txtDocNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub UserForm_Initialize()
    cboDocType.List = Array("Factura", "Boleta de venta")
End Sub
