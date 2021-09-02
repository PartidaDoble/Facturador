VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCancelDocument 
   Caption         =   "ANULAR COMPROBANTE DE PAGO"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5910
   OleObjectBlob   =   "frmCancelDocument.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCancelDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDocType_Change()
    On Error Resume Next
    Dim DocType As String
    
    DocType = GetDocTypeCode
    
    If DocType = "01" Then
        cboDocSerie.List = CollectionToArray(GetInvoiceSeries)
    ElseIf DocType = "03" Then
        cboDocSerie.List = CollectionToArray(GetBoletaSeries)
    ElseIf DocType = "07" Then
        cboDocSerie.List = CollectionToArray(GetCreditNoteSeries)
    ElseIf DocType = "08" Then
        cboDocSerie.List = CollectionToArray(GetDebitNoteSeries)
    End If
End Sub

Private Sub cmdShowDocument_Click()
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    Dim DocumentNumber As String
    
    If cboDocType = Empty Or cboDocSerie = Empty Or txtDocNumber = Empty Then
        MsgBox "Debe especificar el tipo, la serie y el n�mero del comprobante electr�nico.", vbInformation, "Subsane la observaci�n"
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

Private Sub UserForm_Initialize()
    cboDocType.List = Array("Factura", "Boleta de venta", "Nota de cr�dito", "Nota de d�bito")
End Sub

Private Sub txtDocNumber_Change()
    On Error Resume Next
    txtDocNumber = CInt(txtDocNumber)
End Sub

Private Sub txtDocNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub txtMotive_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ToUppercase(KeyAscii)
End Sub

Private Sub cmdAnular_Click()
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    Dim DocumentNumber As String
    Dim Answer As Integer
    
    If Not ValidFields Then Exit Sub
    
    DocumentNumber = GetDocTypeCode & "-" & cboDocSerie & "-" & Format(txtDocNumber, "00000000")
    Set Document = DocumentRepo.GetItem(DocumentNumber)
    
    If Not Document Is Nothing Then
        If Document.EmitedMoreSevenDaysAgo Then
            MsgBox "Tiene hasta 7 d�as para anular un comprobante electr�nico. El comprobante " & DocumentNumber & _
            " se emiti� hace mas de siete d�as. En este caso, debe emitir una Nota de Cr�dito", vbExclamation, "No cumple requisito"
            GoTo EndSub
        End If
        
        If Document.IsAccepted Then
            If Document.IsCanceled Then
                MsgBox "No puede anular el comprobante " & DocumentNumber & " porque ya se encuentra anulado.", vbExclamation, "Anulado"
                GoTo EndSub
            End If
        Else
            MsgBox "Para anular el comprobante " & DocumentNumber & ", previamente debe estar enviado y aceptado.", vbExclamation, "No cumple condici�n"
            GoTo EndSub
        End If
        
        Answer = MsgBox("�Est� seguro que desea anular el comprobante " & Document.Id & "?", vbYesNo + vbQuestion, "Anular comprobante")
        If Answer = vbYes Then
            Document.CancelInfo = "X|" & Trim(txtMotive)
            
            DocumentRepo.Update Document
            MsgBox "El comprobante " & Document.Id & " a sido registrado como anulado.", vbInformation, "Comprobante registrado"
            InfoLog "El comprobante " & Document.Id & " a sido registrado como anulado.", "frmCancelDocument.cmdAnular_Click"
            Unload Me
        End If
    Else
        MsgBox "El comprobante " & DocumentNumber & " no est� emitido y no se encuentra registrado en la hoja ""Comprobantes de Pago"".", vbExclamation, "Comprobante no emitido"
    End If
EndSub:
End Sub

Private Function ValidFields() As Boolean
    If Trim(cboDocType) = Empty Then
        MsgBox "Debe seleccionar el tipo de comprobante de pago.", vbExclamation, "Subsane la observaci�n"
        cboDocType.SetFocus
        Exit Function
    End If
    If Trim(cboDocSerie) = Empty Then
        MsgBox "Debe ingresar la serie del comprobante de pago.", vbExclamation, "Subsane la observaci�n"
        cboDocSerie.SetFocus
        Exit Function
    End If
    If Len(Trim(cboDocSerie)) <> 4 Then
        MsgBox "El n�mero de serie del comprobante debe tener 4 d�gitos.", vbInformation, "Subsane la observaci�n"
        cboDocSerie.SetFocus
        Exit Function
    End If
    If Trim(txtDocNumber) = Empty Then
        MsgBox "Debe ingresar el n�mero del comprobante de pago.", vbExclamation, "Subsane la observaci�n"
        txtDocNumber.SetFocus
        Exit Function
    End If
    If Trim(txtMotive) = Empty Then
        MsgBox "Debe ingresar el motivo por el cual est� eliminando el comprobante de pago.", vbExclamation, "Subsane la observaci�n"
        txtMotive.SetFocus
        Exit Function
    End If
    
    ValidFields = True
End Function

Private Function GetDocTypeCode() As String
    Select Case cboDocType
        Case "Factura"
            GetDocTypeCode = "01"
        Case "Boleta de venta"
            GetDocTypeCode = "03"
        Case "Nota de cr�dito"
            GetDocTypeCode = "07"
        Case "Nota de d�bito"
            GetDocTypeCode = "08"
    End Select
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub
