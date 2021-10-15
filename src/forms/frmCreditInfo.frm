VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreditInfo 
   Caption         =   "INFORMACIÓN DEL CRÉDITO"
   ClientHeight    =   3285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmCreditInfo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCreditInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Dim Installment1 As String
    Dim Installment2 As String
    Dim Installment3 As String
    
    If Not ValidFields Then Exit Sub
    
    frmDocument.txtPaymentData = Join(Array("Credito", Format(CDbl(txtNetAmountPending), "0.00"), txtTypeCurrency), "-")
    frmDocument.lblNetAmountPending = IIf(txtTypeCurrency = "PEN", "S/ ", "US$ ") & txtNetAmountPending
    
    If txtAmount1 <> Empty And txtPaymentDate1 <> Empty Then
        Installment1 = Join(Array(Format(CDbl(txtAmount1), "0.00"), txtPaymentDate1, txtTypeCurrency), "-")
        frmDocument.lblInstallment1.Visible = True
        frmDocument.lblInstallmentDate1 = txtPaymentDate1
        frmDocument.lblInstallmentAmount1 = IIf(txtTypeCurrency = "PEN", "S/ ", "US$ ") & txtAmount1
    End If
    If txtAmount2 <> Empty And txtPaymentDate2 <> Empty Then
        Installment2 = Join(Array(Format(CDbl(txtAmount2), "0.00"), txtPaymentDate2, txtTypeCurrency), "-")
        frmDocument.lblInstallment2.Visible = True
        frmDocument.lblInstallmentDate2 = txtPaymentDate2
        frmDocument.lblInstallmentAmount2 = IIf(txtTypeCurrency = "PEN", "S/ ", "US$ ") & txtAmount2
    End If
    If txtAmount3 <> Empty And txtPaymentDate3 <> Empty Then
        Installment3 = Join(Array(Format(CDbl(txtAmount3), "0.00"), txtPaymentDate3, txtTypeCurrency), "-")
        frmDocument.lblInstallment3.Visible = True
        frmDocument.lblInstallmentDate3 = txtPaymentDate3
        frmDocument.lblInstallmentAmount3 = IIf(txtTypeCurrency = "PEN", "S/ ", "US$ ") & txtAmount3
    End If
    
    frmDocument.txtPaymentDetail = Join(Array(Installment1, Installment2, Installment3), "|")
    
    frmDocument.fmeWayPayDetail.Top = 324
    frmDocument.Height = 463
    frmDocument.cmdSave.Top = 402
    frmDocument.cmdCancel.Top = 402
    frmDocument.fmeWayPayDetail.Visible = True
    Unload Me
End Sub

Private Function ValidFields() As Boolean
    Dim NetAmountPending As Long
    Dim Amount1 As Long
    Dim Amount2 As Long
    Dim Amount3 As Long
    
    If Not IsValidDate(txtPaymentDate1) Then
        MsgBox "Ingrese una fecha de emisión válida. El formato de fecha es dd/mm/yyyy.", vbExclamation, "Subsane la observación"
        txtPaymentDate1.SetFocus
        Exit Function
    End If
    If CDate(txtEmissionDate) >= CDate(txtPaymentDate1) Then
        MsgBox "La fecha de pago de la cuota debe ser posterior a la fecha de emisión de la factura.", vbExclamation, "Subsane la observación"
        txtPaymentDate1.SetFocus
        Exit Function
    End If
    If Trim(txtAmount1) = Empty Then
        MsgBox "Ingrese el monto de la primera cuota.", vbExclamation, "Subsane la observación"
        txtAmount1.SetFocus
        Exit Function
    End If
    
    If txtPaymentDate2 <> Empty And txtAmount2 <> Empty Then
        If Not IsValidDate(txtPaymentDate2) Then
            MsgBox "Ingrese una fecha de emisión válida. El formato de fecha es dd/mm/yyyy.", vbExclamation, "Subsane la observación"
            txtPaymentDate2.SetFocus
            Exit Function
        End If
        If CDate(txtPaymentDate1) >= CDate(txtPaymentDate2) Then
            MsgBox "La fecha de pago de la cuota 2 debe ser posterior a la fecha de pago de la couta 1.", vbExclamation, "Subsane la observación"
            txtPaymentDate2.SetFocus
            Exit Function
        End If
    End If
    
    If txtPaymentDate3 <> Empty And txtAmount3 <> Empty Then
        If Not IsValidDate(txtPaymentDate3) Then
            MsgBox "Ingrese una fecha de emisión válida. El formato de fecha es dd/mm/yyyy.", vbExclamation, "Subsane la observación"
            txtPaymentDate3.SetFocus
            Exit Function
        End If
        If CDate(txtPaymentDate2) >= CDate(txtPaymentDate3) Then
            MsgBox "La fecha de pago de la cuota 3 debe ser posterior a la fecha de pago de la couta 2.", vbExclamation, "Subsane la observación"
            txtPaymentDate3.SetFocus
            Exit Function
        End If
    End If
    
    NetAmountPending = CLng(CDbl(txtNetAmountPending) * 100)
    If IsNumeric(txtAmount1) Then
        Amount1 = CLng(CDbl(txtAmount1) * 100)
    End If
    If IsNumeric(txtAmount2) Then
        Amount2 = CLng(CDbl(txtAmount2) * 100)
    End If
    If IsNumeric(txtAmount3) Then
        Amount3 = CLng(CDbl(txtAmount3) * 100)
    End If
    
    If NetAmountPending <> Amount1 + Amount2 + Amount3 Then
        MsgBox "La suma de las cuotas debe ser igual al monto neto pendiente de pago.", vbExclamation, "Subsane la observación"
        Exit Function
    End If
    
    ValidFields = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtAmount1_AfterUpdate()
    txtAmount1 = Format(txtAmount1, "#,##0.00")
End Sub

Private Sub txtAmount1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If InStr(txtAmount1, ".") > 0 And KeyAscii = Asc(".") Then KeyAscii = 0
    KeyAscii = OnlyAmount(KeyAscii)
End Sub

Private Sub txtAmount2_AfterUpdate()
    txtAmount2 = Format(txtAmount2, "#,##0.00")
End Sub

Private Sub txtAmount2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If InStr(txtAmount2, ".") > 0 And KeyAscii = Asc(".") Then KeyAscii = 0
    KeyAscii = OnlyAmount(KeyAscii)
End Sub

Private Sub txtAmount3_AfterUpdate()
    txtAmount3 = Format(txtAmount3, "#,##0.00")
End Sub

Private Sub txtAmount3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If InStr(txtAmount3, ".") > 0 And KeyAscii = Asc(".") Then KeyAscii = 0
    KeyAscii = OnlyAmount(KeyAscii)
End Sub

Private Sub txtNetAmountPending_AfterUpdate()
    txtNetAmountPending = Format(txtNetAmountPending, "#,##0.00")
End Sub

Private Sub txtNetAmountPending_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If InStr(txtNetAmountPending, ".") > 0 And KeyAscii = Asc(".") Then KeyAscii = 0
    KeyAscii = OnlyAmount(KeyAscii)
End Sub
