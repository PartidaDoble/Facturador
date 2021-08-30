Attribute VB_Name = "Ribbon"
Option Explicit

Public Sub EmitInvoice(Control As IRibbonControl)
    If Not SfsPrepared Then Exit Sub
    
    frmInvoice.Caption = "FACTURA"
    frmInvoice.txtDocType = "01"
    frmInvoice.cboDocSerie.List = CollectionToArray(GetInvoiceSeries)
    frmInvoice.cboDocSerie = sheetSetting.Range("O1").Value
    frmInvoice.txtDocNumber = NextCorrelativeNumber(sheetSetting.Range("O1"))
    frmInvoice.lblCustomerDocType = "RUC:"
    
    If Trim(Prop.Company.NroCtaDetraction) = Empty Then
        frmInvoice.cmdShowDetraction.Visible = False
    End If
    
    frmInvoice.Show
End Sub

Public Sub EmitBoleta(Control As IRibbonControl)
    If Not SfsPrepared Then Exit Sub
    
    frmInvoice.Caption = "BOLETA DE VENTA"
    frmInvoice.txtDocType = "03"
    frmInvoice.cboDocSerie.List = CollectionToArray(GetBoletaSeries)
    frmInvoice.cboDocSerie = sheetSetting.Range("O2").Value
    frmInvoice.txtDocNumber = NextCorrelativeNumber(sheetSetting.Range("O2"))
    frmInvoice.lblCustomerDocType = "DNI:"
    frmInvoice.cmdShowDetraction.Visible = False
    frmInvoice.Show
End Sub

Public Sub EmitCreditNote(Control As IRibbonControl)
    If Not SfsPrepared Then Exit Sub
    
    frmNote.Caption = "NOTA DE CRÉDITO"
    frmNote.cboDocSerie.List = CollectionToArray(GetCreditNoteSeries)
    frmNote.txtDocType = "07"
    frmNote.cboMotive.List = Array("Anulación de la operación", "Anulación por error en el RUC", "Corrección por error en la descripción", "Descuento global", "Descuento por ítem", "Devolución total", "Devolución por ítem", "Bonificación", "Disminución en el valor", "Otros Conceptos")
    frmNote.Show
End Sub

Public Sub EmitDebitNote(Control As IRibbonControl)
    If Not SfsPrepared Then Exit Sub
    
    frmNote.Caption = "NOTA DE DÉBITO"
    frmNote.cboDocSerie.List = CollectionToArray(GetDebitNoteSeries)
    frmNote.txtDocType = "08"
    frmNote.cboMotive.List = Array("Intereses por mora", "Aumento en el valor", "Penalidades / otros conceptos")
    frmNote.Show
End Sub

Public Sub CancelDocument(Control As IRibbonControl)
    frmCancelDocument.Show
End Sub

Public Sub SendInvoicesAndNotes(Control As IRibbonControl)
    On Error GoTo HandleErrors
    
    If Not ThereIsInternet Then Exit Sub
    If Not SfsPrepared Then Exit Sub
    
    Application.StatusBar = "Enviando facturas y notas electrónicas..."
    SendGeneratedInvoicesAndNotes
    SaveSentInvoicesAndNotes
    
    Application.StatusBar = "Enviando facturas y notas anuladas..."
    SendCanceledInvoicesAndNotes
    
    Application.StatusBar = Empty
    ThisWorkbook.Save
    MsgBox "El envío de Facturas y Notas vinculadas a terminado.", vbInformation, "Operación terminada"
    InfoLog "El envío de Facturas y Notas vinculadas a terminado.", "SendInvoicesAndNotes"
    Exit Sub
HandleErrors:
    Application.StatusBar = Empty
    MsgBox "Error al enviar Facturas y Notas electrónicas.", vbCritical, "ERROR"
    ErrorLog "Error al enviar Facturas y Notas electrónicas", "SendInvoicesAndNotes", Err.Number
End Sub

Public Sub SendBoletasAndNotes(Control As IRibbonControl)
    On Error GoTo HandleErrors
    Dim Answer As Integer
    
    If Not ThereIsInternet Then Exit Sub
    If Not SfsPrepared Then Exit Sub
    
    Answer = MsgBox("Las Boletas de Venta y Notas vinculadas se envían en grupos de hasta 500 comprobantes. " & _
                    "Es recomendable hacer el envío pocas veces al día, de preferencia una sola vez." & Chr(13) & Chr(13) & _
                    "¿Está seguro que desea continuar?", vbYesNo + vbQuestion, "Enviar Boletas y Notas vinculadas")
    If Answer = vbNo Then Exit Sub
    
    Application.StatusBar = "Enviando Boletas de Venta y Notas electrónicas..."
    SendGeneratedBoletasAndNotesLoop
    
    Application.StatusBar = Empty
    ThisWorkbook.Save
    MsgBox "El envío de Boletas de Venta y Notas vinculadas a terminado.", vbInformation, "Operación terminada"
    InfoLog "El envío de Boletas de Venta y Notas vinculadas a terminado.", "SendBoletasAndNotes"
    Exit Sub
HandleErrors:
    Application.StatusBar = Empty
    MsgBox "Error al enviar Boletas de Venta y Notas electrónicas.", vbCritical, "ERROR"
    ErrorLog "Error al enviar Boletas de Venta y Notas electrónicas", "SendBoletasAndNotes", Err.Number
End Sub

Public Sub CheckTickets(Control As IRibbonControl)
    On Error GoTo HandleErrors
    If Not ThereIsInternet Then Exit Sub
    If Not SfsPrepared Then Exit Sub
    
    Application.StatusBar = "Consultando tickets de resumenes diarios de boletas y notas..."
    UpdateStatusDailySummary
    SaveSentBoletasAndNotes
    
    Application.StatusBar = "Consultando tickets de comprobantes anulados..."
    UpdateStatusCanceledInvoicesAndNotes
    SaveSentCanceledInvoicesAndNotes
    
    Application.StatusBar = Empty
    ThisWorkbook.Save
    MsgBox "La consulta de tickets a terminado.", vbInformation, "Operación terminada"
    InfoLog "La consulta de tickets a terminado.", "ConsultarTickets"
    Exit Sub
HandleErrors:
    Application.StatusBar = Empty
    MsgBox "Error al consultar los tickets.", vbCritical, "ERROR"
    ErrorLog "Error al consultar los tickets.", "ConsultarTickets", Err.Number
End Sub

Public Sub SendEmails(Control As IRibbonControl)
    On Error GoTo HandleErrors
    Dim Answer As Integer
    
    If Not ThereIsInternet Then Exit Sub
    
    If Not Prop.App.Premium Then
        MsgBox "Esta funcionalidad no está disponible en la versión libre. ", vbInformation, "No disponible"
        Exit Sub
    End If
    
    Answer = MsgBox("Se procederá al envío de todas las Facturas y Notas vinculadas con situación ""enviado y aceptado sunat"", " & _
                    "que aún no hayan sido envidas al cliente." & Chr(13) & Chr(13) & _
                    "¿Está seguro que desea continuar?", vbYesNo + vbQuestion, "Enviar correos electrónicos")
    If Answer = vbNo Then Exit Sub
    
    If Prop.Email.Provider = GmailProv Then
        Answer = MsgBox("Dado que está usando Gmail como proveedor de correo electrónico, " & _
                    "esta operación puede demorar entre 4 a 8 segundos por correo enviado. " & _
                    "No realice ninguna otra tarea en la aplicación mientras no termine la operación." & Chr(13) & Chr(13) & _
                    "¿Está seguro que desea continuar?", vbYesNo + vbQuestion, "Enviar correos electrónicos")
        If Answer = vbNo Then Exit Sub
    End If
    
    Application.StatusBar = "Enviando correos electrónicos..."
    Run "SendMassEmails"
    
    Application.StatusBar = Empty
    ThisWorkbook.Save
    MsgBox "El envío de correos electrónicos a terminado.", vbInformation, "Operación terminada"
    InfoLog "El envío de correos electrónicos a terminado.", "SendEmails"
    Exit Sub
HandleErrors:
    Application.StatusBar = Empty
    MsgBox "Error al enviar correos electrónicos.", vbCritical, "ERROR"
    ErrorLog "Error al enviar correos electrónicos.", "SendEmails", Err.Number
End Sub

Public Sub NewClient(Control As IRibbonControl)
    frmNewCustomer.Show
End Sub

Public Sub NewProduct(Control As IRibbonControl)
    frmNewProduct.Show
End Sub
