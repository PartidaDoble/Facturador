Attribute VB_Name = "Ribbon"
Option Explicit

Public Sub EmitInvoice(Control As IRibbonControl)
    If Not SfsPrepared Then Exit Sub
    
    frmDocument.Caption = "FACTURA"
    frmDocument.txtDocType = "01"
    frmDocument.cboDocSerie.List = CollectionToArray(GetInvoiceSeries)
    frmDocument.cboDocSerie = sheetSetting.Range("O1").Value
    frmDocument.txtDocNumber = NextCorrelativeNumber(sheetSetting.Range("O1"))
    frmDocument.lblCustomerDocType = "RUC:"
    frmDocument.cmdReferenceDocument.Visible = False
    
    If Trim(Prop.Company.NroCtaDetraction) = Empty Then
        frmDocument.cmdShowDetraction.Visible = False
    End If
    
    frmDocument.Show
End Sub

Public Sub EmitBoleta(Control As IRibbonControl)
    If Not SfsPrepared Then Exit Sub
    
    frmDocument.Caption = "BOLETA DE VENTA"
    frmDocument.txtDocType = "03"
    frmDocument.cboDocSerie.List = CollectionToArray(GetBoletaSeries)
    frmDocument.cboDocSerie = sheetSetting.Range("O2").Value
    frmDocument.txtDocNumber = NextCorrelativeNumber(sheetSetting.Range("O2"))
    frmDocument.lblCustomerDocType = "DNI:"
    frmDocument.cmdShowDetraction.Visible = False
    frmDocument.cmdReferenceDocument.Visible = False
    
    frmDocument.Show
End Sub

Public Sub EmitCreditNote(Control As IRibbonControl)
    If Not SfsPrepared Then Exit Sub
    
    frmDocument.Caption = "NOTA DE CRÉDITO"
    frmDocument.cboDocSerie.List = CollectionToArray(GetCreditNoteSeries)
    frmDocument.txtDocType = "07"
    frmDocument.cmdShowDetraction.Visible = False
    
    frmDocument.Show
End Sub

Public Sub EmitDebitNote(Control As IRibbonControl)
    If Not SfsPrepared Then Exit Sub
    
    frmDocument.Caption = "NOTA DE DÉBITO"
    frmDocument.cboDocSerie.List = CollectionToArray(GetDebitNoteSeries)
    frmDocument.txtDocType = "08"
    frmDocument.cmdShowDetraction.Visible = False
    
    frmDocument.Show
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
