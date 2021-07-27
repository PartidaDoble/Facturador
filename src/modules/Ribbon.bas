Attribute VB_Name = "Ribbon"
Option Explicit

Public Sub EmitirBoleta(Control As IRibbonControl)
    frmInvoice.Caption = "BOLETA DE VENTA"
    frmInvoice.cboDocSerie.List = Array("B001", "B002")
    frmInvoice.txtDocType = "03"
    frmInvoice.lblCustomerDocType = "DNI:"
    frmInvoice.txtCustomerDocType = "1"
    frmInvoice.Show
End Sub

Public Sub EmitirFactura(Control As IRibbonControl)
    frmInvoice.Caption = "FACTURA"
    frmInvoice.cboDocSerie.List = Array("F001", "F002")
    frmInvoice.txtDocType = "01"
    frmInvoice.lblCustomerDocType = "RUC:"
    frmInvoice.txtCustomerDocType = "6"
    frmInvoice.Show
End Sub
