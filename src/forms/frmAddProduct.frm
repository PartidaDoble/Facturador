VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddProduct 
   Caption         =   "AGREGAR PRODUCTO"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6120
   OleObjectBlob   =   "frmAddProduct.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmAddProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSearch_Click()
    Dim ProductRepository As New ProductRepositoryClass
    Dim Product As Variant
    Dim p As Long

    If Trim(txtProduct.Value) = Empty Then Exit Sub

    lstProducts.Clear

    For Each Product In ProductRepository.GetAll
        If InStr(UCase(Product.Description), UCase(Trim(txtProduct))) <> 0 Then
            With lstProducts
                If .ListCount >= 7 Then .Width = 295
                p = .ListCount
                .AddItem Product.Description & Space(200)
                .List(p, 1) = Format(Product.UnitPrice, "#,##0.00")
            End With
        End If
    Next Product
End Sub

Private Sub lstProducts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    txtDescription = Trim(lstProducts.Column(0))
    txtUnitPrice = lstProducts.Column(1)
    txtQuantity.SetFocus
End Sub

Private Sub txtQuantity_Change()
    lblTotal.Caption = Format(txtQuantity.Value * txtUnitPrice.Value, "#,##0.00")
End Sub

Private Sub cmdAdd_Click()
    Dim p As Long

    With frmInvoice.lstItems
        If .ListCount >= 8 Then frmInvoice.lstItems.Width = 497
        p = .ListCount
        .AddItem txtDescription & Space(200)
        .List(p, 1) = Format(txtQuantity, "#,##0.00")
        .List(p, 2) = txtUnitPrice
        .List(p, 3) = lblTotal
    End With
    
    FrmInvoiceShowInformation
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
