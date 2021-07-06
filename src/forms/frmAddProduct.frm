VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddProduct 
   Caption         =   "AGREGAR PRODUCTO O SERVICIO"
   ClientHeight    =   5340
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

Private Sub cmdSearch_Click()
    Dim ProductRepository As New ProductRepositoryClass
    Dim Product As Variant
    Dim Index As Long

    If Trim(txtSearchProduct.Value) = Empty Then Exit Sub

    lstProducts.Clear

    For Each Product In ProductRepository.GetAll
        If InStr(UCase(Product.Description), UCase(Trim(txtSearchProduct))) <> 0 Then
            With lstProducts
                If .ListCount >= 7 Then .Width = 295
                Index = .ListCount
                .AddItem Product.Description & Space(200)
                .List(Index, 1) = Format(Product.UnitPrice, "#,##0.00")
                .List(Index, 2) = Product.Code
                .List(Index, 3) = Product.UnitMeasurement
            End With
        End If
    Next Product
End Sub

Private Sub cmdShowFormNewProduct_Click()
    frmNewProduct.Show
End Sub

Private Sub lstProducts_Click()
    On Error Resume Next
    txtDescription = Trim(lstProducts.Column(0))
    txtUnitPrice = lstProducts.Column(1)
    txtCode = lstProducts.Column(2)
    txtUnitMeasurement = lstProducts.Column(3)
    txtQuantity.SetFocus
End Sub

Private Sub lstProducts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    txtDescription = Trim(lstProducts.Column(0))
    txtUnitPrice = lstProducts.Column(1)
    txtCode = lstProducts.Column(2)
    txtUnitMeasurement = lstProducts.Column(3)
    txtQuantity.SetFocus
End Sub

Private Sub txtDescription_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ToUppercase(KeyAscii)
End Sub

Private Sub txtQuantity_AfterUpdate()
    txtQuantity = Format(txtQuantity, "#,##0.00")
End Sub

Private Sub txtQuantity_Change()
    On Error Resume Next
    txtTotal = Format(txtQuantity.Value * txtUnitPrice.Value, "#,##0.00")
End Sub

Private Sub cmdAdd_Click()
    Dim p As Long

    With frmInvoice.lstItems
        If .ListCount >= 8 Then frmInvoice.lstItems.Width = 497
        p = .ListCount
        .AddItem txtDescription & Space(200)
        .List(p, 1) = Format(txtQuantity, "#,##0.00")
        .List(p, 2) = txtUnitPrice
        .List(p, 3) = txtTotal
        .List(p, 4) = txtCode
        .List(p, 5) = txtUnitMeasurement
    End With
    
    FrmInvoiceShowInformation
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtQuantity_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If InStr(txtQuantity, ".") > 0 And KeyAscii = Asc(".") Then KeyAscii = 0
    KeyAscii = OnlyAmount(KeyAscii)
End Sub

Private Sub txtSearchProduct_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ToUppercase(KeyAscii)
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    txtUnitPrice = Format(txtUnitPrice, "#,##0.00")
End Sub

Private Sub txtUnitPrice_Change()
    On Error Resume Next
    txtTotal = Format(txtQuantity * txtUnitPrice, "#,##0.00")
End Sub

Private Sub txtUnitPrice_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If InStr(txtUnitPrice, ".") > 0 And KeyAscii = Asc(".") Then KeyAscii = 0
    KeyAscii = OnlyAmount(KeyAscii)
End Sub
