VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNewProduct 
   Caption         =   "NUEVO PRODUCTO O SERVICIO"
   ClientHeight    =   2445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6150
   OleObjectBlob   =   "frmNewProduct.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmNewProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub txtCode_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyAlphanumeric(KeyAscii)
End Sub

Private Sub txtDescription_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ToUppercase(KeyAscii)
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    txtUnitPrice = Format(txtUnitPrice, "#,##0.00")
End Sub

Private Sub txtUnitPrice_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If InStr(txtUnitPrice, ".") > 0 And KeyAscii = Asc(".") Then KeyAscii = 0
    KeyAscii = OnlyAmount(KeyAscii)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
