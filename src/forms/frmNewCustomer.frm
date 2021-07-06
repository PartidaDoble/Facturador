VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNewCustomer 
   Caption         =   "NUEVO CLIENTE"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6240
   OleObjectBlob   =   "frmNewCustomer.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmNewCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    cboDocType.List = Array("DNI", "RUC")
End Sub

Private Sub txtDocNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub txtName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ToUppercase(KeyAscii)
End Sub

Private Sub cmdSave_Click()
    Dim Index As Integer
    
    frmSearchCustomer.lstCustomers.Clear
    
    Index = frmSearchCustomer.lstCustomers.ListCount
    With frmSearchCustomer.lstCustomers
        .AddItem txtDocNumber
        .List(Index, 1) = txtName
    End With
    
    MsgBox "Los datos del cliente se almacenaron con éxito", vbInformation, "Datos almacenados"

    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
