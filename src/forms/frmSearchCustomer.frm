VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchCustomer 
   Caption         =   "BUSCAR CLIENTE"
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6120
   OleObjectBlob   =   "frmSearchCustomer.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSearchCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdShowFormNewCustomer_Click()
    frmNewCustomer.Show
End Sub

Private Sub txtSearchCustomer_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ToUppercase(KeyAscii)
End Sub

Private Sub cmdSearch_Click()
    Dim CustomerRepo As New CustomerRepository
    Dim Customer As Variant
    Dim p As Long

    If Trim(txtSearchCustomer.Value) = Empty Then Exit Sub

    lstCustomers.Clear

    For Each Customer In CustomerRepo.GetAll
        If InStr(UCase(Customer.Name), UCase(Trim(txtSearchCustomer))) <> 0 Or InStr(UCase(Customer.DocNumber), UCase(Trim(txtSearchCustomer))) <> 0 Then
            With lstCustomers
                p = .ListCount
                .AddItem Customer.DocNumber
                .List(p, 1) = Customer.Name
                .List(p, 2) = Customer.DocType
                .List(p, 3) = Customer.Address
                .List(p, 4) = Customer.Ubigeo
            End With
        End If
    Next Customer
End Sub

Private Sub lstCustomers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    
    If Trim(lstCustomers.Column(0)) <> Empty Or Trim(lstCustomers.Column(1)) <> Empty Then
        frmDocument.txtCustomerDocNumber = Trim(lstCustomers.Column(0))
        frmDocument.txtCustomerName = Trim(lstCustomers.Column(1))
        frmDocument.txtCustomerDocType = Trim(lstCustomers.Column(2))
        frmDocument.txtCustomerAddress = Trim(lstCustomers.Column(3))
        frmDocument.txtCustomerUbigeo = Trim(lstCustomers.Column(4))
        
        Unload Me
    End If
End Sub

Private Sub cmdAdd_Click()
    If lstCustomers.ListCount < 1 Then Exit Sub
    
    If lstCustomers.ListIndex < 0 Then
        MsgBox "Debe seleccionar a un cliente. Otra forma es hacer doble click sobre un cliente.", vbExclamation, "Subsane la observación"
        Exit Sub
    End If
    
    With lstCustomers
        frmDocument.txtCustomerDocNumber = Trim(.List(.ListIndex, 0))
        frmDocument.txtCustomerName = Trim(.List(.ListIndex, 1))
        frmDocument.txtCustomerDocType = Trim(.List(.ListIndex, 2))
        frmDocument.txtCustomerAddress = Trim(.List(.ListIndex, 3))
        frmDocument.txtCustomerUbigeo = Trim(.List(.ListIndex, 4))
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
