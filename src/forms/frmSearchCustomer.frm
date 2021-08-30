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
        frmInvoice.txtCustomerDocNumber = Trim(lstCustomers.Column(0))
        frmInvoice.txtCustomerName = Trim(lstCustomers.Column(1))
        frmInvoice.txtCustomerDocType = Trim(lstCustomers.Column(2))
        frmInvoice.txtCustomerAddress = Trim(lstCustomers.Column(3))
        frmInvoice.txtCustomerUbigeo = Trim(lstCustomers.Column(4))
        
        frmNote.txtCustomerDocNumber = Trim(lstCustomers.Column(0))
        frmNote.txtCustomerName = Trim(lstCustomers.Column(1))
        frmNote.txtCustomerDocType = Trim(lstCustomers.Column(2))
        frmNote.txtCustomerAddress = Trim(lstCustomers.Column(3))
        frmNote.txtCustomerUbigeo = Trim(lstCustomers.Column(4))
        
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
        frmInvoice.txtCustomerDocNumber = Trim(.List(.ListIndex, 0))
        frmInvoice.txtCustomerName = Trim(.List(.ListIndex, 1))
        frmInvoice.txtCustomerDocType = Trim(.List(.ListIndex, 2))
        frmInvoice.txtCustomerAddress = Trim(.List(.ListIndex, 3))
        frmInvoice.txtCustomerUbigeo = Trim(.List(.ListIndex, 4))
    End With
    
    With lstCustomers
        frmNote.txtCustomerDocNumber = Trim(.List(.ListIndex, 0))
        frmNote.txtCustomerName = Trim(.List(.ListIndex, 1))
        frmNote.txtCustomerDocType = Trim(.List(.ListIndex, 2))
        frmNote.txtCustomerAddress = Trim(.List(.ListIndex, 3))
        frmNote.txtCustomerUbigeo = Trim(.List(.ListIndex, 4))
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
