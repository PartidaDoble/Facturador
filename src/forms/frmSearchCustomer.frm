VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchCustomer 
   Caption         =   "AGREGAR CLIENTE"
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

Private Sub cmdSearch_Click()
    Dim CustomerRepository As New CustomerRepositoryClass
    Dim Customer As Variant
    Dim p As Long

    If Trim(txtSearchCustomer.Value) = Empty Then Exit Sub

    lstCustomers.Clear

    For Each Customer In CustomerRepository.GetAll
        If InStr(UCase(Customer.Name), UCase(Trim(txtSearchCustomer))) <> 0 Or InStr(UCase(Customer.DocNumber), UCase(Trim(txtSearchCustomer))) <> 0 Then
            With lstCustomers
                p = .ListCount
                .AddItem Customer.DocNumber
                .List(p, 1) = Customer.Name
            End With
        End If
    Next Customer
End Sub

Private Sub lstCustomers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    frmInvoice.txtCustomerDocNumber = Trim(lstCustomers.Column(0))
    frmInvoice.txtCustomerName = Trim(lstCustomers.Column(1))
    Unload Me
End Sub

Private Sub cmdAdd_Click()
    If lstCustomers.ListCount < 1 Then Exit Sub
    
    With lstCustomers
        frmInvoice.txtCustomerDocNumber = Trim(.List(.ListIndex, 0))
        frmInvoice.txtCustomerName = Trim(.List(.ListIndex, 1))
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
