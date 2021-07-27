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

Private Sub cboDocType_Change()
    If cboDocType = "DNI" Then cmdConsultSunat.Visible = False
    If cboDocType = "RUC" Then cmdConsultSunat.Visible = True
End Sub

Private Sub txtDocNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub cmdConsultSunat_Click()
    Dim CustomerName As String
    
    If Len(txtDocNumber) <> 11 Then
        MsgBox "El número de RUC debe tener 11 dígitos.", vbExclamation, "Subsane la observación"
        Exit Sub
    End If
    
    CustomerName = GetCustomerName(txtDocNumber)
    If CustomerName = Empty Then
        MsgBox "Hubo problemas al obtener la información desde la página web de la SUNAT.", vbCritical, ""
    Else
        txtName = GetCustomerName(txtDocNumber)
    End If
End Sub

Private Sub txtName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ToUppercase(KeyAscii)
End Sub

Private Sub cmdSave_Click()
    Dim Index As Integer
    Dim Customer As New CustomerEntity
    Dim CustomerRepo As New CustomerRepository

    If Not ValidFields Then Exit Sub
    
    frmSearchCustomer.lstCustomers.Clear
    Index = frmSearchCustomer.lstCustomers.ListCount
    With frmSearchCustomer.lstCustomers
        .AddItem txtDocNumber
        .List(Index, 1) = txtName
    End With
    
    Customer.DocType = IIf(cboDocType = "RUC", "6", "1")
    Customer.DocNumber = txtDocNumber
    Customer.Name = Trim(txtName)
    
    If CustomerRepo.Contains(Customer) Then
        MsgBox "El cliente con documento número " & txtDocNumber & " ya existe. " & _
            "No puede haber dos registros con el mismo número de documento.", vbExclamation, ""
    Else
        CustomerRepo.Add Customer
        MsgBox "Los datos del cliente se almacenaron con éxito.", vbInformation, "Datos almacenados"
        InfoLog "Los datos del cliente se almacenaron con éxito. Número de documento: " & txtDocNumber
        Unload Me
    End If
End Sub

Private Function ValidFields() As Boolean
    If cboDocType = Empty Or (cboDocType <> "DNI" And cboDocType <> "RUC") Then
        MsgBox "Debe seleccionar el tipo de documento.", vbExclamation, "Subsane la observación"
        cboDocType.SetFocus
        Exit Function
    End If
    If txtDocNumber = Empty Then
        MsgBox "Debe ingresar el número de documento.", vbExclamation, "Subsane la observación"
        txtDocNumber.SetFocus
        Exit Function
    End If
    If cboDocType = "DNI" And Len(txtDocNumber) <> 8 Then
        MsgBox "El número de DNI debe tener 8 dígitos", vbExclamation, "Subsane la observación"
        txtDocNumber.SetFocus
        Exit Function
    End If
    If cboDocType = "RUC" And Len(txtDocNumber) <> 11 Then
        MsgBox "El número de RUC debe tener 11 dígitos", vbExclamation, "Subsane la observación"
        txtName.SetFocus
        Exit Function
    End If
    If Trim(txtName) = Empty Then
        MsgBox "Debe ingresar el nombre del cliente.", vbExclamation, "Subsane la observación"
        txtName.SetFocus
        Exit Function
    End If
    
    ValidFields = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub
