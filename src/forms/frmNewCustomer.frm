VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNewCustomer 
   Caption         =   "NUEVO CLIENTE"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6735
   OleObjectBlob   =   "frmNewCustomer.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmNewCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtAddress_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ToUppercase(KeyAscii)
End Sub

Private Sub UserForm_Initialize()
    cboDocType.List = Array("DNI", "RUC")
    cboDepartment.List = CollectionToArray(GetDepartments)
End Sub

Private Sub cboDocType_Change()
    If cboDocType = "DNI" Then cmdConsultSunat.Visible = False
    If cboDocType = "RUC" Then cmdConsultSunat.Visible = True
End Sub

Private Sub txtDocNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub cmdConsultSunat_Click()
    On Error Resume Next
    Dim CustomerInfo As Dictionary
    Dim Address As String
    Dim i As Integer
    Dim DistrictFullName As String
    Dim Location() As String
    Dim Departament As Variant
    
    If Not ThereIsInternet Then Exit Sub
    
    If Len(txtDocNumber) <> 11 Then
        MsgBox "El número de RUC debe tener 11 dígitos.", vbExclamation, "Subsane la observación"
        Exit Sub
    End If
    
    Set CustomerInfo = GetCustomerInfo(txtDocNumber)
    If CustomerInfo Is Nothing Then
        MsgBox "Hubo problemas al obtener la información desde la página web de la SUNAT.", vbCritical, ""
    Else
        txtName = CustomerInfo("name")
        
        If CustomerInfo("domicilio") <> Empty Then
            Address = CustomerInfo("domicilio")

            For Each Departament In GetDepartments
                If InStr(1, Address, Departament & " - ") > 0 Then
                    i = InStr(1, Address, Departament & " - ")
                    Exit For
                End If
            Next Departament
            
            txtAddress = Mid(Address, 1, i - 2)
            
            DistrictFullName = Mid(Address, i)
            If DistrictExists(DistrictFullName) Then
                Location = Split(DistrictFullName, " - ")
                cboDepartment = Location(0)
                cboProvince = Location(1)
                cboDistrict = Location(2)
            End If
        End If
    End If
End Sub

Private Sub txtName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ToUppercase(KeyAscii)
End Sub

Private Sub cboDepartment_Change()
    On Error Resume Next
    cboProvince.Clear
    cboProvince.List = CollectionToArray(GetProvinces(cboDepartment))
End Sub

Private Sub cboProvince_Change()
    On Error Resume Next
    cboDistrict.Clear
    cboDistrict.List = CollectionToArray(GetDistricts(cboProvince))
End Sub

Private Sub cmdSave_Click()
    Dim Index As Integer
    Dim Customer As New CustomerEntity
    Dim CustomerRepo As New CustomerRepository
    Dim DistrictFullName As String

    If Not ValidFields Then Exit Sub
    
    Customer.DocType = IIf(Trim(cboDocType) = "RUC", "6", "1")
    Customer.DocNumber = txtDocNumber
    Customer.Name = Trim(txtName)
    
    If CustomerRepo.Contains(Customer) Then
        MsgBox "El cliente con documento número " & txtDocNumber & " ya existe. " & _
               "No puede haber dos registros con el mismo número de documento.", vbExclamation, ""
        Exit Sub
    End If
    
    If (cboDepartment <> Empty And cboProvince <> Empty And cboDistrict <> Empty) Then
        DistrictFullName = Trim(cboDepartment) & " - " & Trim(cboProvince) & " - " & Trim(cboDistrict)
        
        If DistrictExists(DistrictFullName) Then
            Customer.Ubigeo = GetUbigeo(Trim(cboDepartment), Trim(cboProvince), Trim(cboDistrict))
            txtAddress = txtAddress & " (" & DistrictFullName & ")"
        End If
    End If
    
    Customer.Address = Trim(txtAddress)
    Customer.Email = LCase(Trim(txtEmail))
    
    CustomerRepo.Add Customer
    
    frmSearchCustomer.lstCustomers.Clear
    Index = frmSearchCustomer.lstCustomers.ListCount
    With frmSearchCustomer.lstCustomers
        .AddItem txtDocNumber
        .List(Index, 1) = Trim(txtName)
        .List(Index, 2) = IIf(Trim(cboDocType) = "RUC", "6", "1")
        .List(Index, 3) = Trim(txtAddress)
        .List(Index, 4) = Customer.Ubigeo
    End With
    
    MsgBox "Los datos del cliente se almacenaron correctamente.", vbInformation, "Datos almacenados"
    InfoLog "Los datos del cliente se almacenaron correctamente. DNI/RUC: " & txtDocNumber
    ThisWorkbook.Save
    Unload Me
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
    If Trim(txtAddress) <> Empty And (cboDepartment = Empty Or cboProvince = Empty Or cboDistrict = Empty) Then
        MsgBox "El domicilio del cliente es opcional. Si va a ingresar el domicilio, también debe ingresar el departamento, la provincia y el distrito.", vbExclamation, "Subsane la observación"
        cboDepartment.SetFocus
        Exit Function
    End If
    If (cboDepartment <> Empty Or cboProvince <> Empty Or cboDistrict <> Empty) And Trim(txtAddress) = Empty Then
        MsgBox "Ingrese el domicilio del cliente. Ademas debe seleccionar el departamento, la provincia y el distrito.", vbExclamation, "Subsane la observación"
        txtAddress.SetFocus
        Exit Function
    End If
    
    ValidFields = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub
