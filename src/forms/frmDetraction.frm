VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDetraction 
   Caption         =   "DETRACCIÓN"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9150.001
   OleObjectBlob   =   "frmDetraction.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmDetraction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Dim DetractionData As String
    
    If Not ValidFields Then Exit Sub
    
    DetractionData = Join(Array(Left(cboCode, 3), txtPercentage, Format(CDbl(txtAmount), "0.00"), Left(cboPaymentMethod, 3), txtCurrencySymbol), "-")
    
    frmDocument.txtDetractionData = DetractionData
    frmDocument.lblDetraction = "OPERACIÓN SUJETA AL SPOT " & txtCurrencySymbol & " " & txtAmount & " (" & txtPercentage & "%)"
    Unload Me
End Sub

Private Function ValidFields() As Boolean
    If Trim(Prop.Company.NroCtaDetraction) = Empty Then
        MsgBox "La cuenta de detracciones de la empresa no está configura.", vbExclamation, "Subsane la observación"
        Exit Function
    End If
    If Trim(cboCode) = Empty Then
        MsgBox "Seleccione un tipo de bien o servicio.", vbExclamation, "Subsane la observación"
        cboCode.SetFocus
        Exit Function
    End If
    If Val(txtPercentage) = 0 Then
        MsgBox "Ingrese el porcentaje de detracción.", vbExclamation, "Subsane la observación"
        txtPercentage.SetFocus
        Exit Function
    End If
    If Trim(txtAmount) = Empty Then
        MsgBox "Ingrese el importe de detracción.", vbExclamation, "Subsane la observación"
        txtAmount.SetFocus
        Exit Function
    End If
    If CDbl(txtAmount) <= 0 Then
        MsgBox "El importe de detracción debe ser un número mayor que cero.", vbExclamation, "Subsane la observación"
        txtAmount.SetFocus
        Exit Function
    End If
    If Trim(cboPaymentMethod) = Empty Then
        MsgBox "Seleccione un medio de pago.", vbExclamation, "Subsane la observación"
        cboPaymentMethod.SetFocus
        Exit Function
    End If
    
    ValidFields = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtAmount_AfterUpdate()
    txtAmount = Format(txtAmount, "#,##0.00")
End Sub

Private Sub txtAmount_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If InStr(txtAmount, ".") > 0 And KeyAscii = Asc(".") Then KeyAscii = 0
    KeyAscii = OnlyAmount(KeyAscii)
End Sub

Private Sub txtPercentage_Change()
    On Error GoTo HandleErrors
    Dim Result As String
    Result = Format((txtTotal * txtPercentage) / 100, "0")
    txtAmount = Format(CDbl(Result), "#,##0.00")
    Exit Sub
HandleErrors:
    txtAmount = Format(0, "#,##0.00")
End Sub

Private Sub txtPercentage_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub UserForm_Initialize()
    cboCode.List = CollectionToArray(GetProductsAndServices)
    cboPaymentMethod.List = CollectionToArray(GetPaymentsMethod)
    
    If Prop.App.Env = EnvProduction Then
        txtTotal.Visible = False
    End If
End Sub

Private Function GetProductsAndServices() As Collection
    Dim Data As New Collection
    
    Data.Add "001 Azúcar y melaza de caña"
    Data.Add "002 Arroz"
    Data.Add "003 Alcohol etílico"
    Data.Add "004 Recursos hidrobiológicos"
    Data.Add "005 Maíz amarillo duro"
    Data.Add "007 Caña de azúcar"
    Data.Add "008 Madera"
    Data.Add "009 Arena y piedra"
    Data.Add "010 Residuos, subproductos, desechos, recortes y desperdicios"
    Data.Add "011 Bienes gravados con el IGV, o renuncia a la exoneración"
    Data.Add "012 Intermediación laboral y tercerización"
    Data.Add "013 Animales vivos"
    Data.Add "014 Carnes y despojos comestibles"
    Data.Add "015 Abonos, cueros y pieles de origen animal"
    Data.Add "016 Aceite de pescado"
    Data.Add "017 Harina, polvo y pellets de pescado, crustáceos, moluscos"
    Data.Add "019 Arrendamiento de bienes muebles"
    Data.Add "020 Mantenimiento y reparación de bienes muebles"
    Data.Add "021 Movimiento de carga"
    Data.Add "022 Otros servicios empresariales"
    Data.Add "023 Leche"
    Data.Add "024 Comisión mercantil"
    Data.Add "025 Fabricación de bienes por encargo"
    Data.Add "026 Servicio de transporte de personas"
    Data.Add "027 Servicio de transporte de carga"
    Data.Add "028 Transporte de pasajeros"
    Data.Add "030 Contratos de construcción"
    Data.Add "031 Oro gravado con el IGV"
    Data.Add "034 Minerales metálicos no auríferos"
    Data.Add "035 Bienes exonerados del IGV"
    Data.Add "036 Oro y demás minerales metálicos exonerados del IGV"
    Data.Add "037 Demás servicios gravados con el IGV"
    Data.Add "039 Minerales no metálicos"
    Data.Add "040 Bien inmueble gravado con IGV"
    
    Set GetProductsAndServices = Data
End Function

Private Function GetPaymentsMethod() As Collection
    Dim Data As New Collection
    
    Data.Add "001 Depósito en cuenta"
    Data.Add "002 Giro"
    Data.Add "003 Transferencia de fondos"
    Data.Add "004 Orden de pago"
    Data.Add "005 Tarjeta de débito"
    Data.Add "006 Tarjeta de crédito emitida en el país por una empresa del sistema financiero"
    Data.Add "007 Cheques con la cláusula de NO NEGOCIABLE, INTRANSFERIBLES, NO A LA ORDEN"
    Data.Add "008 Efectivo, por operaciones en las que no existe obligación de utilizar medio de pago"
    Data.Add "009 Efectivo, en los demás casos"
    Data.Add "010 Medios de pago usados en comercio exterior"
    Data.Add "011 Documentos emitidos por las EDPYMES y las cooperativas de ahorro y crédito no..."
    Data.Add "012 Tarjeta de crédito emitida en el país o en el exterior por una empresa no..."
    Data.Add "013 Tarjetas de crédito emitidas en el exterior por empresas bancarias o..."
    Data.Add "999 Otros medios de pago"

    Set GetPaymentsMethod = Data
End Function
