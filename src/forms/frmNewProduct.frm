VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNewProduct 
   Caption         =   "NUEVO PRODUCTO O SERVICIO"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6240
   OleObjectBlob   =   "frmNewProduct.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmNewProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    cboUnitMeasure.List = Array("UNIDAD", "KILOGRAMO", "LIBRA", "GRAMO", "CAJA", "GALON", "BARRIL", "LATA", "MILLAR", "METRO CUBICO", "METRO")
    cboUnitMeasure = "UNIDAD"
    txtDescription.SetFocus
    
    If Prop.App.AutoProdCode Then
        txtCode.Locked = True
        If WorksheetFunction.Max(sheetProducts.Columns(1)) = 0 Then
            txtCode = 10000
        Else
            txtCode = WorksheetFunction.Max(sheetProducts.Columns(1)) + 1
        End If
    Else
        txtCode.SetFocus
    End If
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

Private Sub cmdSave_Click()
    Dim Index As Integer
    Dim Product As New ProductEntity
    Dim ProductRepo As New ProductRepository
    
    If Not ValidFields Then Exit Sub
    
    frmSearchProduct.lstProducts.Clear
    
    Index = frmSearchProduct.lstProducts.ListCount
    With frmSearchProduct.lstProducts
        .AddItem Trim(txtDescription) & Space(200)
        .List(Index, 1) = txtUnitPrice
        .List(Index, 2) = Trim(txtCode)
        .List(Index, 3) = GetUnitMeasureFromName(Trim(cboUnitMeasure))
    End With
    
    Product.Code = Trim(txtCode)
    Product.UnitMeasure = GetUnitMeasureFromName(Trim(cboUnitMeasure))
    Product.Description = Trim(txtDescription)
    Product.UnitPrice = txtUnitPrice
    
    If ProductRepo.Contains(Product) Then
        MsgBox "El producto con código " & Trim(txtCode) & " ya existe. " & _
            "No puede haber dos registros con el mismo código.", vbExclamation, ""
    Else
        ProductRepo.Add Product
        MsgBox "Los datos del producto se almacenaron con éxito.", vbInformation, "Datos almacenados"
        InfoLog "Los datos del producto se almacenaron con éxito. Producto: " & Trim(txtDescription)
        ThisWorkbook.Save
        Unload Me
    End If
End Sub

Private Function ValidFields() As Boolean
    If Trim(txtCode) = Empty Then
        MsgBox "Debe ingresar el código del producto.", vbExclamation, "Subsane la observación"
        txtCode.SetFocus
        Exit Function
    End If
    If Trim(cboUnitMeasure) = Empty Then
        MsgBox "Seleccione una unidad de medida de la lista.", vbExclamation, "Subsane la observación"
        cboUnitMeasure.SetFocus
        Exit Function
    End If
    If Trim(txtDescription) = Empty Then
        MsgBox "Debe ingresar la descripción del producto o servicio.", vbExclamation, "Subsane la observación"
        txtDescription.SetFocus
        Exit Function
    End If
    If Trim(txtUnitPrice) = Empty Or Not IsNumeric(Trim(txtUnitPrice)) Then
        MsgBox "Debe ingresar el precio unitario del producto.", vbExclamation, "Subsane la observación"
        txtUnitPrice.SetFocus
        Exit Function
    End If
    If CDbl(Trim(txtUnitPrice)) <= 0 Then
        MsgBox "El precio unitario debe ser mayor a cero.", vbExclamation, "Subsane la observación"
        txtUnitPrice.SetFocus
        Exit Function
    End If
    
    ValidFields = True
End Function

Private Function GetUnitMeasureFromName(UnitMeasureName As String) As String
    Select Case UnitMeasureName
        Case "UNIDAD"
            GetUnitMeasureFromName = "NIU"
        Case "KILOGRAMO"
            GetUnitMeasureFromName = "KGM"
        Case "LIBRA"
            GetUnitMeasureFromName = "LBR"
        Case "GRAMO"
            GetUnitMeasureFromName = "GRM"
        Case "CAJA"
            GetUnitMeasureFromName = "BX"
        Case "GALON"
            GetUnitMeasureFromName = "GLL"
        Case "BARRIL"
            GetUnitMeasureFromName = "BLL"
        Case "LATA"
            GetUnitMeasureFromName = "CA"
        Case "MILLAR"
            GetUnitMeasureFromName = "MIL"
        Case "METRO CUBICO"
            GetUnitMeasureFromName = "MTQ"
        Case "METRO"
            GetUnitMeasureFromName = "MTR"
        Case Else
            GetUnitMeasureFromName = "NIU"
    End Select
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub
