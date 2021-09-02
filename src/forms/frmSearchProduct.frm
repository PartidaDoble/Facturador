VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchProduct 
   Caption         =   "AGREGAR PRODUCTO O SERVICIO"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6120
   OleObjectBlob   =   "frmSearchProduct.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSearchProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdShowFormNewProduct_Click()
    frmNewProduct.Show
End Sub

Private Sub txtSearchProduct_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ToUppercase(KeyAscii)
End Sub

Private Sub cmdSearch_Click()
    Dim ProductRepo As New ProductRepository
    Dim Product As Variant
    Dim Index As Long

    If Trim(txtSearchProduct.Value) = Empty Then Exit Sub

    lstProducts.Clear

    For Each Product In ProductRepo.GetAll
        If InStr(UCase(Product.Description), UCase(Trim(txtSearchProduct))) <> 0 Then
            With lstProducts
                If .ListCount >= 7 Then .Width = 295
                If .ListCount < 7 Then .Width = 282
                Index = .ListCount
                .AddItem Product.Description & Space(200)
                .List(Index, 1) = Format(Product.UnitPrice, "#,##0.00")
                .List(Index, 2) = Product.Code
                .List(Index, 3) = Product.UnitMeasure
            End With
        End If
    Next Product
End Sub

Private Sub lstProducts_Click()
    On Error Resume Next
    txtCode = lstProducts.Column(2)
    txtUnitMeasure = lstProducts.Column(3)
    txtDescription = Trim(lstProducts.Column(0))
    txtUnitPrice = lstProducts.Column(1)
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

Private Sub txtQuantity_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If InStr(txtQuantity, ".") > 0 And KeyAscii = Asc(".") Then KeyAscii = 0
    KeyAscii = OnlyAmount(KeyAscii)
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

Private Sub cmdAdd_Click()
    Dim p As Long
    
    If Not ValidFields Then Exit Sub

    With frmDocument.lstItems
        If .ListCount >= 8 Then frmDocument.lstItems.Width = 497
        p = .ListCount
        .AddItem txtDescription & Space(200)
        .List(p, 1) = Format(txtQuantity, "#,##0.00")
        .List(p, 2) = txtUnitPrice
        .List(p, 3) = txtTotal
        .List(p, 4) = txtCode
        .List(p, 5) = txtUnitMeasure
    End With
    FrmDocumentCalculateTotals
    
    Unload Me
End Sub

Private Function ValidFields() As Boolean
    If Trim(txtDescription) = Empty Then
        MsgBox "El producto debe tener una descripción.", vbExclamation, "Subsane la observación"
        txtDescription.SetFocus
        Exit Function
    End If
    If Trim(txtCode) = Empty Then
        MsgBox "El producto o servicio no tiene un Código. Revise la base de datos de productos.", vbExclamation, "Subsane la observación"
        Exit Function
    End If
    If Trim(txtUnitMeasure) = Empty Then
        MsgBox "El producto o servicio no tiene una Unidad de Medida. Revise la base de datos de productos.", vbExclamation, "Subsane la observación"
        Exit Function
    End If
    If Trim(txtQuantity) = Empty Or Not IsNumeric(Trim(txtQuantity)) Then
        MsgBox "Debe ingresar la cantidad de productos.", vbExclamation, "Subsane la observación"
        txtQuantity.SetFocus
        Exit Function
    End If
    If CDbl(Trim(txtQuantity)) <= 0 Then
        MsgBox "La cantidad de productos debe ser mayor a cero.", vbExclamation, "Subsane la observación"
        txtQuantity.SetFocus
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
    If Trim(txtTotal) = Empty Or Not IsNumeric(Trim(txtTotal)) Then
        MsgBox "El precio total presenta inconsistencias.", vbExclamation, "Subsane la observación"
        txtTotal.SetFocus
        Exit Function
    End If
    If CDbl(Trim(txtTotal)) <= 0 Then
        MsgBox "El precio total debe ser mayor a cero.", vbExclamation, "Subsane la observación"
        txtTotal.SetFocus
        Exit Function
    End If
    
    ValidFields = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    If Prop.App.Env = EnvProduction Then
        txtCode.Visible = False
        txtUnitMeasure.Visible = False
    End If
End Sub
