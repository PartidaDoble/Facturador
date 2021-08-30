Attribute VB_Name = "FormsHelpers"
Option Explicit

Public Function NextCorrelativeNumber(Serie As String) As String
    NextCorrelativeNumber = WorksheetFunction.MaxIfs(sheetDocuments.Columns(5), sheetDocuments.Columns(4), Serie) + 1
End Function

Public Function GetInvoiceSeries() As Collection
    Set GetInvoiceSeries = GetSeries(6)
End Function

Public Function GetBoletaSeries() As Collection
    Set GetBoletaSeries = GetSeries(10)
End Function

Public Function GetCreditNoteSeries() As Collection
    Set GetCreditNoteSeries = JoinCollections(GetSeries(7), GetSeries(11))
End Function

Public Function GetDebitNoteSeries()
    Set GetDebitNoteSeries = JoinCollections(GetSeries(8), GetSeries(12))
End Function

Public Function JoinCollections(FirstColl As Collection, SecondColl As Collection) As Collection
    Dim JoinColl As New Collection
    Dim Item As Variant
    
    For Each Item In FirstColl
        JoinColl.Add Item
    Next Item
    
    For Each Item In SecondColl
        JoinColl.Add Item
    Next Item
    
    Set JoinCollections = JoinColl
End Function

Public Function GetSeries(Column As Integer) As Collection
    Dim Cell As Range
    Dim Series As New Collection

    For Each Cell In sheetSetting.Range(sheetSetting.Cells(3, Column), sheetSetting.Cells(12, Column))
        If Trim(Cell.Value) <> Empty Then
            Series.Add Trim(Cell.Value)
        End If
    Next Cell
    
    Set GetSeries = Series
End Function

Public Function GetUbigeo(Department As String, Province As String, District As String) As String
    Dim Value As String
    Value = Department & " - " & Province & " - " & District
    GetUbigeo = WorksheetFunction.VLookup(Value, sheetUbigeo.Columns("D:E"), 2, 0)
End Function

Public Function DistrictExists(DistrictFullName As String) As Boolean
    DistrictExists = GetMatchRow(sheetUbigeo, 4, DistrictFullName) <> 0
End Function

Public Function GetDepartments() As Collection
    Set GetDepartments = FilterUnique(LoadColumnData(sheetUbigeo, 1, 2, GetLastRow(sheetUbigeo)))
End Function

Public Function GetProvinces(Department As String) As Collection
    Dim StartRow As Long
    Dim EndRow As Long
    
    StartRow = GetMatchRow(sheetUbigeo, 1, Department)
    EndRow = GetRowWithDifferentValue(sheetUbigeo, StartRow, 1, Department) - 1
    Set GetProvinces = FilterUnique(LoadColumnData(sheetUbigeo, 2, StartRow, EndRow))
End Function

Public Function GetDistricts(Province As String) As Collection
    Dim StartRow As Long
    Dim EndRow As Long
    Dim ControlRow As Long
    
    StartRow = GetMatchRow(sheetUbigeo, 2, Province)
    EndRow = GetRowWithDifferentValue(sheetUbigeo, StartRow, 2, Province) - 1
    Set GetDistricts = FilterUnique(LoadColumnData(sheetUbigeo, 3, StartRow, EndRow))
End Function

Private Function GetRowWithDifferentValue(Sheet As Worksheet, StartRow As Long, Column As Long, Value As String) As Long
    Dim Row As Long

    Row = StartRow
    Do While True
        If Sheet.Cells(Row, Column) <> Value Then
            GetRowWithDifferentValue = Row
            Exit Do
        End If
        Row = Row + 1
    Loop
End Function

Public Function ToUppercase(InputKey) As Variant
    ToUppercase = Asc(UCase(Chr(InputKey)))
End Function

Public Function OnlyAmount(KeyAscci)
    Dim Keys As String
    Keys = "1234567890." & Chr(vbKeyBack)
    OnlyAmount = IIf(InStr(Keys, Chr(KeyAscci)) = 0, 0, KeyAscci)
End Function

Function OnlyInteger(InputKey) As Variant
    Dim Keys As String
    Keys = "1234567890" & Chr(vbKeyBack)
    OnlyInteger = IIf(InStr(Keys, Chr(InputKey)) = 0, 0, InputKey)
End Function

Public Function OnlyAlphanumeric(KeyAscii)
    OnlyAlphanumeric = IIf(IsAlphanumeric(KeyAscii), Asc(UCase(Chr(KeyAscii))), 0)
End Function

Private Function IsAlphanumeric(KeyAscii As Variant) As Boolean
    Dim KeyIsNumber As Boolean
    Dim KeyIsLowercase As Boolean
    Dim KeyIsUppercase As Boolean
    
    KeyIsNumber = 48 <= KeyAscii And KeyAscii <= 57
    KeyIsLowercase = 97 <= KeyAscii And KeyAscii <= 122
    KeyIsUppercase = 65 <= KeyAscii And KeyAscii <= 90
    
    IsAlphanumeric = KeyIsNumber Or KeyIsLowercase Or KeyIsUppercase
End Function

Public Sub FrmInvoiceCalculateTotals()
    Dim TypeCurrency As String
    Dim TotalPrice As Double
    Dim SubTotal As Double
    Dim Igv As Double

    TotalPrice = FrmInvoiceSumTotalItems
    SubTotal = TotalPrice / (Prop.Rate.Igv + 1)
    Igv = TotalPrice - SubTotal

    frmInvoice.lblSubTotal = Format(SubTotal, "#,##0.00")
    frmInvoice.lblIGV = Format(Igv, "#,##0.00")
    frmInvoice.lblTotal = Format(TotalPrice, "#,##0.00")

    TypeCurrency = IIf(frmInvoice.cboTypeCurrency = "Soles", "PEN", "USD")
    frmInvoice.lblTotalInLetters.Caption = "SON: " & AmountInLetters(TotalPrice, TypeCurrency)
End Sub

Public Function FrmInvoiceSumTotalItems() As Double
    Dim Sum As Double
    Dim Index As Integer

    With frmInvoice.lstItems
        If .ListCount < 1 Then
            Sum = 0
        Else
            For Index = 0 To .ListCount - 1
                Sum = Sum + .List(Index, 3)
            Next Index
        End If
    End With

    FrmInvoiceSumTotalItems = Sum
End Function

Public Sub FrmNoteCalculateTotals()
    Dim TypeCurrency As String
    Dim TotalPrice As Double
    Dim SubTotal As Double
    Dim Igv As Double

    TotalPrice = FrmNoteSumTotalItems
    SubTotal = TotalPrice / (Prop.Rate.Igv + 1)
    Igv = TotalPrice - SubTotal

    frmNote.lblSubTotal = Format(SubTotal, "#,##0.00")
    frmNote.lblIGV = Format(Igv, "#,##0.00")
    frmNote.lblTotal = Format(TotalPrice, "#,##0.00")

    TypeCurrency = IIf(frmNote.cboTypeCurrency = "Soles", "PEN", "USD")
    frmNote.lblTotalInLetters.Caption = "SON: " & AmountInLetters(TotalPrice, TypeCurrency)
End Sub

Public Function FrmNoteSumTotalItems() As Double
    Dim Sum As Double
    Dim Index As Integer

    With frmNote.lstItems
        If .ListCount < 1 Then
            Sum = 0
        Else
            For Index = 0 To .ListCount - 1
                Sum = Sum + .List(Index, 3)
            Next Index
        End If
    End With

    FrmNoteSumTotalItems = Sum
End Function

Public Function GetCustomerInfo(Ruc As String) As Dictionary
    On Error GoTo HandleErrors
    Dim Doc As New Scraping
    Dim Name As String
    Dim Address As String
    Dim Info As New Dictionary
    Dim i As Integer
    
    Doc.gotoPage "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp"
    Doc.Id("txtRuc").FieldValue Ruc
    Doc.Id("btnAceptar").click sleep:=3
    
    For i = 0 To Doc.css(".list-group .col-sm-5").Count - 1
        If Trim(Doc.css(".list-group .col-sm-5").Index(i).text) = "Número de RUC:" Then
            Name = Trim(Doc.css(".list-group .col-sm-7").Index(i).text)
            Info.Add "name", Trim(Mid(Name, InStr(Name, "-") + 1))
        End If
        If Trim(Doc.css(".list-group .col-sm-5").Index(i).text) = "Domicilio Fiscal:" Then
            Address = Trim(Doc.css(".list-group .col-sm-7").Index(i).text)
            Info.Add "domicilio", IIf(Address = "-", Empty, Address)
        End If
        If Trim(Doc.css(".list-group .col-sm-5").Index(i).text) = "Estado del Contribuyente:" Then
            Info.Add "estado", Trim(Doc.css(".list-group .col-sm-7").Index(i).text)
        End If
        If Trim(Doc.css(".list-group .col-sm-5").Index(i).text) = "Condición del Contribuyente:" Then
            Info.Add "condicion", Trim(Doc.css(".list-group .col-sm-7").Index(i).text)
        End If
    Next i
    
    Set GetCustomerInfo = Info
    DebugLog "Se obtiene el nombre del cliente desde la página web de la SUNAT | RUC: " & Ruc
    Exit Function
HandleErrors:
    DebugLog "Hubo problemas al obtener la información desde la página web de la SUNAT | RUC: " & Ruc, "GetCustomerInfo"
    Set GetCustomerInfo = Nothing
End Function
