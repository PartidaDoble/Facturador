Attribute VB_Name = "Utils"
Option Explicit

Public Function IsValidDate(DateStr As String) As Boolean
    On Error GoTo HandleErrors
    If IIf(IsDate(DateStr), Format(CDate(DateStr), "dd/mm/yyyy"), "") = DateStr Then
        IsValidDate = True
    End If
HandleErrors:
End Function

Public Sub MaximizeRibbon()
    On Error Resume Next
    If CommandBars("Ribbon").Height < 100 Then CommandBars.ExecuteMso "MinimizeRibbon"
End Sub

Public Function ThereIsInternet() As Boolean
    On Error GoTo HandleErrors
    Dim XMLPage As New MSXML2.XMLHTTP60
    
    XMLPage.Open "HEAD", "https://www.google.com/", False
    XMLPage.send
    
    ThereIsInternet = XMLPage.Status = 200
    Exit Function
HandleErrors:
    MsgBox "Necesita una conexión a internet para realizar esta operación.", vbCritical, "No tiene conexión a internet"
End Function

Public Function FilterUnique(Data As Collection) As Collection
    Dim Unique As New Collection
    Dim Item As Variant
    Dim Dict As New Dictionary
    Dim Key As Variant
    
    For Each Item In Data
        Dict(Item) = 0
    Next Item
    
    For Each Key In Dict.Keys
        Unique.Add Key
    Next Key
    
    Set FilterUnique = Unique
End Function

Public Function LoadColumnData(Sheet As Worksheet, Column As Long, StartRow As Long, EndRow As Long) As Collection
    Dim Row As Long
    Dim Data As New Collection
    
    For Row = StartRow To EndRow
        Data.Add Sheet.Cells(Row, Column).Value
    Next Row
    
    Set LoadColumnData = Data
End Function

Public Function PathJoin(ParamArray Values() As Variant) As String
    Dim fs As New FileSystemObject
    Dim Value As Variant
    Dim Path As String
    
    For Each Value In Values
        Path = fs.BuildPath(Path, Value)
    Next Value
    
    PathJoin = Path
End Function

Public Function CollectionToArray(Data As Collection) As Variant
    Dim Result As Variant
    Dim i As Long

    ReDim Result(Data.Count - 1)

    For i = 0 To Data.Count - 1
        Result(i) = Data(i + 1)
    Next i

    CollectionToArray = Result
End Function

Public Function CharCount(Str As String, Char As String)
    CharCount = Len(Str) - Len(Replace(Str, Char, ""))
End Function

Public Sub WriteFile(FilePath As String, Data As String)
    Dim fs As New FileSystemObject
    Dim Stream As TextStream
    Set Stream = fs.CreateTextFile(FilePath)
    Stream.Write Data
    Stream.Close
End Sub

Public Function GetLastRow(Sheet As Worksheet, Optional Column As Long = 1) As Long
    GetLastRow = Sheet.Cells(Rows.Count, Column).End(xlUp).Row
End Function

Public Function GetMatchRow(Sheet As Worksheet, Column As Long, Value As String) As Long
    On Error Resume Next
    GetMatchRow = WorksheetFunction.Match(Value, Sheet.Columns(Column), 0)
End Function

Public Function GetXmlContentFromZip(ZipPath As String, XmlName As String) As String
    On Error GoTo HandleErrors
    Dim fs As New FileSystemObject
    Dim Stream As TextStream
    Dim Shell As New Shell32.Shell
    Dim Source As Shell32.FolderItem
    Dim Target As Shell32.Folder
    Dim TempPath As String
    
    Set Source = Shell.Namespace(ZipPath).Items.Item(XmlName)
    TempPath = fs.BuildPath(ThisWorkbook.Path, "temp")
    If Not fs.FolderExists(TempPath) Then fs.CreateFolder TempPath
    Set Target = Shell.Namespace(TempPath)
    Target.CopyHere Source, 256
    
    Set Stream = fs.OpenTextFile(fs.BuildPath(TempPath, XmlName))
    GetXmlContentFromZip = Stream.ReadAll
    Stream.Close
    
    fs.DeleteFolder TempPath
    
    Set Shell = Nothing
    Set Source = Nothing
    Set Target = Nothing
    Exit Function
HandleErrors:
    ErrorLog "Error al extraer el contenido del archivo " & XmlName, "GetXmlContentFromZip", Err.Number
End Function

Public Function TaxLess(amount As Double, TaxRate As Double) As Double
    TaxLess = amount / (TaxRate + 1)
End Function

Public Function TaxPlus(TaxBase As Double, TaxRate As Double) As Double
    TaxPlus = TaxBase + TaxBase * TaxRate
End Function

Public Function AmountInLetters(amount As Double, TypeCurrency As String) As String
    On Error Resume Next
    Dim WholePart As Long
    Dim DecimalPart As Long
    Dim WholePartInLetters As String
    Dim DecimalPartInLetters As String
    Dim CurrencyName As String

    WholePart = Int(amount)
    DecimalPart = Round(amount - WholePart, 2) * 100

    WholePartInLetters = UCase(NumberToWords(WholePart))
    DecimalPartInLetters = "CON " & Format(DecimalPart, "00") & "/100"
    CurrencyName = IIf(TypeCurrency = "PEN", "SOLES", "DÓLARES AMERICANOS")
    
    AmountInLetters = WholePartInLetters & " " & DecimalPartInLetters & " " & CurrencyName

    If 1000 <= amount And amount < 2000 Then
        AmountInLetters = "UN " & AmountInLetters
    End If
End Function

Private Function NumberToWords(valor) As String
   Select Case Int(valor)
       Case 0: NumberToWords = "cero"
       Case 1: NumberToWords = "un"
       Case 2: NumberToWords = "dos"
       Case 3: NumberToWords = "tres"
       Case 4: NumberToWords = "cuatro"
       Case 5: NumberToWords = "cinco"
       Case 6: NumberToWords = "seis"
       Case 7: NumberToWords = "siete"
       Case 8: NumberToWords = "ocho"
       Case 9: NumberToWords = "nueve"
       Case 10: NumberToWords = "diez"
       Case 11: NumberToWords = "once"
       Case 12: NumberToWords = "doce"
       Case 13: NumberToWords = "trece"
       Case 14: NumberToWords = "catorce"
       Case 15: NumberToWords = "quince"
       Case Is < 20: NumberToWords = "dieci" & NumberToWords(valor - 10)
       Case 20: NumberToWords = "veinte"
       Case Is < 30: NumberToWords = "veinti" & NumberToWords(valor - 20)
       Case 30: NumberToWords = "treinta"
       Case 40: NumberToWords = "cuarenta"
       Case 50: NumberToWords = "cincuenta"
       Case 60: NumberToWords = "sesenta"
       Case 70: NumberToWords = "setenta"
       Case 80: NumberToWords = "ochenta"
       Case 90: NumberToWords = "noventa"
       Case Is < 100: NumberToWords = NumberToWords(Int(valor \ 10) * 10) & " y " & NumberToWords(valor Mod 10)
       Case 100: NumberToWords = "cien"
       Case Is < 200: NumberToWords = "ciento " & NumberToWords(valor - 100)
       Case 200, 300, 400, 600, 800: NumberToWords = NumberToWords(Int(valor \ 100)) & "cientos"
       Case 500: NumberToWords = "quinientos"
       Case 700: NumberToWords = "setecientos"
       Case 900: NumberToWords = "novecientos"
       Case Is < 1000: NumberToWords = NumberToWords(Int(valor \ 100) * 100) & " " & NumberToWords(valor Mod 100)
       Case 1000: NumberToWords = "mil"
       Case Is < 2000: NumberToWords = "mil " & NumberToWords(valor Mod 1000)
       Case Is < 1000000: NumberToWords = NumberToWords(Int(valor \ 1000)) & " mil"
           If valor Mod 1000 Then NumberToWords = NumberToWords & " " & NumberToWords(valor Mod 1000)
       Case 1000000: NumberToWords = "un millón "
       Case Is < 2000000: NumberToWords = "un millón " & NumberToWords(valor Mod 1000000)
       Case Is < 1000000000000#: NumberToWords = NumberToWords(Int(valor / 1000000)) & "  millones  "
           If (valor - Int(valor / 1000000) * 1000000) _
           Then NumberToWords = NumberToWords & NumberToWords(valor - Int(valor / 1000000) * 1000000)
       Case 1000000000000#: NumberToWords = "un billón "
       Case Is < 2000000000000#
           NumberToWords = "un billón " & NumberToWords(valor - Int(valor / 1000000000000#) * 1000000000000#)
       Case Else: NumberToWords = NumberToWords(Int(valor / 1000000000000#)) & " billones "
           If (valor - Int(valor / 1000000000000#) * 1000000000000#) _
               Then NumberToWords = NumberToWords & " " & NumberToWords(valor - Int(valor / 1000000000000#) * 1000000000000#)
   End Select
End Function
