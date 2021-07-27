Attribute VB_Name = "Utils"
Option Explicit

Public Function GetLastRow(Sheet As Worksheet) As Long
    GetLastRow = Sheet.Cells(Rows.Count, 1).End(xlUp).Row
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

Public Function TaxLess(Amount As Double, TaxRate As Double) As Double
    TaxLess = Amount / (TaxRate + 1)
End Function

Public Function TaxPlus(TaxBase As Double, TaxRate As Double) As Double
    TaxPlus = TaxBase + TaxBase * TaxRate
End Function

Public Sub TraceLog(Message As String, Optional From As String = Empty)
    SettingLog
    LogTrace Message, From
End Sub

Public Sub DebugLog(Message As String, Optional From As String = Empty)
    SettingLog
    LogDebug Message, From
End Sub

Public Sub InfoLog(Message As String, Optional From As String = Empty)
    SettingLog
    LogInfo Message, From
End Sub

Public Sub WarnLog(Message As String, Optional From As String = Empty)
    SettingLog
    LogWarn Message, From
End Sub

Public Sub ErrorLog(Message As String, Optional From As String = Empty, Optional ErrNumber As Long = 0)
    SettingLog
    LogError Message, From, ErrNumber
End Sub

Private Sub SettingLog()
    LogEnabled = True

    If Prop.App.Env = "production" Then
        LogCallback = "LogFile"
        LogThreshold = 3
    End If
End Sub

Public Sub LogFile(Level As Long, Message As String, From As String)
    On Error Resume Next
    Dim fs As New FileSystemObject
    Dim Stream As TextStream
    Dim LevelName As String
    Dim LogFilePath As String
    
    LogEnabled = True
    
    Select Case Level
        Case 1
            LevelName = "Trace"
        Case 2
            LevelName = "Debug"
        Case 3
            LevelName = "Info "
        Case 4
            LevelName = "WARN "
        Case 5
            LevelName = "ERROR"
    End Select

    LogFilePath = fs.BuildPath(ThisWorkbook.Path, "facturador.log")
    If fs.FileExists(LogFilePath) Then
        Set Stream = fs.OpenTextFile(LogFilePath, ForAppending)
    Else
        Set Stream = fs.CreateTextFile(LogFilePath)
    End If
    
    Stream.WriteLine Now & " " & LevelName & " - " & IIf(From <> "", From & ": ", "") & Message
    Stream.Close
End Sub

Public Function AmountInLetters(Amount As Double, TypeCurrency As String) As String
    On Error Resume Next
    Dim WholePart As Long
    Dim DecimalPart As Long
    Dim WholePartInLetters As String
    Dim DecimalPartInLetters As String
    Dim CurrencyName As String

    WholePart = Int(Amount)
    DecimalPart = Round(Amount - WholePart, 2) * 100

    WholePartInLetters = UCase(NumberToWords(WholePart))
    DecimalPartInLetters = "CON " & Format(DecimalPart, "00") & "/100"
    CurrencyName = IIf(TypeCurrency = "PEN", "SOLES", "DÓLARES AMERICANOS")
    
    AmountInLetters = WholePartInLetters & " " & DecimalPartInLetters & " " & CurrencyName

    If 1000 <= Amount And Amount < 2000 Then
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

Public Function GetUnitMeasureFromCode(Code As String) As String
    Select Case Code
        Case "NIU"
            GetUnitMeasureFromCode = "UNIDAD"
        Case "KGM"
            GetUnitMeasureFromCode = "KILOGRAMO"
        Case "LBR"
            GetUnitMeasureFromCode = "LIBRA"
        Case "GRM"
            GetUnitMeasureFromCode = "GRAMO"
        Case "BX"
            GetUnitMeasureFromCode = "CAJA"
        Case "GLL"
            GetUnitMeasureFromCode = "GALON"
        Case "BLL"
            GetUnitMeasureFromCode = "BARRIL"
        Case "CA"
            GetUnitMeasureFromCode = "LATA"
        Case "MIL"
            GetUnitMeasureFromCode = "MILLAR"
        Case "MTQ"
            GetUnitMeasureFromCode = "METRO CUBICO"
        Case "MTR"
            GetUnitMeasureFromCode = "METRO"
        Case Else
            GetUnitMeasureFromCode = "UNIDAD"
    End Select
End Function

Public Function GetUnitMeasureFromName(UnitMeasureName As String) As String
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

Public Function GetSituationFromEnum(Situation As SituationEnum) As String
    Select Case Situation
        Case CdpPorGenerarXml
            GetSituationFromEnum = "01 - por generar xml"
        Case CdpXmlGenerado
            GetSituationFromEnum = "02 - xml generado"
        Case CdpEnviadoAceptado
            GetSituationFromEnum = "03 - enviado y aceptado sunat"
        Case CdpEnviadoAceptadoConObs
            GetSituationFromEnum = "04 - enviado y aceptado sunat con obs."
        Case CdpRechazado
            GetSituationFromEnum = "05 - rechazado por sunat"
        Case CdpConErrores
            GetSituationFromEnum = "06 - con errores"
        Case CdpPorValidarXml
            GetSituationFromEnum = "07 - por validar xml"
        Case CdpEnviadoPorProcesar
            GetSituationFromEnum = "08 - enviado a sunat por procesar"
        Case CdpEnviadoProcesando
            GetSituationFromEnum = "09 - enviado a sunat procesando"
        Case CdpRechazado10
            GetSituationFromEnum = "10 - rechazado por sunat"
        Case CdpEnviadoAceptado11
            GetSituationFromEnum = "11 - enviado y aceptado sunat"
        Case CdpEnviadoAceptadoConObs12
            GetSituationFromEnum = "12 - enviado y aceptado sunat"
    End Select
End Function

Public Function GetSituationFromCode(Code As String) As SituationEnum
    Select Case Code
        Case "01"
            GetSituationFromCode = CdpPorGenerarXml
        Case "02"
            GetSituationFromCode = CdpXmlGenerado
        Case "03"
            GetSituationFromCode = CdpEnviadoAceptado
        Case "04"
            GetSituationFromCode = CdpEnviadoAceptadoConObs
        Case "05"
            GetSituationFromCode = CdpRechazado
        Case "06"
            GetSituationFromCode = CdpConErrores
        Case "07"
            GetSituationFromCode = CdpPorValidarXml
        Case "08"
            GetSituationFromCode = CdpEnviadoPorProcesar
        Case "09"
            GetSituationFromCode = CdpEnviadoProcesando
        Case "10"
            GetSituationFromCode = CdpRechazado10
        Case "11"
            GetSituationFromCode = CdpEnviadoAceptado11
        Case "12"
            GetSituationFromCode = CdpEnviadoAceptadoConObs12
    End Select
End Function
