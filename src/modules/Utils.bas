Attribute VB_Name = "Utils"
Option Explicit

Public Function TaxLess(Amount As Double, TaxRate As Double) As Double
    TaxLess = Amount / (TaxRate + 1)
End Function

Public Function TaxPlus(TaxBase As Double, TaxRate As Double) As Double
    TaxPlus = TaxBase + TaxBase * TaxRate
End Function

Public Function AmountInLetters(Amount As Double, TypeCurrency As AppTypeCurrency) As String
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
    If TypeCurrency = AppTypeCurrencyPEN Then CurrencyName = "SOLES"
    If TypeCurrency = AppTypeCurrencyUSD Then CurrencyName = "DÓLARES AMERICANOS"
    
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
