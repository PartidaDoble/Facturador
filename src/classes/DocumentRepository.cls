VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Add(Document As DocumentEntity)
    Dim Row As Long
    
    Row = GetLastRow(sheetDocuments) + 1
    With sheetDocuments
        .Cells(Row, 1) = Document.Id
        .Cells(Row, 2) = Document.Emission
        .Cells(Row, 3) = Document.DocType
        .Cells(Row, 4) = Document.DocSerie
        .Cells(Row, 5) = Document.DocNumber
        .Cells(Row, 6) = Document.CustomerDocType
        .Cells(Row, 7) = "'" & Document.CustomerDocNumber
        .Cells(Row, 8) = Document.CustomerName
        .Cells(Row, 9) = IIf(Document.TypeCurrency = "PEN", "Soles", "Dólares")
        .Cells(Row, 10) = Document.SubTotal
        .Cells(Row, 11) = Document.Igv
        .Cells(Row, 12) = Document.Total
        .Cells(Row, 13) = GetSituationFromEnum(Document.Situation)
        .Cells(Row, 14) = Document.Observation
    End With
End Sub

Public Function GetItem(Id As String) As DocumentEntity
    On Error GoTo HandleErrors
    Dim Document As New DocumentEntity
    Dim Row As Long
    
    Row = WorksheetFunction.Match(Id, sheetDocuments.Columns(1), 0)
    With sheetDocuments
        Document.Emission = .Cells(Row, 2)
        Document.DocType = Format(.Cells(Row, 3), "00")
        Document.DocSerie = .Cells(Row, 4)
        Document.DocNumber = .Cells(Row, 5)
        Document.CustomerDocType = .Cells(Row, 6)
        Document.CustomerDocNumber = .Cells(Row, 7)
        Document.CustomerName = .Cells(Row, 8)
        Document.TypeCurrency = IIf(.Cells(Row, 9) = "Soles", "PEN", "USD")
        Document.SubTotal = .Cells(Row, 10)
        Document.Igv = .Cells(Row, 11)
        Document.Total = .Cells(Row, 12)
        Document.Situation = GetSituationFromCode(Left(.Cells(Row, 13), 2))
        Document.Observation = .Cells(Row, 14)
    End With
    
    Set GetItem = Document
    Exit Function
HandleErrors:
    Set GetItem = Nothing
End Function
