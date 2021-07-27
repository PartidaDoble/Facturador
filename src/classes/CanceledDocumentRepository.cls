VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CanceledDocumentRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetAll() As Collection
    Dim CanceledDocument As CanceledDocumentEntity
    Dim CanceledDocuments As New Collection
    Dim Row As Long

    For Row = 2 To GetLastRow(sheetCanceledDocuments)
        Set CanceledDocument = New CanceledDocumentEntity
        
        With sheetCanceledDocuments
            CanceledDocument.GenerationDate = .Cells(Row, 2)
            CanceledDocument.CommunicationDate = .Cells(Row, 3)
            CanceledDocument.DocType = Format(.Cells(Row, 4), "00")
            CanceledDocument.DocNumber = .Cells(Row, 5)
            CanceledDocument.Motivo = .Cells(Row, 6)
            CanceledDocument.Correlative = .Cells(Row, 7)
            CanceledDocument.Situation = GetSituationFromCode(Left(.Cells(Row, 8), 2))
            CanceledDocument.Observation = .Cells(Row, 9)
        End With
        
        CanceledDocuments.Add CanceledDocument
    Next Row
    
    Set GetAll = CanceledDocuments
End Function

Public Sub Add(CanceledDocument As CanceledDocumentEntity)
    Dim Row As Long
    Row = GetLastRow(sheetCanceledDocuments) + 1
    Save CanceledDocument, Row
End Sub

Public Sub Update(CanceledDocument As CanceledDocumentEntity)
    Dim Row As Long
    Row = WorksheetFunction.Match(CanceledDocument.Id, sheetCanceledDocuments.Columns(1), 0)
    Save CanceledDocument, Row
End Sub

Private Sub Save(CanceledDocument As CanceledDocumentEntity, Row As Long)
    With sheetCanceledDocuments
        .Cells(Row, 1) = CanceledDocument.Id
        .Cells(Row, 2) = CanceledDocument.GenerationDate
        .Cells(Row, 3) = CanceledDocument.CommunicationDate
        .Cells(Row, 4) = CanceledDocument.DocType
        .Cells(Row, 5) = CanceledDocument.DocNumber
        .Cells(Row, 6) = CanceledDocument.Motivo
        .Cells(Row, 7) = CanceledDocument.Correlative
        .Cells(Row, 8) = GetSituationFromEnum(CanceledDocument.Situation)
        .Cells(Row, 9) = CanceledDocument.Observation
    End With
End Sub
