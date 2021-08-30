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
    Dim CanceledDocuments As New Collection
    Dim Row As Long
    
    For Row = 2 To GetLastRow(sheetCanceledDocuments)
        CanceledDocuments.Add GetInstance(Row)
    Next Row
    
    Set GetAll = CanceledDocuments
End Function

Public Function GetItem(Id As String) As CanceledDocumentEntity
    On Error GoTo HandleErrors
    Set GetItem = GetInstance(GetMatchRow(sheetCanceledDocuments, 1, Id))
    Exit Function
HandleErrors:
    Set GetItem = Nothing
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

Private Function GetInstance(Row As Long) As CanceledDocumentEntity
    Dim CanceledDocument As New CanceledDocumentEntity
    
    With sheetCanceledDocuments
        CanceledDocument.CommunicationDate = .Cells(Row, 2)
        CanceledDocument.Correlative = .Cells(Row, 3)
        CanceledDocument.Situation = GetSituationFromCode(Left(.Cells(Row, 4), 2))
        CanceledDocument.Observation = .Cells(Row, 5)
        CanceledDocument.Ticket = .Cells(Row, 6)
        CanceledDocument.Stored = IIf(.Cells(Row, 7) = "Sí", True, False)
    End With
    
    Set GetInstance = CanceledDocument
End Function

Private Sub Save(CanceledDocument As CanceledDocumentEntity, Row As Long)
    With sheetCanceledDocuments
        .Cells(Row, 1) = CanceledDocument.Id
        .Cells(Row, 2) = CanceledDocument.CommunicationDate
        .Cells(Row, 3) = CanceledDocument.Correlative
        .Cells(Row, 4) = GetSituationFromEnum(CanceledDocument.Situation)
        .Cells(Row, 5) = CanceledDocument.Observation
        .Cells(Row, 6) = "'" & CanceledDocument.Ticket
        .Cells(Row, 7) = IIf(CanceledDocument.Stored, "Sí", "No")
    End With
End Sub
