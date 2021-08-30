VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DailySummaryRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetAll() As Collection
    Dim DailySummary As DailySummaryEntity
    Dim DailySummaries As New Collection
    Dim Row As Long
    
    For Row = 2 To GetLastRow(sheetDailySummary)
        DailySummaries.Add GetInstance(Row)
    Next Row
    
    Set GetAll = DailySummaries
End Function

Public Function GetItem(Id As String) As DailySummaryEntity
    On Error GoTo HandleErrors
    Set GetItem = GetInstance(GetMatchRow(sheetDailySummary, 1, Id))
    Exit Function
HandleErrors:
    Set GetItem = Nothing
End Function

Public Sub Add(DailySummary As DailySummaryEntity)
    Save DailySummary, GetLastRow(sheetDailySummary) + 1
End Sub

Public Sub Update(DailySummary As DailySummaryEntity)
    Save DailySummary, GetMatchRow(sheetDailySummary, 1, DailySummary.Id)
End Sub

Private Function GetInstance(Row As Long) As DailySummaryEntity
    Dim DailySummary As New DailySummaryEntity
    
    With sheetDailySummary
        DailySummary.GenerationDate = .Cells(Row, 2)
        DailySummary.Correlative = .Cells(Row, 3)
        DailySummary.Situation = GetSituationFromCode(Left(.Cells(Row, 4), 2))
        DailySummary.Observation = .Cells(Row, 5)
        DailySummary.Ticket = .Cells(Row, 6)
        DailySummary.Stored = IIf(.Cells(Row, 7) = "Sí", True, False)
    End With
    
    Set GetInstance = DailySummary
End Function

Private Sub Save(DailySummary As DailySummaryEntity, Row As Long)
    With sheetDailySummary
        .Cells(Row, 1) = DailySummary.Id
        .Cells(Row, 2) = DailySummary.GenerationDate
        .Cells(Row, 3) = DailySummary.Correlative
        .Cells(Row, 4) = GetSituationFromEnum(DailySummary.Situation)
        .Cells(Row, 5) = DailySummary.Observation
        .Cells(Row, 6) = "'" & DailySummary.Ticket
        .Cells(Row, 7) = IIf(DailySummary.Stored, "Sí", "No")
    End With
End Sub
