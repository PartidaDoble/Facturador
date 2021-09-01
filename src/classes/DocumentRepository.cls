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

Public Function GetAll() As Collection
    Dim Documents As New Collection
    Dim Row As Long
    
    For Row = 2 To GetLastRow(sheetDocuments)
        Documents.Add GetInstance(Row)
    Next Row
    
    Set GetAll = Documents
End Function

Public Function GetItem(Id As String) As DocumentEntity
    On Error GoTo HandleErrors
    Set GetItem = GetInstance(GetMatchRow(sheetDocuments, 1, Id))
    Exit Function
HandleErrors:
    Set GetItem = Nothing
End Function

Public Sub Add(Document As DocumentEntity)
    Save Document, GetLastRow(sheetDocuments) + 1
End Sub

Public Sub Update(Document As DocumentEntity)
    Save Document, GetMatchRow(sheetDocuments, 1, Document.Id)
End Sub

Private Function GetInstance(Row As Long) As DocumentEntity
    Dim Document As New DocumentEntity
    Dim Customer As New CustomerEntity
    Dim NoteInfo As New NoteInfoEntity
    Dim Detraction As New DetractionEntity
    Dim Item As New ItemEntity
    Dim NoteArray As Variant
    Dim DetractionArray As Variant
    
    With sheetDocuments
        Document.Emission = .Cells(Row, 2)
        Document.DocType = Format(.Cells(Row, 3), "00")
        Document.DocSerie = .Cells(Row, 4)
        Document.DocNumber = .Cells(Row, 5)
        
        If .Cells(Row, 6) <> Empty And .Cells(Row, 7) <> Empty And .Cells(Row, 8) <> Empty Then
            Customer.DocType = .Cells(Row, 6)
            Customer.DocNumber = .Cells(Row, 7)
            Customer.Name = .Cells(Row, 8)
            
            Set Document.Customer = Customer
        End If
        
        Document.TypeCurrency = IIf(.Cells(Row, 9) = "Soles", "PEN", "USD")
        
        Item.Quantity = 1
        Item.UnitValue = .Cells(Row, 10)
        Item.IgvRate = .Cells(Row, 11) / .Cells(Row, 10)
        Document.AddItem Item
        
        Document.CancelInfo = .Cells(Row, 13)
        
        If CharCount(.Cells(Row, 14), "-") = 3 Then
            NoteArray = Split(.Cells(Row, 14), "-")
            
            NoteInfo.RefDocEmission = NoteArray(0)
            NoteInfo.RefDocType = NoteArray(1)
            NoteInfo.RefDocSerie = NoteArray(2)
            NoteInfo.RefDocNumber = NoteArray(3)
            
            Set Document.NoteInfo = NoteInfo
        End If
        
        If CharCount(.Cells(Row, 15), "-") = 3 Then
            DetractionArray = Split(.Cells(Row, 15), "-")
            
            Detraction.Code = DetractionArray(0)
            Detraction.Percentage = DetractionArray(1)
            Detraction.Amount = DetractionArray(2)
            Detraction.PaymentMethod = DetractionArray(3)
            
            Set Document.Detraction = Detraction
        End If
        
        Document.Situation = GetSituationFromCode(Left(.Cells(Row, 16), 2))
        Document.Observation = .Cells(Row, 17)
        Document.DailySummary = .Cells(Row, 18)
        Document.Stored = IIf(.Cells(Row, 19) = "Sí", True, False)
        Document.EmailSent = IIf(.Cells(Row, 20) = "Sí", True, False)
    End With
    
    Set GetInstance = Document
End Function

Private Sub Save(Document As DocumentEntity, Row As Long)
    With sheetDocuments
        .Cells(Row, 1) = Document.Id
        .Cells(Row, 2) = Document.Emission
        .Cells(Row, 3) = Document.DocType
        .Cells(Row, 4) = Document.DocSerie
        .Cells(Row, 5) = Document.DocNumber
        
        If Not Document.Customer Is Nothing Then
            .Cells(Row, 6) = Document.Customer.DocType
            .Cells(Row, 7) = "'" & Document.Customer.DocNumber
            .Cells(Row, 8) = Document.Customer.Name
        End If
        
        .Cells(Row, 9) = IIf(Document.TypeCurrency = "PEN", "Soles", "Dólares")
        .Cells(Row, 10) = Format(Document.SubTotal, "#,##0.00")
        .Cells(Row, 11) = Format(Document.Igv, "#,##0.00")
        .Cells(Row, 12) = Document.Total
        .Cells(Row, 13) = Document.CancelInfo

        If Not Document.NoteInfo Is Nothing Then
            .Cells(Row, 14) = Format(Document.NoteInfo.RefDocEmission, "dd/mm/yyyy") & "-" & Document.NoteInfo.Id
        End If
        
        If Not Document.Detraction Is Nothing Then
            .Cells(Row, 15) = Document.Detraction.ConcatData
        End If
        
        .Cells(Row, 16) = GetSituationFromEnum(Document.Situation)
        .Cells(Row, 17) = Document.Observation
        .Cells(Row, 18) = Document.DailySummary
        .Cells(Row, 19) = IIf(Document.Stored, "Sí", "No")
        .Cells(Row, 20) = IIf(Document.DocType <> "03", IIf(Document.EmailSent, "Sí", "No"), Empty)
    End With
End Sub
