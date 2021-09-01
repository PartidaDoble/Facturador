Attribute VB_Name = "App"
Option Explicit

' Envía facturas y notas
Public Sub SendGeneratedInvoicesAndNotes()
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    Dim FileName As String
    
    For Each Document In DocumentRepo.GetAll
        If Document.Situation = CdpXmlGenerado And (Document.IsInvoice Or Document.IsInvoiceNote) Then
            SendInvoiceOrNote Document
        End If
        
        If Document.Situation = CdpConErrores And (Document.IsInvoice Or Document.IsInvoiceNote) Then
            If Document.Observation = "Error al invocar el servicio de SUNAT." Then
                FileName = Prop.Company.Ruc & "-" & Document.Id
                DbUpdateDocumentSituation FileName, "02"
                
                SendInvoiceOrNote Document
            End If
        End If
    Next Document
End Sub

Public Sub SendInvoiceOrNote(Document As DocumentEntity)
    Dim DocumentRepo As New DocumentRepository
    Dim FileName As String
    
    SendElectronicDocument Document.DocType, Document.DocSerie & "-" & Format(Document.DocNumber, "00000000")
    
    FileName = Prop.Company.Ruc & "-" & Document.Id
    Document.Situation = GetSituationFromCode(DbGetDocumentSituation(FileName))
    Document.Observation = DbGetDocumentObservation(FileName)
    
    DocumentRepo.Update Document
End Sub

' Genera y envía facturas y notas anuladas
Public Sub SendCanceledInvoicesAndNotes()
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    Dim CanceledDocumentRepo As New CanceledDocumentRepository
    Dim CanceledDocument As New CanceledDocumentEntity
    Dim Documents As New Collection
    Dim Correlative As Long
    Dim CanceledDocumentNumber As String
    Dim FileName As String
    
    For Each Document In DocumentRepo.GetAll
        If Document.IsInvoice Or Document.IsInvoiceNote Then
            If Document.IsAccepted And Document.IsCanceledNotSent Then
                Documents.Add Document
            End If
        End If
    Next Document
    
    If Documents.Count = 0 Then Exit Sub
    
    Correlative = WorksheetFunction.MaxIfs(sheetCanceledDocuments.Columns(3), sheetCanceledDocuments.Columns(2), Date) + 1
    CanceledDocumentNumber = "RA-" & Format(Date, "yyyymmdd") & "-" & Format(Correlative, "000")
    FileName = Prop.Company.Ruc & "-" & CanceledDocumentNumber
    CreateCanceledDocumentJsonFile FileName & ".json", Documents, Date
    
    RefreshSfsScreen
    GenerateElectronicDocument "RA", CanceledDocumentNumber
    
    If ElectronicDocumentExists(FileName & ".zip") Then
        SendElectronicDocument "RA", CanceledDocumentNumber
        
        CanceledDocument.CommunicationDate = Date
        CanceledDocument.Correlative = Correlative
        CanceledDocument.Situation = GetSituationFromCode(DbGetDocumentSituation(FileName))
        CanceledDocument.Observation = DbGetDocumentObservation(FileName)
        CanceledDocument.Ticket = GetTicketNumber(CanceledDocument.Observation)
        
        CanceledDocumentRepo.Add CanceledDocument
        
        For Each Document In Documents
            Document.Situation = CanceledDocument.Situation
            Document.Observation = CanceledDocument.Observation
            Document.DailySummary = CanceledDocument.Id
            
            DocumentRepo.Update Document
        Next Document
        
        If CanceledDocument.Ticket <> Empty Then
            InfoLog "El resumen de Facturas y Notas anuladas " & CanceledDocumentNumber & " se ha enviado correctamente.", "SendCanceledInvoicesAndNotes"
        End If
    Else
        ErrorLog "Error al generar el xml para anular Facturas y Notas " & CanceledDocumentNumber, "SendCanceledInvoicesAndNotes"
    End If
End Sub

' Envía resumen de boletas, notas y anulados
Public Sub SendGeneratedBoletasAndNotesLoop()
    Dim Documents As Collection
    Dim DocumentsByDay As Collection
    Dim FirstFiveHundred As Collection
    Dim i As Integer
    
    For i = 0 To 4
        Set Documents = GetBoletasAndNotesNotSentOrCanceled
        
        If Documents.Count > 0 Then
            Set DocumentsByDay = FilterDocumentsByDay(Documents)
            Set FirstFiveHundred = GetFirstFiveHundred(DocumentsByDay)
            SendGeneratedBoletasAndNotes FirstFiveHundred
        Else
            Exit For
        End If
    Next i
End Sub

Public Function GetBoletasAndNotesNotSentOrCanceled() As Collection
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    Dim Documents As New Collection
    
    For Each Document In DocumentRepo.GetAll
        If Document.IsBoleta Or Document.IsBoletaNote Then
            If Document.Situation = CdpXmlGenerado And Not Document.SentSummary Then
                Documents.Add Document
            ElseIf Document.IsAccepted And Document.IsCanceledNotSent Then
                Documents.Add Document
            End If
        End If
    Next Document
    
    Set GetBoletasAndNotesNotSentOrCanceled = Documents
End Function

Public Function FilterDocumentsByDay(Documents As Collection) As Collection
    Dim FilteredDocuments As New Collection
    Dim Document As Object
    Dim RefDate As Date
    
    If Documents.Count > 0 Then
        RefDate = Documents(1).Emission
    End If
    
    For Each Document In Documents
        If Document.Emission = RefDate Then
            FilteredDocuments.Add Document
        End If
    Next Document
    
    Set FilterDocumentsByDay = FilteredDocuments
End Function

Public Function GetFirstFiveHundred(Documents As Collection) As Collection
    Dim FirstFiveHundred As New Collection
    Dim Document As DocumentEntity
    Dim Counter As Integer
    
    For Each Document In Documents
        If Counter >= 500 Then Exit For
        FirstFiveHundred.Add Document
        Counter = Counter + 1
    Next Document
    
    Set GetFirstFiveHundred = FirstFiveHundred
End Function

Public Sub SendGeneratedBoletasAndNotes(Documents As Collection)
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    Dim DailySummaryRepo As New DailySummaryRepository
    Dim DailySummary As New DailySummaryEntity
    Dim DailySummaryNumber As String
    Dim Correlative As Integer
    Dim FileName As String
    Dim Counter As Integer
    
    Counter = 0
    Correlative = WorksheetFunction.MaxIfs(sheetDailySummary.Columns(3), sheetDailySummary.Columns(2), Date) + 1
    
    DailySummaryNumber = "RC-" & Format(Date, "yyyymmdd") & "-" & Format(Correlative, "000")
    FileName = Prop.Company.Ruc & "-" & DailySummaryNumber
    CreateDailySummaryJsonFile FileName & ".json", Documents, Date
    
    RefreshSfsScreen
    GenerateElectronicDocument "RC", DailySummaryNumber
    
    If ElectronicDocumentExists(FileName & ".zip") Then
        SendElectronicDocument "RC", DailySummaryNumber
        
        DailySummary.GenerationDate = Date
        DailySummary.Correlative = Correlative
        DailySummary.Situation = GetSituationFromCode(DbGetDocumentSituation(FileName))
        DailySummary.Observation = DbGetDocumentObservation(FileName)
        DailySummary.Ticket = GetTicketNumber(DailySummary.Observation)
        DailySummary.Stored = False
        
        DailySummaryRepo.Add DailySummary
        
        If DailySummary.Ticket <> Empty Then
            For Each Document In Documents
                Document.DailySummary = DailySummaryNumber
                Document.Situation = DailySummary.Situation
                Document.Observation = DailySummary.Observation
                
                DocumentRepo.Update Document
            Next Document
        Else
            ErrorLog "Error al enviar el resumen diario " & DailySummaryNumber, "SendGeneratedBoletasAndNotes"
        End If
        
        InfoLog "El resumen diario de boletas " & DailySummaryNumber & " fue enviado a la SUNAT correctamente.", "SendGeneratedBoletasAndNotes"
    Else
        ErrorLog "Error al generar el xml para el resumen diario de boletas " & DailySummaryNumber, "SendGeneratedBoletasAndNotes"
    End If
End Sub

' Actualiza estado de resumenes diarios
Public Sub UpdateStatusDailySummary()
    Dim DailySummaryRepo As New DailySummaryRepository
    Dim DailySummary As DailySummaryEntity
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    Dim FileName As String
    Dim NewSituation As String
    
    For Each DailySummary In DailySummaryRepo.GetAll
        If DailySummary.Situation = CdpEnviadoPorProcesar Then
            FileName = Prop.Company.Ruc & "-" & DailySummary.Id
            NewSituation = DbGetDocumentSituation(Prop.Company.Ruc & "-" & DailySummary.Id)
            
            If NewSituation <> "08" Then
                DailySummary.Situation = GetSituationFromCode(NewSituation)
                DailySummary.Observation = DbGetDocumentObservation(FileName)
                
                DailySummaryRepo.Update DailySummary
                
                UpdateStatusBoletasAndNotes DailySummary
            End If
        End If
    Next DailySummary
End Sub

Public Sub UpdateStatusBoletasAndNotes(DailySummary As DailySummaryEntity)
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    
    For Each Document In DocumentRepo.GetAll
        If Document.IsBoleta Or Document.IsBoletaNote Then
            If Document.Situation = CdpEnviadoPorProcesar And Document.DailySummary = DailySummary.Id Then
                Document.Situation = DailySummary.Situation
                Document.Observation = DailySummary.Observation
                
                If DailySummary.IsAccepted And Document.IsCanceledNotSent Then
                    Document.CancelInfo = "Sí"
                End If
                
                DocumentRepo.Update Document
            End If
        End If
    Next Document
End Sub

' Actualiza estado de facturas y notas anuladas
Public Sub UpdateStatusCanceledInvoicesAndNotes()
    Dim CanceledDocumentRepo As New CanceledDocumentRepository
    Dim CanceledDocument As CanceledDocumentEntity
    Dim FileName As String
    Dim NewSituation As String
    
    For Each CanceledDocument In CanceledDocumentRepo.GetAll
        If CanceledDocument.Situation = CdpEnviadoPorProcesar Then
            FileName = Prop.Company.Ruc & "-" & CanceledDocument.Id
            NewSituation = DbGetDocumentSituation(FileName)
            
            If NewSituation <> "08" Then
                CanceledDocument.Situation = GetSituationFromCode(NewSituation)
                CanceledDocument.Observation = DbGetDocumentObservation(FileName)
                
                CanceledDocumentRepo.Update CanceledDocument
                
                UpdateCanceledInvoice CanceledDocument
            End If
        End If
    Next CanceledDocument
End Sub

Public Sub UpdateCanceledInvoice(CanceledDocument As CanceledDocumentEntity)
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    
    If CanceledDocument.IsAccepted Then
        For Each Document In DocumentRepo.GetAll
            If Document.DailySummary = CanceledDocument.Id Then
                Document.Situation = CanceledDocument.Situation
                Document.Observation = CanceledDocument.Observation
                
                If CanceledDocument.IsAccepted Then
                    Document.CancelInfo = "Sí"
                End If
                
                DocumentRepo.Update Document
            End If
        Next Document
    End If
End Sub

' Mueve comprobantes aceptados
Public Sub SaveSentInvoicesAndNotes()
    On Error GoTo HandleErrors
    Dim fs As New FileSystemObject
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    Dim MainPath As String
    Dim MonthPath As String
    Dim DocumentPath As String
    Dim JsonFilePath As String
    Dim ElectronicDocumentPath As String
    Dim CdrPath As String
    Dim PdfPath As String
    Dim FileName As String
    
    MainPath = PathJoin(ThisWorkbook.Path, Prop.App.DocDirName)
    If Not fs.FolderExists(MainPath) Then fs.CreateFolder MainPath
    
    For Each Document In DocumentRepo.GetAll
        If Document.IsAccepted And Not Document.Stored And (Document.IsInvoice Or Document.IsInvoiceNote) Then
            MonthPath = PathJoin(MainPath, Format(Document.Emission, "yyyy-mm"))
            If Not fs.FolderExists(MonthPath) Then fs.CreateFolder MonthPath
            
            DocumentPath = PathJoin(MonthPath, Document.Id)
            
            If Not fs.FolderExists(DocumentPath) Then
                fs.CreateFolder DocumentPath
                
                FileName = Prop.Company.Ruc & "-" & Document.Id
                
                JsonFilePath = PathJoin(Prop.Sfs.DATAPath, FileName & ".json")
                If fs.FileExists(JsonFilePath) Then fs.MoveFile JsonFilePath, PathJoin(DocumentPath, FileName & ".json")
                
                ElectronicDocumentPath = PathJoin(Prop.Sfs.ENVIOPath, FileName & ".zip")
                If fs.FileExists(ElectronicDocumentPath) Then fs.MoveFile ElectronicDocumentPath, PathJoin(DocumentPath, FileName & ".zip")
                
                CdrPath = PathJoin(Prop.Sfs.RPTAPath, "R" & FileName & ".zip")
                If fs.FileExists(CdrPath) Then fs.MoveFile CdrPath, PathJoin(DocumentPath, "R" & FileName & ".zip")
                
                PdfPath = PathJoin(Prop.Sfs.REPOPath, FileName & ".pdf")
                If fs.FileExists(PdfPath) Then fs.CopyFile PdfPath, PathJoin(DocumentPath, FileName & ".pdf")
                
                Document.Stored = True
                DocumentRepo.Update Document
                
                DbDeleteRow FileName
            End If
        End If
    Next Document
    
    InfoLog "Los facturas electrónicas se almacenaron (movieron) correctamente.", "SaveSentInvoicesAndNotes"
    Exit Sub
HandleErrors:
    ErrorLog "Error al almacenar (mover) las facturas electrónicas.", "SaveSentInvoicesAndNotes", Err.Number
End Sub

Public Sub SaveSentCanceledInvoicesAndNotes()
    On Error GoTo HandleErrors
    Dim fs As New FileSystemObject
    Dim CanceledDocumentRepo As New CanceledDocumentRepository
    Dim CanceledDocument As CanceledDocumentEntity
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    Dim FileName As String
    Dim JsonFilePath As String
    Dim ElectronicDocumentPath As String
    Dim CdrPath As String
    Dim MainPath As String
    Dim MonthPath As String
    Dim DocumentPath As String
    
    For Each CanceledDocument In CanceledDocumentRepo.GetAll
        If CanceledDocument.IsAccepted And Not CanceledDocument.Stored Then
            FileName = Prop.Company.Ruc & "-" & CanceledDocument.Id
            
            For Each Document In DocumentRepo.GetAll
                If Document.DailySummary = CanceledDocument.Id Then
                    MainPath = PathJoin(ThisWorkbook.Path, Prop.App.DocDirName)
                    MonthPath = PathJoin(MainPath, Format(Document.Emission, "yyyy-mm"))
                    DocumentPath = PathJoin(MonthPath, Document.Id)
                    
                    If fs.FolderExists(DocumentPath) Then
                        ElectronicDocumentPath = PathJoin(Prop.Sfs.ENVIOPath, FileName & ".zip")
                        If fs.FileExists(ElectronicDocumentPath) Then fs.CopyFile ElectronicDocumentPath, PathJoin(DocumentPath, FileName & ".zip")
                        
                        CdrPath = PathJoin(Prop.Sfs.RPTAPath, CanceledDocument.Ticket & ".zip")
                        If fs.FileExists(CdrPath) Then fs.CopyFile CdrPath, PathJoin(DocumentPath, CanceledDocument.Ticket & ".zip")
                    End If
                End If
            Next Document
            
            JsonFilePath = PathJoin(Prop.Sfs.DATAPath, FileName & ".json")
            If fs.FileExists(JsonFilePath) Then fs.DeleteFile JsonFilePath
            
            ElectronicDocumentPath = PathJoin(Prop.Sfs.ENVIOPath, FileName & ".zip")
            If fs.FileExists(ElectronicDocumentPath) Then fs.DeleteFile ElectronicDocumentPath
            
            CdrPath = PathJoin(Prop.Sfs.RPTAPath, CanceledDocument.Ticket & ".zip")
            If fs.FileExists(CdrPath) Then fs.DeleteFile CdrPath
            
            DbDeleteRow FileName
        End If
    Next CanceledDocument
    
    InfoLog "Terminó el proceso de almacenamiento de Facutas y Notas anuladas.", "SaveSentCanceledInvoicesAndNotes"
    Exit Sub
HandleErrors:
    ErrorLog "Error en el proceso de almacenamiento de Facutas y Notas anuladas.", "SaveSentCanceledInvoicesAndNotes", Err.Number
End Sub

Public Sub SaveSentBoletasAndNotes()
    On Error GoTo HandleErrors
    Dim fs As New FileSystemObject
    Dim DocumentRepo As New DocumentRepository
    Dim Document As DocumentEntity
    Dim DailySummaryRepo As New DailySummaryRepository
    Dim DailySummary As DailySummaryEntity
    Dim MainPath As String
    Dim MonthPath As String
    Dim DocumentPath As String
    Dim JsonFilePath As String
    Dim ElectronicDocumentPath As String
    Dim CdrPath As String
    Dim PdfPath As String
    Dim FileName As String
    Dim DailySummaryFileName As String
    
    For Each DailySummary In DailySummaryRepo.GetAll
        If DailySummary.IsAccepted And Not DailySummary.Stored Then
            For Each Document In DocumentRepo.GetAll
                If Document.IsAccepted And Document.DailySummary = DailySummary.Id And Not Document.Stored Then
                    MainPath = PathJoin(ThisWorkbook.Path, Prop.App.DocDirName)
                    If Not fs.FolderExists(MainPath) Then fs.CreateFolder MainPath
                    
                    MonthPath = PathJoin(MainPath, Format(Document.Emission, "yyyy-mm"))
                    If Not fs.FolderExists(MonthPath) Then fs.CreateFolder MonthPath
                    
                    DocumentPath = PathJoin(MonthPath, Document.Id)
                    
                    If Not fs.FolderExists(DocumentPath) Then
                        fs.CreateFolder DocumentPath
                        
                        FileName = Prop.Company.Ruc & "-" & Document.Id
                        
                        JsonFilePath = PathJoin(Prop.Sfs.DATAPath, FileName & ".json")
                        If fs.FileExists(JsonFilePath) Then fs.MoveFile JsonFilePath, PathJoin(DocumentPath, FileName & ".json")
                        
                        ElectronicDocumentPath = PathJoin(Prop.Sfs.ENVIOPath, FileName & ".zip")
                        If fs.FileExists(ElectronicDocumentPath) Then fs.MoveFile ElectronicDocumentPath, PathJoin(DocumentPath, FileName & ".zip")
                        
                        CdrPath = PathJoin(Prop.Sfs.RPTAPath, DailySummary.Ticket & ".zip")
                        If fs.FileExists(CdrPath) Then fs.CopyFile CdrPath, PathJoin(DocumentPath, DailySummary.Ticket & ".zip")
                        
                        PdfPath = PathJoin(Prop.Sfs.REPOPath, FileName & ".pdf")
                        If fs.FileExists(PdfPath) Then fs.CopyFile PdfPath, PathJoin(DocumentPath, FileName & ".pdf")
                        
                        Document.Stored = True
                        DocumentRepo.Update Document
                        
                        DbDeleteRow FileName
                    End If
                End If
            Next Document
            
            DailySummaryFileName = Prop.Company.Ruc & "-" & DailySummary.Id
            
            JsonFilePath = PathJoin(Prop.Sfs.DATAPath, DailySummaryFileName & ".json")
            If fs.FileExists(JsonFilePath) Then fs.DeleteFile JsonFilePath
            
            ElectronicDocumentPath = PathJoin(Prop.Sfs.ENVIOPath, DailySummaryFileName & ".zip")
            If fs.FileExists(ElectronicDocumentPath) Then fs.DeleteFile ElectronicDocumentPath
            
            CdrPath = PathJoin(Prop.Sfs.RPTAPath, DailySummary.Ticket & ".zip")
            If fs.FileExists(CdrPath) Then fs.DeleteFile CdrPath
            
            DailySummary.Stored = True
            DailySummaryRepo.Update DailySummary
            
            DbDeleteRow DailySummaryFileName
        End If
    Next DailySummary
    
    InfoLog "El proceso de almacenamiento de boletas y notas vinculadas a terminado.", "SaveSentBoletasAndNotes"
    Exit Sub
HandleErrors:
    ErrorLog "Error en el proceso de almacenamiento de boletas y notas vinculadas.", "SaveSentBoletasAndNotes", Err.Number
End Sub
