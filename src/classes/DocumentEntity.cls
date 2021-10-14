VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public OperationCode As String

Public Emission As Date

Public EmissionTime As Date

Public TypeCurrency As String

Public DocType As String

Public DocSerie As String

Public DocNumber As Long

Public Customer As CustomerEntity

Public NoteInfo As NoteInfoEntity

Public Detraction As DetractionEntity

Public WayPay As WayPayEntity

Private pItems As New Collection


' Additional info
Public CancelInfo As String

Public Situation As SituationEnum

Public Observation As String

Public DailySummary As String

Public Stored As Boolean

Public EmailSent As Boolean

Property Get Id() As String
    Id = DocType & "-" & DocSerie & "-" & Format(DocNumber, "00000000")
End Property

Public Property Get Total() As Double
    Dim Item As Variant
    Dim PriceSum As Double

    For Each Item In pItems
        PriceSum = PriceSum + Item.SalePrice
    Next Item

    Total = PriceSum
End Property

Public Property Get SubTotal() As Double
    Dim Item As Variant
    Dim ValueSum As Double

    For Each Item In pItems
        ValueSum = ValueSum + Item.SaleValue
    Next Item

    SubTotal = ValueSum
End Property

Public Property Get Igv() As Double
    Dim Item As Variant
    Dim IgvSum As Double
    
    For Each Item In pItems
        IgvSum = IgvSum + Item.Igv
    Next Item
    
    Igv = IgvSum
End Property

Public Property Get Items() As Collection
    Set Items = pItems
End Property

Public Sub AddItem(Item As ItemEntity)
    pItems.Add Item
End Sub

Public Function IsAccepted() As Boolean
    IsAccepted = Situation = CdpEnviadoAceptado Or Situation = CdpEnviadoAceptadoConObs
End Function

Public Function IsInvoice() As Boolean
    IsInvoice = DocType = "01"
End Function

Public Function IsBoleta() As Boolean
    IsBoleta = DocType = "03"
End Function

Public Function IsNote() As Boolean
    IsNote = DocType = "07" Or DocType = "08"
End Function

Public Function IsInvoiceNote() As Boolean
    IsInvoiceNote = IsNote And Left(DocSerie, 1) = "F"
End Function

Public Function IsBoletaNote() As Boolean
    IsBoletaNote = IsNote And Left(DocSerie, 1) = "B"
End Function

Public Function IsCanceled() As Boolean
    IsCanceled = CancelInfo = "Sí"
End Function

Public Function IsCanceledNotSent() As Boolean
    IsCanceledNotSent = Left(CancelInfo, 2) = "X|"
End Function

Public Function SentSummary() As Boolean
    SentSummary = Left(DailySummary, 2) = "RC" Or Left(DailySummary, 2) = "RA"
End Function

Public Function EmitedMoreSevenDaysAgo() As Boolean
    EmitedMoreSevenDaysAgo = Date - Emission > 7
End Function

Public Function GetState() As String
    If IsBoleta Then GetState = "1"
    If IsBoletaNote Then GetState = "2"
    If IsCanceledNotSent Then GetState = "3"
End Function

Public Function GetName() As String
    Select Case DocType
        Case "01"
            GetName = "Factura Electrónica"
        Case "03"
            GetName = "Boleta de Venta Electrónica"
        Case "07"
            GetName = "Nota de Crédito Electrónica"
        Case "08"
            GetName = "Nota de Débito Electrónica"
    End Select
End Function
