VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DailySummaryEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public GenerationDate As Date

Public Correlative As Long

Public Situation As SituationEnum

Public Observation As String

Public Ticket As String

Public Stored As Boolean

Property Get Id() As String
    Id = "RC-" & Format(GenerationDate, "yyyymmdd") & "-" & Format(Correlative, "000")
End Property

Public Function IsAccepted() As Boolean
    IsAccepted = Situation = CdpEnviadoAceptado Or Situation = CdpEnviadoAceptadoConObs
End Function

