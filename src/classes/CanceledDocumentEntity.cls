VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CanceledDocumentEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public GenerationDate As Date

Public CommunicationDate As Date

Public DocType As String

Public DocNumber As String

Public Motivo As String

Public Correlative As Long

Public Situation As SituationEnum

Public Observation As String

Property Get Id() As String
    Id = "RA-" & Format(CommunicationDate, "yyyymmdd") & "-" & Format(Correlative, "000")
End Property
