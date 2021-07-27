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

Public Emission As Date

Public DocType As String

Public DocSerie As String

Public DocNumber As Long

Public CustomerDocType As String

Public CustomerDocNumber As String

Public CustomerName As String

Public TypeCurrency As String

Public SubTotal As Double

Public Igv As Double

Public Total As Double

Public Situation As SituationEnum

Public Observation As String

Property Get Id() As String
    Id = DocType & "-" & DocSerie & "-" & Format(DocNumber, "00000000")
End Property
