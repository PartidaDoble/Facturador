VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NoteInfoEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MotiveCode As String

Public Motive As String

Public RefDocEmission As Date

Public RefDocType As String

Public RefDocSerie As String

Public RefDocNumber As String

Property Get Id() As String
    Id = RefDocType & "-" & RefDocSerie & "-" & Format(RefDocNumber, "00000000")
End Property
