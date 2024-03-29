VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomerEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DocType As String

Public DocNumber As String

Public Name As String

Public Address As String

Public Ubigeo As String

Public Email As String

Public Function HasValidAddress() As Boolean
    HasValidAddress = Address <> Empty And Len(Ubigeo) = 6
End Function

Public Function GetDocName() As String
    GetDocName = IIf(DocType = "6", "RUC", "DNI")
End Function
