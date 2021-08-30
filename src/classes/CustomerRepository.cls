VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomerRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetAll() As Collection
    Dim Customer As CustomerEntity
    Dim Customers As New Collection
    Dim Row As Long

    For Row = 2 To GetLastRow(sheetCustomers)
        Customers.Add GetInstance(Row)
    Next Row
    
    Set GetAll = Customers
End Function

Public Function GetItem(Id As String) As CustomerEntity
    On Error GoTo HandleErrors
    Set GetItem = GetInstance(GetMatchRow(sheetCustomers, 2, Id))
    Exit Function
HandleErrors:
    Set GetItem = Nothing
End Function

Public Sub Add(Customer As CustomerEntity)
    Dim Row As Long
    Row = GetLastRow(sheetCustomers) + 1
    With sheetCustomers
        .Cells(Row, 1) = Customer.DocType
        .Cells(Row, 2) = "'" & Customer.DocNumber
        .Cells(Row, 3) = Customer.Name
        .Cells(Row, 4) = Customer.Address
        .Cells(Row, 5) = "'" & Customer.Ubigeo
        .Cells(Row, 6) = Customer.Email
    End With
End Sub

Private Function GetInstance(Row As Long) As CustomerEntity
    Dim Customer As New CustomerEntity
    
    With sheetCustomers
        Customer.DocType = .Cells(Row, 1)
        Customer.DocNumber = .Cells(Row, 2)
        Customer.Name = .Cells(Row, 3)
        Customer.Address = .Cells(Row, 4)
        Customer.Ubigeo = .Cells(Row, 5)
        Customer.Email = .Cells(Row, 6)
    End With
    
    Set GetInstance = Customer
End Function

Public Function Contains(Customer As CustomerEntity) As Boolean
    Contains = WorksheetFunction.CountIf(sheetCustomers.Columns(2), Customer.DocNumber) > 0
End Function
