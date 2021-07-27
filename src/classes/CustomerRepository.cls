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
        Set Customer = New CustomerEntity
        
        With sheetCustomers
            Customer.DocType = .Cells(Row, 1)
            Customer.DocNumber = .Cells(Row, 2)
            Customer.Name = .Cells(Row, 3)
        End With
        
        Customers.Add Customer
    Next Row
    
    Set GetAll = Customers
End Function

Public Function Add(Customer As CustomerEntity) As Boolean
    Dim Row As Long
    Row = GetLastRow(sheetCustomers) + 1
    With sheetCustomers
        .Cells(Row, 1) = Customer.DocType
        .Cells(Row, 2) = "'" & Customer.DocNumber
        .Cells(Row, 3) = Customer.Name
    End With
End Function

Public Function Contains(Customer As CustomerEntity) As Boolean
    Contains = WorksheetFunction.CountIf(sheetCustomers.Columns(2), Customer.DocNumber) > 0
End Function
