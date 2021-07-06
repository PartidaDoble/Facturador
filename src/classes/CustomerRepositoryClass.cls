VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomerRepositoryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetAll() As Collection
    Dim Customer1 As New CustomerEntity
    Dim Customer2 As New CustomerEntity
    Dim Customer3 As New CustomerEntity
    Dim Customers As New Collection

    Customer1.DocNumber = "04454845"
    Customer1.Name = "VICTOR HUGO ZEVALLOS BENEGAS"
    
    Customer2.DocNumber = "10123456789"
    Customer2.Name = "VICTOR HUGO ZEVALLOS BENEGAS"
    
    Customer3.DocNumber = "20222444666"
    Customer3.Name = "EMPRESA DE COSAS EMPRESA DE COSAS EMPRESA DE COSAS EMPRESA DE COSAS"
    
    Customers.Add Customer1
    Customers.Add Customer2
    Customers.Add Customer3

    Set GetAll = Customers
End Function
