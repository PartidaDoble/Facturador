Attribute VB_Name = "Utils"
Option Explicit

Public Type RateType
    Igv As Double
End Type

Public Function Prop() As AppProperties
    Dim Properties As New AppProperties
    Set Prop = Properties
End Function

Public Function Test() As VBAUnit
    Dim UnitTest As New VBAUnit
    Set Test = UnitTest
End Function
