Attribute VB_Name = "Utils"
Option Explicit

Public Type RateType
    Igv As Double
End Type

Public Function Prop() As AppProperties
    Dim properties As New AppProperties
    Set Prop = properties
End Function

Public Function Test() As VBAUnit
    Dim t As New VBAUnit
    Set Test = t
End Function
