Attribute VB_Name = "UtilsTests"
Option Explicit

Private Sub RunTests()
    AmountInLettersTest
    TaxLessTest
    TaxPlusTest
    PathJoinTest
    CharCountTests
    CollectionToArrayTest
    FilterUniqueTest
    IsValidDateTest
End Sub

Public Function Test() As VBAUnit
    Dim UnitTest As New VBAUnit
    Set Test = UnitTest
End Function

Private Sub AmountInLettersTest()
    With Test.It("AmountInLetters")
        .AssertEquals "UN CON 00/100 SOLES", AmountInLetters(1, "PEN")
        .AssertEquals "CIENTO CINCUENTA Y TRES CON 45/100 SOLES", AmountInLetters(153.45, "PEN")
        .AssertEquals "UN MIL OCHOCIENTOS CUARENTA Y CINCO CON 40/100 SOLES", AmountInLetters(1845.4, "PEN")
        .AssertEquals "TREINTA Y CINCO MIL OCHOCIENTOS SESENTA Y DOS CON 80/100 DÓLARES AMERICANOS", AmountInLetters(35862.8, "USD")
    End With
End Sub

Private Sub TaxLessTest()
    With Test.It("TaxLess")
        .AssertEquals 100, TaxLess(118, 0.18)
        .AssertEquals 84.7457627118644, TaxLess(100, 0.18)
    End With
End Sub

Private Sub TaxPlusTest()
    With Test.It("TaxPlus")
        .AssertEquals 118, TaxPlus(100, 0.18)
        .AssertEquals 100, TaxPlus(84.7457627118644, 0.18)
    End With
End Sub

Private Sub PathJoinTest()
    With Test.It("PathJoin")
        .AssertEquals "", PathJoin()
        .AssertEquals "foo\bar", PathJoin("foo", "bar")
        .AssertEquals "foo\bar\baz", PathJoin("foo", "bar", "baz")
    End With
End Sub

Private Sub CharCountTests()
    With Test.It("CharCount")
        .AssertEquals 0, CharCount("foo bar baz", "-")
        .AssertEquals 2, CharCount("foo-bar-baz", "-")
    End With
End Sub

Private Sub CollectionToArrayTest()
    Dim Data As New Collection
    
    Data.Add "foo"
    Data.Add "bar"
    Data.Add "baz"
    
    With Test.It("CollectionToArray")
        .AssertEquals 2, UBound(CollectionToArray(Data))
        .AssertEquals "foo", CollectionToArray(Data)(0)
        .AssertEquals "baz", CollectionToArray(Data)(2)
    End With
End Sub

Private Sub FilterUniqueTest()
    Dim Data As New Collection

    Data.Add "foo"
    Data.Add "bar"
    Data.Add "baz"
    Data.Add "foo"
    Data.Add "bar"
    Data.Add "baz"
    
    With Test.It("FilterUnique")
        .AssertEquals 3, FilterUnique(Data).Count
        .AssertEquals "foo", FilterUnique(Data)(1)
        .AssertEquals "bar", FilterUnique(Data)(2)
        .AssertEquals "baz", FilterUnique(Data)(3)
    End With
End Sub

Private Sub IsValidDateTest()
    With Test.It("FilterUnique")
        .AssertTrue IsValidDate("15/09/2021")
        .AssertFalse IsValidDate("15/19/2021")
        .AssertFalse IsValidDate("15/9/2021")
        .AssertFalse IsValidDate("5/09/2021")
        .AssertFalse IsValidDate("5/09/21")
    End With
End Sub
