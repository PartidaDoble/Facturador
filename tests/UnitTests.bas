Attribute VB_Name = "UnitTests"
Option Explicit

Private Sub RunAllModuleTests()
    AmountInLettersTest
    TaxLessTest
    TaxPlusTest
End Sub

Private Sub AmountInLettersTest()
    With Test.It("AmountInLettersTest")
        .AssertEquals "UN CON 00/100 SOLES", AmountInLetters(1, "PEN")
        .AssertEquals "CIENTO CINCUENTA Y TRES CON 45/100 SOLES", AmountInLetters(153.45, "PEN")
        .AssertEquals "UN MIL OCHOCIENTOS CUARENTA Y CINCO CON 40/100 SOLES", AmountInLetters(1845.4, "PEN")
        .AssertEquals "TREINTA Y CINCO MIL OCHOCIENTOS SESENTA Y DOS CON 80/100 DÓLARES AMERICANOS", AmountInLetters(35862.8, "USD")
    End With
End Sub

Private Sub TaxLessTest()
    With Test.It("TaxLessTest")
        .AssertEquals 100, TaxLess(118, 0.18)
        .AssertEquals 84.7457627118644, TaxLess(100, 0.18)
    End With
End Sub

Private Sub TaxPlusTest()
    With Test.It("TaxPlusTest")
        .AssertEquals 118, TaxPlus(100, 0.18)
        .AssertEquals 100, TaxPlus(84.7457627118644, 0.18)
    End With
End Sub

Private Sub PathJoinTest()
    With Test.It("TaxPlusTest")
        .AssertEquals "", PathJoin()
        .AssertEquals "foo\bar", PathJoin("foo", "bar")
        .AssertEquals "foo\bar\baz", PathJoin("foo", "bar", "baz")
    End With
End Sub
