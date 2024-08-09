Imports Barcode.Encoding

<TestClass>
Public Class Code128UnitTest

    ' ReSharper disable StringLiteralTypo
    <TestMethod>
    <DataRow("a", "ÌabÎ")>
    <DataRow("ALLCAPSTEST", "ÌALLCAPSTESTPÎ")>
    <DataRow("123456789123456789123456789123456789123412345678", "Í,BXn{7Mcy,BXn{7Mcy,B,BXnÅÎ")>
    <DataRow("!@#$%^&*()_+<?>"":}{|,./;", "Ì!@#$%^&*()_+<?>"":}{|,./;5Î")>
    <DataRow("mixedUpperLower123&$test", "ÌmixedUpperLower123&$testMÎ")>
    <DataRow("switchfromC123456backtoB", "ÌswitchfromCÇ,BXÈbacktoBiÎ")>
    <DataRow("1234StartWithFour", "Í,BÈStartWithFourYÎ")>
    <DataRow("EndWithFour1234", "ÌEndWithFourÇ,ByÎ")>
    Public Sub Encode_Text_To_Known_Values(text As String, expectedEncoded As String)
        Assert.AreEqual(expectedEncoded, Code128.GetCode128EncodedString(text))
    End Sub

    <TestMethod>
    <DataRow("ThisOneIsLessThen48butStillTooLong")>
    <DataRow("ThisOneExceedsTheMaxiumum48CharactersAcceptedByCode128")>
    <DataRow("InvalidCharacter:Ç")>
    Public Sub Encoded_Text_Out_Of_Bounds(text As String)
        Assert.ThrowsException(Of ArgumentException)(Sub() Code128.GetCode128EncodedString(text))
    End Sub

    <TestMethod>
    <DataRow("a65", "a"c)>
    <DataRow("a97", "Å"c)>
    Public Sub GetTwoDigitsToAscii_Should_Return_Known_Values(text As String, expectedEncoded As Char)
        Assert.AreEqual(expectedEncoded, Code128.GetTwoDigitsToAscii(text, 1))
    End Sub

    <TestMethod>
    <DataRow("ThisOneIsLessThen48butStillTooLong")>
    <DataRow("ThisOneExceedsTheMaxiumum48CharactersAcceptedByCode128")>
    <DataRow("InvalidCharacter:Ç")>
    Public Sub GetTwoDigitsToAscii_Only_Accepts_Numeric(text As String)
        Assert.ThrowsException(Of InvalidCastException)(Sub() Code128.GetTwoDigitsToAscii(text, 1))
    End Sub

    <TestMethod>
    <DataRow("a"c, 65)>
    <DataRow("Å"c, 97)>
    Public Sub CheckSumChar_Should_Return_Known_Values(text As Char, expectedEncoded As Integer)
        Assert.AreEqual(expectedEncoded, Code128.CheckSumChar(text, 1))
    End Sub

    <TestMethod>
    <DataRow("0123456789")>
    Public Sub IsAllNumbers_True(text As String)
        Assert.IsTrue(Code128.IsAllNumbers(text, 0, text.Length))
    End Sub
End Class