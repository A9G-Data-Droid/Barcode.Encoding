<TestClass>
Public Class Code128UnitTest

    ' ReSharper disable StringLiteralTypo
    <TestMethod>
    <DataRow("a", "ÑabÓ")>
    <DataRow("ALLCAPSTEST", "ÑALLCAPSTESTPÓ")>
    <DataRow("123456789123456789123456789123456789123412345678", "Ò,BXn{7Mcy,BXn{7Mcy,B,BXnÊÓ")>
    <DataRow("!@#$%^&*()_+<?>"":}{|,./;", "Ñ!@#$%^&*()_+<?>"":}{|,./;5Ó")>
    <DataRow("mixedUpperLower123&$test", "ÑmixedUpperLower123&$testMÓ")>
    <DataRow("switchfromC123456backtoB", "ÑswitchfromCÌ,BXÍbacktoBiÓ")>
    <DataRow("1234StartWithFour", "Ò,BÍStartWithFourYÓ")>
    <DataRow("EndWithFour1234", "ÑEndWithFourÌ,ByÓ")>
    <DataRow("1234start4567", "Ò,BÍstartÌMc'Ó")>
    Public Sub Encode_Text_To_Known_Values(text As String, expectedEncoded As String)
        Assert.AreEqual(expectedEncoded, Barcode.Encoding.GetCode128EncodedString(text))
    End Sub

    <TestMethod>
    <DataRow("ThisOneIsLessThen48butStillTooLong")>
    <DataRow("ThisOneExceedsTheMaxiumum48CharactersAcceptedByCode128")>
    <DataRow("InvalidCharacter:Ç")>
    Public Sub Encoded_Text_Out_Of_Bounds(text As String)
        Assert.ThrowsException(Of ArgumentOutOfRangeException)(Sub() Barcode.Encoding.GetCode128EncodedString(text))
    End Sub

    <TestMethod>
    <DataRow("a65", "a"c)>
    <DataRow("a97", "Ê"c)>
    Public Sub GetTwoDigitsToAscii_Should_Return_Known_Values(text As String, expectedEncoded As Char)
        Assert.AreEqual(expectedEncoded, Barcode.Encoding.GetTwoDigitsToAscii(text, 1))
    End Sub

    <TestMethod>
    <DataRow("ThisOneIsLessThen48butStillTooLong")>
    <DataRow("ThisOneExceedsTheMaxiumum48CharactersAcceptedByCode128")>
    <DataRow("InvalidCharacter:Ç")>
    Public Sub GetTwoDigitsToAscii_Only_Accepts_Numeric(text As String)
        Assert.ThrowsException(Of InvalidCastException)(Sub() Barcode.Encoding.GetTwoDigitsToAscii(text, 1))
    End Sub

    <TestMethod>
    <DataRow("a"c, 65)>
    <DataRow("Ê"c, 97)>
    Public Sub CheckSumChar_Should_Return_Known_Values(text As Char, expectedEncoded As Integer)
        Assert.AreEqual(expectedEncoded, Barcode.Encoding.CheckSumChar(text, 1))
    End Sub

    <TestMethod>
    <DataRow("0123456789")>
    Public Sub IsAllNumbers_True(text As String)
        Assert.IsTrue(Barcode.Encoding.IsAllNumbers(text, 0, text.Length))
    End Sub
End Class