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
End Class