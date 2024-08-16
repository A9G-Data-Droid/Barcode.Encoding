Imports System.Text

Public Module TextRender
    ''' <summary>
    ''' The symbols used to portray the lines in the barcode.
    ''' </summary>
    Private ReadOnly BarcodeLines As New Dictionary(Of String, Char) From {
        {"00", " "c},
        {"01", ChrW(9616)},
        {"10", ChrW(9612)},
        {"11", ChrW(9608)}
    }

    Private ReadOnly Code128BarPerChar As New Dictionary(Of Char, String) From {
        {" "c, "11011001100"},
        {"!"c, "11001101100"},
        {""""c, "11001100110"},
        {"#"c, "10010011000"},
        {"$"c, "10010001100"},
        {"%"c, "10001001100"},
        {"&"c, "10011001000"},
        {"'"c, "10011000100"},
        {"("c, "10001100100"},
        {")"c, "11001001000"},
        {"*"c, "11000100100"},
        {"+"c, "10110011100"},
        {","c, "10001101000"},
        {"-"c, "10101101000"},
        {"."c, "10010000100"},
        {"/"c, "10110110000"},
        {"0"c, "11000010100"},
        {"1"c, "11000100010"},
        {"2"c, "11000110110"},
        {"3"c, "11001000010"},
        {"4"c, "11001010110"},
        {"5"c, "11001101010"},
        {"6"c, "11010001110"},
        {"7"c, "11010010010"},
        {"8"c, "11010110010"},
        {"9"c, "11011001010"},
        {":"c, "10100010100"},
        {";"c, "10101010000"},
        {"<"c, "10100011000"},
        {"="c, "10100011100"},
        {">"c, "10100101100"},
        {"?"c, "10101001000"},
        {"@"c, "10100101000"},
        {"A"c, "10110011110"},
        {"B"c, "10110101110"},
        {"C"c, "10110110010"},
        {"D"c, "10111000110"},
        {"E"c, "10111010010"},
        {"F"c, "10111100010"},
        {"G"c, "11010011110"},
        {"H"c, "11010101110"},
        {"I"c, "11010110010"},
        {"J"c, "11011000110"},
        {"K"c, "11011010010"},
        {"L"c, "11011100010"},
        {"M"c, "11100011110"},
        {"N"c, "11100101110"},
        {"O"c, "11100110010"},
        {"P"c, "11101000110"},
        {"Q"c, "11101010010"},
        {"R"c, "11101100010"},
        {"S"c, "11110011110"},
        {"T"c, "11110101110"},
        {"U"c, "11110110010"},
        {"V"c, "11111000110"},
        {"W"c, "11111010010"},
        {"X"c, "11111100010"},
        {"Y"c, "11100011010"},
        {"Z"c, "11100101010"},
        {"["c, "11100110100"},
        {"\"c, "11101001010"},
        {"]"c, "11101010100"},
        {"^"c, "11101100100"},
        {"_"c, "11101101000"},
        {"`"c, "11110000100"},
        {"a"c, "11110010000"},
        {"b"c, "11110100000"},
        {"c"c, "11110110000"},
        {"d"c, "11111000000"},
        {"e"c, "11111010000"},
        {"f"c, "11111100000"},
        {"g"c, "11111110000"},
        {"h"c, "00000011000"},
        {"i"c, "00000100100"},
        {"j"c, "00000110000"},
        {"k"c, "00001000000"},
        {"l"c, "00001010100"},
        {"m"c, "00001100100"},
        {"n"c, "00001110000"},
        {"o"c, "00010001000"},
        {"p"c, "00010010100"},
        {"q"c, "00010100100"},
        {"r"c, "00010110000"},
        {"s"c, "00011001000"},
        {"t"c, "00011010100"},
        {"u"c, "00011100100"},
        {"v"c, "00011110000"},
        {"w"c, "00100001000"},
        {"x"c, "00100010100"},
        {"y"c, "00100100100"},
        {"z"c, "00100110000"},
        {"{"c, "00101001000"},
        {"|"c, "00101010100"},
        {"}"c, "00101100100"},
        {"~"c, "00101110000"},
        {StartB, "11010000100"},
        {StartC, "11010011100"},
        {StopCode, "11101011000"}
    }

    ''' <summary>
    ''' Use this to display a barcode using Unicode text.
    ''' Always combine with a monospace font for accurate scanning.
    ''' </summary>
    ''' <param name="text">The original text.</param>
    ''' <param name="height">How many lines high should the barcode be?</param>
    ''' <returns>A barcode you can scan when displayed in monospace font.</returns>
    Public Function Code128Barcode(text As String, Optional height As Integer = 8) As String
        Dim barcodeLine As String = GetBarcodeLines(Code128.GetCode128EncodedString(text)) & Environment.NewLine
        Dim output As New StringBuilder(barcodeLine)
        For i As Integer = 0 To height
            output.AppendLine(barcodeLine)
        Next i

        Return output.ToString()
    End Function

    ''' <summary>
    ''' Translates encoded text into the barcode glyphs used to represent those characters.
    ''' </summary>
    ''' <param name="encoded">Encoded text, from GetCode128EncodedString</param>
    ''' <returns>Lines and spaces.</returns>
    Public Function GetBarcodeLines(encoded As String) As String
        Dim barcodeGlyphs As New StringBuilder
        For Each character As Char In encoded
            Dim barChar As String = Code128BarPerChar(character)
            Dim count As Integer = barChar.Length - 1
            For i As Integer = 0 To count Step 2
                Dim glyph As String = barChar(i)
                If i < count Then
                    glyph &= barChar(i + 1)
                Else
                    glyph &= "0"
                End If

                barcodeGlyphs.Append(BarcodeLines(glyph))
            Next i
        Next character

        Return barcodeGlyphs.ToString()
    End Function
End Module