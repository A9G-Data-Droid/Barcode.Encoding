Imports System.Text
Imports System.Text.RegularExpressions


''' <summary>
''' Code 128 Barcode Encoder
''' </summary>
Public Module Code128
    ''' <summary>
    ''' Switches to Code 128 table B 
    ''' </summary>
    Private Const SwitchB As Char = ChrW(200)

    ''' <summary>
    ''' Switches to Code 128 table C 
    ''' </summary>
    Private Const SwitchC As Char = ChrW(199)

    '''' <summary>
    '''' Code 128 table A start character, not used
    '''' </summary>
    'Private Const StartA As Char = ChrW(203)

    ''' <summary>
    ''' Code 128 table B start character 
    ''' </summary>
    Private Const StartB As Char = ChrW(204)

    ''' <summary>
    ''' Code 128 table C start character 
    ''' </summary>
    Private Const StartC As Char = ChrW(205)

    ''' <summary>
    ''' Code 128 stop character 
    ''' </summary>
    Private Const StopCode As Char = ChrW(206)

    Private Const AsciiLowerBounds As Integer = 127
    Private Const AsciiLowerOffset As Integer = 32
    Private Const AsciiUpperOffset As Integer = 100
    Private Const MaxEncodedLength As Integer = 27
    Private Const AsciiCodePageBoundary As Integer = 95

    ''' <summary>
    ''' Converts the input text to a Code 128 encoded string that can be used with a barcode font.
    ''' </summary>
    ''' <param name="text">The text you want to convert to a barcode.</param>
    ''' <returns>An encoded string which produces a bar code when displayed using a Code128 font.</returns>
    Public Function GetCode128EncodedString(text As String) As String
        ' Validate input
        If text.Length < 1 Then
            Return text
        ElseIf text.Length > 48 Then
            Throw New ArgumentException("Input is too long and would not scan properly. Please use less than 48 characters.", NameOf(text))
        End If

        Dim invalidCharacters As New Regex("[^ -~]", RegexOptions.Multiline Or RegexOptions.Compiled)
        If invalidCharacters.IsMatch(text) Then
            Throw New ArgumentException("Invalid character in barcode string. Please only use the lower 127 ASCII characters", NameOf(text))
        End If

        ' Process input
        Dim optimizedBarcode As New StringBuilder
        Dim useTableB As Boolean = True
        Dim position As Integer
        For position = 0 To text.Length - 1
            ' Decide if a switch to table C would save space
            If useTableB Then
                ' Number of digits for Table C optimization to be worth it
                Dim dataChunk As Integer = If(position = 0 OrElse position + 3 = text.Length - 1, 4, 6)

                If IsAllNumbers(text, position, dataChunk) Then
                    ' Use Table C
                    If position = 0 Then
                        optimizedBarcode.Append(StartC)
                    Else
                        optimizedBarcode.Append(SwitchC)
                    End If

                    useTableB = False
                Else
                    If position = 0 Then optimizedBarcode.Append(StartB)
                End If
            End If

            If Not useTableB Then
                ' We are using Table C, try to process 2 digits
                Const tableCDataWidth As Long = 2

                If IsAllNumbers(text, position, tableCDataWidth) Then
                    Dim asciiValue As Integer = CInt(text.Substring(position, tableCDataWidth))
                    asciiValue = If(asciiValue < AsciiCodePageBoundary, asciiValue + AsciiLowerOffset, asciiValue + AsciiUpperOffset)

                    optimizedBarcode.Append(ChrW(asciiValue))

                    ' Increment because 2 digits were consumed in this pass
                    position += 1
                Else
                    ' Doesn't have 2 digits left, switch to Table B
                    optimizedBarcode.Append(SwitchB)
                    useTableB = True
                End If
            End If

            If useTableB Then
                ' Process 1 digit with table B
                optimizedBarcode.Append(text(position))
            End If

            If optimizedBarcode.Length > MaxEncodedLength - 2 Then
                Throw New ArgumentException("Input is too long and would not scan properly. Compressed length should not exceed 27 characters.", NameOf(text))
            End If
        Next position

        ' Calculation of the Weighted modulo-103 checksum
        Dim checkSum As Integer
        For position = 0 To optimizedBarcode.Length - 1
            Dim asciiValue As Integer = AscW(optimizedBarcode.Chars(position))

            ' Convert the ASCII value to the checksum value
            asciiValue = If(asciiValue < AsciiLowerBounds, asciiValue - AsciiLowerOffset, asciiValue - AsciiUpperOffset)

            If position = 0 Then checkSum = asciiValue

            checkSum += position * asciiValue
        Next position

        checkSum = checkSum Mod 103

        ' Convert the checksum to ASCII code
        checkSum = If(checkSum < AsciiCodePageBoundary, checkSum + AsciiLowerOffset, checkSum + AsciiUpperOffset)

        ' Add the checksum and STOP characters
        optimizedBarcode.Append(ChrW(checkSum)).Append(StopCode)

        Return optimizedBarcode.ToString()
    End Function

    ''' <summary>
    ''' Looks at a section of a string and test of all those characters are numbers.
    ''' </summary>
    ''' <param name="sourceString">The string to test.</param>
    ''' <param name="startPos">First character position.</param>
    ''' <param name="numChars">How many characters to test.</param>
    ''' <returns>True when all the checked characters are numeric</returns>
    Public Function IsAllNumbers(sourceString As String, startPos As Integer, numChars As Integer) As Boolean
        If startPos < 0 OrElse startPos + numChars > sourceString.Length Then Return False

        Dim i As Integer
        For i = startPos To startPos + numChars - 1
            If Not Char.IsDigit(sourceString(i)) Then Return False
        Next i

        Return True
    End Function
End Module