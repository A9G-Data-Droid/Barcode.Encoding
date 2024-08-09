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
    Private Const TableCDataWidth As Long = 2

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

        ' Preamble and first character
        Dim optimizedBarcode As New StringBuilder
        Dim useTableB As Boolean = True
        Dim checkSum As Integer
        Dim startAt As Integer
        If IsAllNumbers(text, 0, 4) Then
            ' Use Table C
            optimizedBarcode.Append(StartC)
            checkSum = CheckSumChar(StartC, 1)
            useTableB = False
            Dim value As Char = GetTwoDigitsToAscii(text, 0)
            optimizedBarcode.Append(value)
            checkSum += CheckSumChar(value, optimizedBarcode.Length - 1)
            startAt = 2
        Else
            optimizedBarcode.Append(StartB)
            checkSum = CheckSumChar(StartB, 1)

            ' Process 1 digit with table B
            optimizedBarcode.Append(text(0))
            checkSum += CheckSumChar(text(0), optimizedBarcode.Length - 1)
            startAt = 1
        End If

        ' Process the remaining characters
        Dim position As Integer
        For position = startAt To text.Length - 1
            If useTableB Then ' Decide if a switch to table C would save space
                ' Number of digits for Table C optimization to be worth it
                Dim dataChunk As Integer = If(position + 3 = text.Length - 1, 4, 6)
                If IsAllNumbers(text, position, dataChunk) Then
                    useTableB = False ' Use Table C
                    optimizedBarcode.Append(SwitchC)
                    checkSum += CheckSumChar(SwitchC, optimizedBarcode.Length - 1)
                End If
            End If

            If Not useTableB Then
                ' Using Table C, try to process 2 digits
                If IsAllNumbers(text, position, TableCDataWidth) Then
                    Dim value As Char = GetTwoDigitsToAscii(text, position)
                    optimizedBarcode.Append(value)
                    checkSum += CheckSumChar(value, optimizedBarcode.Length - 1)

                    ' Increment because 2 digits were consumed in this pass
                    position += 1
                Else
                    ' Doesn't have 2 digits left, switch to Table B
                    optimizedBarcode.Append(SwitchB)
                    checkSum += CheckSumChar(SwitchB, optimizedBarcode.Length - 1)
                    useTableB = True
                End If
            End If

            If useTableB Then
                ' Process 1 digit with table B
                optimizedBarcode.Append(text(position))
                checkSum += CheckSumChar(text(position), optimizedBarcode.Length - 1)
            End If

            If optimizedBarcode.Length > MaxEncodedLength - 2 Then
                Throw New ArgumentException("Input is too long and would not scan properly. Compressed length should not exceed 27 characters.", NameOf(text))
            End If
        Next position

        checkSum = checkSum Mod 103

        ' Convert the checksum to ASCII code
        checkSum = If(checkSum < AsciiCodePageBoundary, checkSum + AsciiLowerOffset, checkSum + AsciiUpperOffset)

        ' Add the checksum and STOP characters
        optimizedBarcode.Append(ChrW(checkSum)).Append(StopCode)

        Return optimizedBarcode.ToString()
    End Function

    ''' <summary>
    ''' Table C takes two digits and represents them with a single ASCII character.
    ''' </summary>
    ''' <param name="text">The text to pull from.</param>
    ''' <param name="startIndex">Starting place in the text.</param>
    ''' <returns>The ASCII character.</returns>
    Public Function GetTwoDigitsToAscii(text As String, startIndex As Integer) As Char
        Dim asciiValue As Integer = CInt(text.Substring(startIndex, TableCDataWidth))
        asciiValue = If(asciiValue < AsciiCodePageBoundary, asciiValue + AsciiLowerOffset, asciiValue + AsciiUpperOffset)

        Return ChrW(asciiValue)
    End Function

    ''' <summary>
    ''' Calculation of the checksum used for Code 128. Perform modulo % 103 on the result to get the final value.
    ''' </summary>
    ''' <param name="check">The character</param>
    ''' <param name="position">The position of that character</param>
    ''' <returns>Checksum value</returns>
    Public Function CheckSumChar(check As Char, position As Integer) As Integer
        Dim asciiValue As Integer = AscW(check)

        ' Convert the ASCII value to the checksum value
        asciiValue = If(asciiValue < AsciiLowerBounds, asciiValue - AsciiLowerOffset, asciiValue - AsciiUpperOffset)

        Return position * asciiValue
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