''' <summary>
''' Code 128 Barcode Encoder
''' Must be used with v2.00 font from http://grandzebu.net/informatique/codbar/code128.htm
''' Download: http://grandzebu.net/informatique/codbar/code128.ttf
''' </summary>
Public Module Code128
    ''' <summary>
    ''' Switches to Code 128 table B 
    ''' </summary>
    Public Const SwitchB As Char = ChrW(205) 'ChrW(199)

    ''' <summary>
    ''' Switches to Code 128 table C 
    ''' </summary>
    Public Const SwitchC As Char = ChrW(204) 'ChrW(200)

    ''' <summary>
    ''' Code 128 table B start character 
    ''' </summary>
    Public Const StartB As Char = ChrW(209)  'ChrW(204) 

    ''' <summary>
    ''' Code 128 table C start character 
    ''' </summary>
    Public Const StartC As Char = ChrW(210)  'ChrW(205) 

    ''' <summary>
    ''' Code 128 stop character 
    ''' </summary>
    Public Const StopCode As Char = ChrW(211) 'ChrW(206) 

    Private Const AsciiLowerBounds As Integer = 127
    Private Const AsciiLowerOffset As Integer = 32
    Private Const AsciiUpperOffset As Integer = 105
    Private Const MaxEncodedLength As Integer = 27
    Private Const AsciiCodePageBoundary As Integer = 95
    Private Const TableCDataWidth As Long = 2
    Private Const Gs1MaximumLength As Integer = 48

    ''' <summary>
    ''' Converts the input text to a Code 128 encoded string that can be used with a barcode font.
    ''' </summary>
    ''' <param name="text">The text you want to convert to a barcode.</param>
    ''' <returns>An encoded string which produces a bar code when displayed using a Code128 font.</returns>
    Public Function GetCode128EncodedString(text As String) As String
        ' Validate input
        If text.Length < 1 Then
            Return text
        ElseIf text.Length > Gs1MaximumLength Then
            Throw New ArgumentOutOfRangeException("text", "Input is too long and would not scan properly. Please use less than 48 characters.")
        End If

        ' Preamble and first character
        Dim optimizedBarcode As String = String.Empty
        Dim useTableB As Boolean = True
        Dim checkSum As Integer
        Dim startAt As Integer
        If IsAllNumbers(text, 0, 4) Then
            ' Use Table C
            optimizedBarcode &= StartC
            checkSum = CheckSumChar(StartC, 1)
            useTableB = False
            Dim value As Char = GetTwoDigitsToAscii(text, 0)
            optimizedBarcode &= value
            checkSum += CheckSumChar(value, optimizedBarcode.Length - 1)
            startAt = 2
        Else
            optimizedBarcode &= StartB
            checkSum = CheckSumChar(StartB, 1)

            ' Process 1 digit with table B
            Dim nextValue As Char = text(0)
            CheckValid(nextValue)
            optimizedBarcode &= nextValue
            checkSum += CheckSumChar(nextValue, optimizedBarcode.Length - 1)
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
                    optimizedBarcode &= SwitchC
                    checkSum += CheckSumChar(SwitchC, optimizedBarcode.Length - 1)
                End If
            End If

            If Not useTableB Then
                ' Using Table C, try to process 2 digits
                If IsAllNumbers(text, position, TableCDataWidth) Then
                    Dim value As Char = GetTwoDigitsToAscii(text, position)
                    optimizedBarcode &= value
                    checkSum += CheckSumChar(value, optimizedBarcode.Length - 1)

                    ' Increment because 2 digits were consumed in this pass
                    position += 1
                Else
                    ' Doesn't have 2 digits left, switch to Table B
                    optimizedBarcode &= SwitchB
                    checkSum += CheckSumChar(SwitchB, optimizedBarcode.Length - 1)
                    useTableB = True
                End If
            End If

            If useTableB Then
                ' Process 1 digit with table B
                Dim nextValue As Char = text(position)
                CheckValid(nextValue)
                optimizedBarcode &= nextValue
                checkSum += CheckSumChar(nextValue, optimizedBarcode.Length - 1)
            End If

            If optimizedBarcode.Length > MaxEncodedLength - 2 Then
                Throw New ArgumentOutOfRangeException("text", "Input is too long and would not scan properly. Compressed length should not exceed 27 characters.")
            End If
        Next position

        checkSum = checkSum Mod 103

        ' Convert the checksum to ASCII code
        checkSum = If(checkSum < AsciiCodePageBoundary, checkSum + AsciiLowerOffset, checkSum + AsciiUpperOffset)

        ' Add the checksum and STOP characters
        optimizedBarcode &= ChrW(checkSum) & StopCode

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

    Public Function IsValid128(c As Char) As Boolean
        Select Case AscW(c)
            Case 32 To 126
                Return True
            Case Else
                Return False
        End Select
    End Function

    Public Sub CheckValid(c As Char)
        If Not IsValid128(c) Then Throw New ArgumentOutOfRangeException("c", "Invalid character in barcode string. Please only use only printable characters in the lower 127 range.")
    End Sub
End Module