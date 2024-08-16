  <Code>    
    ''' &lt;summary&gt;
    ''' Switches to Code 128 table B 
    ''' &lt;/summary&gt;
    Public Const SwitchB As Char = ChrW(205) 'ChrW(199)

    ''' &lt;summary&gt;
    ''' Switches to Code 128 table C 
    ''' &lt;/summary&gt;
    Public Const SwitchC As Char = ChrW(204) 'ChrW(200)

    ''' &lt;summary&gt;
    ''' Code 128 table B start character 
    ''' &lt;/summary&gt;
    Public Const StartB As Char = ChrW(209)  'ChrW(204) 

    ''' &lt;summary&gt;
    ''' Code 128 table C start character 
    ''' &lt;/summary&gt;
    Public Const StartC As Char = ChrW(210)  'ChrW(205) 

    ''' &lt;summary&gt;
    ''' Code 128 stop character 
    ''' &lt;/summary&gt;
    Public Const StopCode As Char = ChrW(211) 'ChrW(206) 

    Private Const AsciiLowerBounds As Integer = 127
    Private Const AsciiLowerOffset As Integer = 32
    Private Const AsciiUpperOffset As Integer = 105
    Private Const MaxEncodedLength As Integer = 27
    Private Const AsciiCodePageBoundary As Integer = 95
    Private Const TableCDataWidth As Long = 2
    Private Const Gs1MaximumLength As Integer = 48
    Private Const ParamName As String = "text"

    ''' &lt;summary&gt;
    ''' Converts the input text to a Code 128 encoded string that can be used with a barcode font.
    ''' &lt;/summary&gt;
    ''' &lt;param name="text"&gt;The text you want to convert to a barcode.&lt;/param&gt;
    ''' &lt;returns&gt;An encoded string which produces a bar code when displayed using a Code128 font.&lt;/returns&gt;
    Public Function GetCode128EncodedString(text As String) As String
        ' Validate input
        If text.Length &lt; 1 Then
            Return text
        ElseIf text.Length &gt; Gs1MaximumLength Then
            Throw New ArgumentOutOfRangeException(ParamName, "Input is too long and would not scan properly. Please use less than 48 characters.")
        End If

        ' Preamble and first character
        Dim optimizedBarcode As String = String.Empty
        Dim useTableB As Boolean = True
        Dim checkSum As Integer
        Dim startAt As Integer
        If IsAllNumbers(text, 0, 4) Then
            ' Use Table C
            optimizedBarcode &amp;= StartC
            checkSum = CheckSumChar(StartC, 1)
            useTableB = False
            Dim value As Char = GetTwoDigitsToAscii(text, 0)
            optimizedBarcode &amp;= value
            checkSum += CheckSumChar(value, optimizedBarcode.Length - 1)
            startAt = 2
        Else
            optimizedBarcode &amp;= StartB
            checkSum = CheckSumChar(StartB, 1)

            ' Process 1 digit with table B
            Dim nextValue As Char = text(0)
            CheckValid(nextValue)
            optimizedBarcode &amp;= nextValue
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
                    optimizedBarcode &amp;= SwitchC
                    checkSum += CheckSumChar(SwitchC, optimizedBarcode.Length - 1)
                End If
            End If

            If Not useTableB Then
                ' Using Table C, try to process 2 digits
                If IsAllNumbers(text, position, TableCDataWidth) Then
                    Dim value As Char = GetTwoDigitsToAscii(text, position)
                    optimizedBarcode &amp;= value
                    checkSum += CheckSumChar(value, optimizedBarcode.Length - 1)

                    ' Increment because 2 digits were consumed in this pass
                    position += 1
                Else
                    ' Doesn't have 2 digits left, switch to Table B
                    optimizedBarcode &amp;= SwitchB
                    checkSum += CheckSumChar(SwitchB, optimizedBarcode.Length - 1)
                    useTableB = True
                End If
            End If

            If useTableB Then
                ' Process 1 digit with table B
                Dim nextValue As Char = text(position)
                CheckValid(nextValue)
                optimizedBarcode &amp;= nextValue
                checkSum += CheckSumChar(nextValue, optimizedBarcode.Length - 1)
            End If

            If optimizedBarcode.Length &gt; MaxEncodedLength - 2 Then
                Throw New ArgumentOutOfRangeException(ParamName, "Input is too long and would not scan properly. Compressed length should not exceed 27 characters.")
            End If
        Next position

        checkSum = checkSum Mod 103

        ' Convert the checksum to ASCII code
        checkSum = If(checkSum &lt; AsciiCodePageBoundary, checkSum + AsciiLowerOffset, checkSum + AsciiUpperOffset)

        ' Add the checksum and STOP characters
        optimizedBarcode &amp;= ChrW(checkSum) &amp; StopCode

        Return optimizedBarcode
    End Function

    ''' &lt;summary&gt;
    ''' Table C takes two digits and represents them with a single ASCII character.
    ''' &lt;/summary&gt;
    ''' &lt;param name="text"&gt;The text to pull from.&lt;/param&gt;
    ''' &lt;param name="startIndex"&gt;Starting place in the text.&lt;/param&gt;
    ''' &lt;returns&gt;The ASCII character.&lt;/returns&gt;
    Public Function GetTwoDigitsToAscii(text As String, startIndex As Integer) As Char
        Dim asciiValue As Integer = CInt(text.Substring(startIndex, TableCDataWidth))
        asciiValue = If(asciiValue &lt; AsciiCodePageBoundary, asciiValue + AsciiLowerOffset, asciiValue + AsciiUpperOffset)

        Return ChrW(asciiValue)
    End Function

    ''' &lt;summary&gt;
    ''' Calculation of the checksum used for Code 128. Perform modulo % 103 on the result to get the final value.
    ''' &lt;/summary&gt;
    ''' &lt;param name="check"&gt;The character&lt;/param&gt;
    ''' &lt;param name="position"&gt;The position of that character&lt;/param&gt;
    ''' &lt;returns&gt;Checksum value&lt;/returns&gt;
    Public Function CheckSumChar(check As Char, position As Integer) As Integer
        Dim asciiValue As Integer = AscW(check)

        ' Convert the ASCII value to the checksum value
        asciiValue = If(asciiValue &lt; AsciiLowerBounds, asciiValue - AsciiLowerOffset, asciiValue - AsciiUpperOffset)

        Return position * asciiValue
    End Function

    ''' &lt;summary&gt;
    ''' Looks at a section of a string and test of all those characters are numbers.
    ''' &lt;/summary&gt;
    ''' &lt;param name="sourceString"&gt;The string to test.&lt;/param&gt;
    ''' &lt;param name="startPos"&gt;First character position.&lt;/param&gt;
    ''' &lt;param name="numChars"&gt;How many characters to test.&lt;/param&gt;
    ''' &lt;returns&gt;True when all the checked characters are numeric&lt;/returns&gt;
    Public Function IsAllNumbers(sourceString As String, startPos As Integer, numChars As Integer) As Boolean
        If startPos &lt; 0 OrElse startPos + numChars &gt; sourceString.Length Then Return False

        Dim i As Integer
        For i = startPos To startPos + numChars - 1
            If Not Char.IsDigit(sourceString(i)) Then Return False
        Next i

        Return True
    End Function

    ''' &lt;summary&gt;
    ''' This implementation only supports ASCII characters lower than 128, excluding control characters (below 33)
    ''' &lt;/summary&gt;
    ''' &lt;param name="c"&gt;A character.&lt;/param&gt;
    ''' &lt;returns&gt;True if supported.&lt;/returns&gt;
    Public Function IsValid128(c As Char) As Boolean
        Select Case AscW(c)
            Case 32 To 126
                Return True
            Case Else
                Return False
        End Select
    End Function

    ''' &lt;summary&gt;
    ''' Throws when an unsupported character is passed.
    ''' &lt;/summary&gt;
    ''' &lt;param name="text"&gt;A character.&lt;/param&gt;
    Public Sub CheckValid(text As Char)
        If Not IsValid128(text) Then Throw New ArgumentOutOfRangeException(ParamName, "Invalid character in barcode string. Please only use only printable characters in the lower 127 range.")
    End Sub
  </Code>