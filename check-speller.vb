Function SpellAmount(amount As Double) As String
    If amount > 999999999 Then
        SpellAmount = "Error: Amount is too large"
    Else
        Dim wholePart As Long
        wholePart = Int(amount)
        Dim fractionalPart As Long
        fractionalPart = CLng((amount - wholePart) * 100)

        SpellAmount = SpellWholeNumber(wholePart)
        If wholePart = 1 Then
            SpellAmount = SpellAmount & " Peso"
        Else
            SpellAmount = SpellAmount & " Pesos"
        End If

        If fractionalPart >= 10 Then
            SpellAmount = SpellAmount & " & " & CStr(fractionalPart) & "/100"
        ElseIf fractionalPart >= 1 Then
            SpellAmount = SpellAmount & " & 0" & CStr(fractionalPart) & "/100"
        End If

        SpellAmount = SpellAmount & " Only"
    End If
End Function

Function SpellWholeNumber(num As Long) As String
    If num >= 1000000 Then
        SpellWholeNumber = SpellWholeNumber(num \ 1000000) & " Million"
        If num Mod 1000000 > 0 Then
            SpellWholeNumber = SpellWholeNumber & " " & SpellWholeNumber(num Mod 1000000)
        End If
    ElseIf num >= 1000 Then
        SpellWholeNumber = SpellWholeNumber(num \ 1000) & " Thousand"
        If num Mod 1000 > 0 Then
            SpellWholeNumber = SpellWholeNumber & " " & SpellWholeNumber(num Mod 1000)
        End If
    ElseIf num >= 100 Then
        SpellWholeNumber = SpellWholeNumber(num \ 100) & " Hundred"
        If num Mod 100 > 0 Then
            SpellWholeNumber = SpellWholeNumber & " " & SpellWholeNumber(num Mod 100)
        End If
    ElseIf num >= 20 Then
        Dim tenWord(2 To 9) As String
        tenWord(2) = "Twenty"
        tenWord(3) = "Thirty"
        tenWord(4) = "Forty"
        tenWord(5) = "Fifty"
        tenWord(6) = "Sixty"
        tenWord(7) = "Seventy"
        tenWord(8) = "Eighty"
        tenWord(9) = "Ninety"

        SpellWholeNumber = tenWord(num \ 10)
        If num Mod 10 > 0 Then
            SpellWholeNumber = SpellWholeNumber & " " & SpellWholeNumber(num Mod 10)
        End If
    Else
        Dim numWord(0 To 19) As String
        numWord(0) = "Zero"
        numWord(1) = "One"
        numWord(2) = "Two"
        numWord(3) = "Three"
        numWord(4) = "Four"
        numWord(5) = "Five"
        numWord(6) = "Six"
        numWord(7) = "Seven"
        numWord(8) = "Eight"
        numWord(9) = "Nine"
        numWord(10) = "Ten"
        numWord(11) = "Eleven"
        numWord(12) = "Twelve"
        numWord(13) = "Thirteen"
        numWord(14) = "Fourteen"
        numWord(15) = "Fifteen"
        numWord(16) = "Sixteen"
        numWord(17) = "Seventeen"
        numWord(18) = "Eighteen"
        numWord(19) = "Nineteen"

        SpellWholeNumber = numWord(num)
    End If
End Function
