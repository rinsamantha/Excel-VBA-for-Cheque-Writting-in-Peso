# Excel-VBA-for-Cheque-Writting-in-Peso
Function NumberToWords(ByVal MyNumber)
    Dim TempStr As String
    Dim DecimalPlace As Integer
    Dim CurrencyName As String
    Dim SubCurrencyName As String
    
    ' Define currency names
    CurrencyName = "pesos"
    SubCurrencyName = "centavos"
    
    ' Convert MyNumber to string and trim white space
    MyNumber = Trim(CStr(MyNumber))
    
    ' Find position of decimal place (0 if none)
    DecimalPlace = InStr(MyNumber, ".")
    
    ' Process decimal and integer parts
    If DecimalPlace > 0 Then
        Dim Cents As String
        Cents = ConvertDecimalPart(Mid(MyNumber, DecimalPlace + 1))
        
        TempStr = ConvertIntegerPart(Left(MyNumber, DecimalPlace - 1)) & " " & CurrencyName
        
        If Cents <> "" Then
            TempStr = TempStr & " and " & Cents & " " & SubCurrencyName
        Else
            TempStr = TempStr & " only"
        End If
    Else
        TempStr = ConvertIntegerPart(MyNumber) & " " & CurrencyName & " only"
    End If
    
    ' Final output
    NumberToWords = Application.Trim(TempStr)
End Function

Function ConvertIntegerPart(ByVal MyNumber)
    Dim Place(9) As String
    Dim TempStr As String
    Dim Count As Integer

    Place(2) = " thousand"
    Place(3) = " million"
    Place(4) = " billion"
    Place(5) = " trillion"
    
    Count = 1
    Do While MyNumber <> ""
        If Val(Right(MyNumber, 3)) > 0 Then
            TempStr = ConvertHundreds(Right(MyNumber, 3)) & Place(Count) & " " & TempStr
        End If
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    
    ConvertIntegerPart = Application.Trim(TempStr)
End Function

Function ConvertDecimalPart(ByVal MyDecimal)
    Dim Result As String
    
    If Len(MyDecimal) = 1 Then
        MyDecimal = MyDecimal & "0"
    ElseIf Len(MyDecimal) > 2 Then
        MyDecimal = Left(MyDecimal, 2)
    End If
    
    If Val(MyDecimal) > 0 Then
        If Val(MyDecimal) < 10 Then
            Result = GetDigit(MyDecimal)
        Else
            Result = ConvertTens(MyDecimal)
        End If
    Else
        Result = ""
    End If
    
    ConvertDecimalPart = Application.Trim(Result)
End Function

Function ConvertHundreds(ByVal MyNumber)
    Dim Result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & " hundred "
    End If
    
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & ConvertTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    
    ConvertHundreds = Result
End Function

Function ConvertTens(ByVal MyTens)
    Dim Result As String
    If Val(Left(MyTens, 1)) = 1 Then
        Select Case Val(MyTens)
            Case 10: Result = "ten"
            Case 11: Result = "eleven"
            Case 12: Result = "twelve"
            Case 13: Result = "thirteen"
            Case 14: Result = "fourteen"
            Case 15: Result = "fifteen"
            Case 16: Result = "sixteen"
            Case 17: Result = "seventeen"
            Case 18: Result = "eighteen"
            Case 19: Result = "nineteen"
        End Select
    Else
        Select Case Val(Left(MyTens, 1))
            Case 2: Result = "twenty "
            Case 3: Result = "thirty "
            Case 4: Result = "forty "
            Case 5: Result = "fifty "
            Case 6: Result = "sixty "
            Case 7: Result = "seventy "
            Case 8: Result = "eighty "
            Case 9: Result = "ninety "
        End Select
        Result = Result & GetDigit(Right(MyTens, 1))
    End If
    ConvertTens = Result
End Function

Function GetDigit(ByVal MyDigit)
    Select Case Val(MyDigit)
        Case 1: GetDigit = "one"
        Case 2: GetDigit = "two"
        Case 3: GetDigit = "three"
        Case 4: GetDigit = "four"
        Case 5: GetDigit = "five"
        Case 6: GetDigit = "six"
        Case 7: GetDigit = "seven"
        Case 8: GetDigit = "eight"
        Case 9: GetDigit = "nine"
        Case 0: GetDigit = ""
    End Select
End Function

