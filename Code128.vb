Option Explicit

' Creates string ready to print as barcode Code128
Function Code128(inString As String) As String
Dim i As Integer
Dim CheckDigit As Integer
Dim currentSubset As String
Dim maxA, maxB, maxC As Long
    ' First we have to test is this string can be printed as barcode Code 128
    For i = 1 To Len(inString)
        If (Asc(Mid(inString, i)) < 0 Or Asc(Mid(inString, i)) > 127) Then
            Exit For
        End If
    Next i
    ' We will proceed only if it is not empty string
    ' and all characters in instring are valid
    If (Len(inString) > 0 And Len(inString) <= i) Then
        ' First we have to find out which subset to use
        maxA = TestA(inString)
        maxB = TestB(inString)
        maxC = TestC(inString)
        ' If we have a posibility to start with subset C - do it.
        If (maxC > 0) Then
            currentSubset = "C"
        Else
            ' More prefered is subset B
            If (maxB >= maxA) Then
                currentSubset = "B"
            Else
                currentSubset = "A"
            End If
        End If
        If (currentSubset = "A") Then
            CheckDigit = 103
        ElseIf (currentSubset = "B") Then
            CheckDigit = 104
        ElseIf (currentSubset = "C") Then
            CheckDigit = 105
        End If
        Code128 = ChrW(CheckDigit + 32)
        While (Len(inString) > 0)
            If (currentSubset = "A" Or currentSubset = "SA") Then
                If (Asc(inString) >= 31) Then
                    CheckDigit = CheckDigit + (Asc(inString) - 32) * Len(Code128)
                    Code128 = Code128 & Left(inString, 1)
                Else
                    CheckDigit = CheckDigit + (Asc(inString) + 64) * Len(Code128)
                    Code128 = Code128 & ChrW(Asc(inString) + 96)
                End If
                inString = Mid(inString, 2)
                If (currentSubset = "SA") Then
                    currentSubset = "B"
                End If
            ElseIf (currentSubset = "B" Or currentSubset = "SB") Then
                CheckDigit = CheckDigit + (Asc(inString) - 32) * Len(Code128)
                Code128 = Code128 & Left(inString, 1)
                inString = Mid(inString, 2)
                If (currentSubset = "SB") Then
                    currentSubset = "A"
                End If
            ElseIf (currentSubset = "C") Then
                CheckDigit = CheckDigit + Val(Mid(inString, 1, 2)) * Len(Code128)
                Code128 = Code128 & ChrW(Val(Mid(inString, 1, 2)) + 32)
                inString = Mid(inString, 3)
            End If
            If (Len(inString) > 0) Then
                maxA = TestA(inString)
                maxB = TestB(inString)
                maxC = TestC(inString)
                If (maxC > 2) Then
                    If (currentSubset <> "C") Then
                        CheckDigit = CheckDigit + 99 * Len(Code128)
                        Code128 = Code128 & ChrW(99 + 32)
                    End If
                    currentSubset = "C"
                Else
                    If (currentSubset = "A" And maxA > 0) Then
                    ElseIf (currentSubset = "B" And maxB > 0) Then
                    ElseIf (currentSubset = "C" And maxC > 0) Then
                    Else
                        If (currentSubset = "A") Then
                            maxA = TestA(Mid(inString, 2))
                            maxB = TestB(Mid(inString, 2))
                            If (maxA >= maxB) Then
                                CheckDigit = CheckDigit + 98 * Len(Code128)
                                Code128 = Code128 & ChrW(98 + 32)
                                currentSubset = "SB"
                            Else
                                CheckDigit = CheckDigit + 100 * Len(Code128)
                                Code128 = Code128 & ChrW(100 + 32)
                                currentSubset = "B"
                            End If
                        ElseIf (currentSubset = "B") Then
                            maxA = TestA(Mid(inString, 2))
                            maxB = TestB(Mid(inString, 2))
                            If (maxB >= maxA) Then
                                CheckDigit = CheckDigit + 98 * Len(Code128)
                                Code128 = Code128 & ChrW(98 + 32)
                                currentSubset = "SA"
                            Else
                                CheckDigit = CheckDigit + 101 * Len(Code128)
                                Code128 = Code128 & ChrW(101 + 32)
                                currentSubset = "A"
                            End If
                        ElseIf (currentSubset = "C") Then
                            If (maxB >= maxA) Then
                                CheckDigit = CheckDigit + 100 * Len(Code128)
                                Code128 = Code128 & ChrW(100 + 32)
                                currentSubset = "B"
                            Else
                                CheckDigit = CheckDigit + 101 * Len(Code128)
                                Code128 = Code128 & ChrW(101 + 32)
                                currentSubset = "A"
                            End If
                        End If
                    End If
                End If
            End If
        Wend
        CheckDigit = CheckDigit Mod 103
        Code128 = Code128 + ChrW(CheckDigit + 32) + ChrW(106 + 32)
    ' For empty or invalid string we return empty string
    Else
        Code128 = ""
    End If
End Function

' Tests how many characters from start of input string can be printed with subset A
' Subset allows to print control codes (from 0 to 31) and
' signs, digits and upper case letters (from 32 to 95)
Function TestA(inString As String) As Long
Dim i As Long
    TestA = 0
    For i = 1 To Len(inString)
        If (Asc(Mid(inString, i)) < 0 Or Asc(Mid(inString, i)) > 95) Then
            Exit For
        End If
        TestA = i
    Next i
End Function

' Tests how many characters from start of input string can be printed with subset B
' Subset allows to print signs, digits, upper and lower case letters (from 32 to 127)
Function TestB(inString As String) As Long
Dim i As Long
    TestB = 0
    For i = 1 To Len(inString)
        If (Asc(Mid(inString, i)) < 32 Or Asc(Mid(inString, i)) > 127) Then
            Exit For
        End If
        TestB = i
    Next i
End Function

' Tests how many characters from start of input string can be printed with subset C
' Subset allows to print pairs of digits (from 48 to 57)
Function TestC(inString As String) As Long
Dim i As Long
    TestC = 0
    For i = 1 To Len(inString) \ 2
        If (Asc(Mid(inString, (i - 1) * 2 + 1)) < 48 Or Asc(Mid(inString, (i - 1) * 2 + 1)) > 57) Then
            Exit For
        End If
        If (Asc(Mid(inString, (i - 1) * 2 + 2)) < 48 Or Asc(Mid(inString, (i - 1) * 2 + 2)) > 57) Then
            Exit For
        End If
        TestC = i
    Next i
    TestC = TestC * 2
End Function
