Attribute VB_Name = "modRoman"
Option Explicit
Option Compare Binary
' --------------------------------------------------------------------------------------
' ROMAN/DECIMAL CONVERSION
' --------------------------------------------------------------------------------------
'	Copyright 2017 Richard S. Tallent, II

'	Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files
'	(the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge,
'	publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to
'	do so, subject to the following conditions:

'	The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

'	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
'	MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
'	LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
'	CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


' *********************************************************************************
' RomanToDecimal
' Pass a Roman numeral string, get back the Decimal equivalent. Supports all numbers
' up to 16384. Does not support bar notation. Can be used as a UDF.
' *********************************************************************************
Public Function RomanToDecimal(ByVal s As String) As Integer
    Application.Volatile False ' Pure function

    Dim l As Integer, x As Integer, total As Integer

    l = Len(s)

    ' Work right to left
    For x = l To 1 Step -1
        Select Case Mid$(s, x, 1)
            Case "I", "i":
                ' 1-3, 4, 9, cannot appear left of L/C/D/M, maximum of 3 occurrances as a "1" value and 1 occurrance left of X or V
                Select Case total
                    Case 5, 10: total = total - 1                       ' IV, IX
                    Case Is >= 9, 3, 4: Exit Function                   ' IIX, IIII, IIV
                    Case Else: total = total + 1                        ' I, I, III
                End Select
            Case "V", "v":
                ' 5, cannot appear left of V/X/L/C/D/M
                If total >= 5 Then Exit Function
                total = total + 5
            Case "X", "x":
                ' 10, 20, 30, 40, 90, cannot appear left of D/M, maximum of 3 occurrances as an "X" value, once before L/C
                Select Case total
                    Case 50 To 59, 100 To 109: total = total - 10       ' XL..., XC...
                    Case Is >= 30: Exit Function                        ' XXXX...
                    Case Else: total = total + 10                       ' X..., XX..., XXX...
                End Select
            Case "L", "l":
                ' 50, cannot appear left of L/C/D/M
                If total >= 50 Then Exit Function
                total = total + 50
            Case "C", "c":
                ' 100, 200, 300, 900, maximum of 3 occurrances
                Select Case total
                    Case 500 To 599, 1000 To 1099: total = total - 100  ' CD..., CM...
                    Case Is >= 300: Exit Function                       ' CCCC, CMM..., CMD...
                    Case Else: total = total + 100                      ' C..., CC..., CCC...
                End Select
            Case "D", "d":
                ' 500, cannot appear left of D/M
                If total >= 500 Then Exit Function
                total = total + 500
            Case "M", "m":
                ' 1000, no limit on occurrances
                total = total + 1000
            Case Else:
                ' Not a Roman numeral
                Exit Function
        End Select
    Next

    RomanToDecimal = total

End Function

' *********************************************************************************
' DecimalToRoman
' Pass a decimal number 1-16384, get back a Roman numeral. Does not support bar
' notation. Can be used as a UDF.
' *********************************************************************************
Public Function DecimalToRoman(ByVal i As Integer) As String
    Application.Volatile False ' Pure function

    If i <= 0 Then Exit Function

    Dim s As String

    ' Thousands
    s = String$(i \ 1000, "M")
    i = i Mod 1000
    If i = 0 Then DecimalToRoman = s: Exit Function

    ' Hundreds
    Select Case i
        Case Is >= 900: s = s & "CM"
        Case 400 To 499: s = s & "CD"
        Case Else
            If i >= 500 Then
                s = s & "D"
                i = i - 500
            End If
            Select Case i
                Case Is >= 300: s = s & "CCC"
                Case Is >= 200: s = s & "CC"
                Case Is >= 100: s = s & "C"
            End Select
    End Select
    i = i Mod 100
    If i = 0 Then DecimalToRoman = s: Exit Function

    ' Tens
    Select Case i
        Case Is >= 90: s = s & "XC"
        Case 40 To 49: s = s & "XL"
        Case Else
            If i >= 50 Then
                s = s & "L"
                i = i - 50
            End If
            Select Case i
                Case Is >= 30: s = s & "XXX"
                Case Is >= 20: s = s & "XX"
                Case Is >= 10: s = s & "X"
            End Select
    End Select
    i = i Mod 10
    If i = 0 Then DecimalToRoman = s: Exit Function

    ' Ones
    Select Case i
        Case 9: s = s & "IX"
        Case 4: s = s & "IV"
        Case Else
            If i >= 5 Then
                s = s & ("V")
                i = i - 5
            End If
            Select Case i
                Case 3: s = s & "III"
                Case 2: s = s & "II"
                Case 1: s = s & "I"
            End Select
    End Select

    DecimalToRoman = s

End Function
