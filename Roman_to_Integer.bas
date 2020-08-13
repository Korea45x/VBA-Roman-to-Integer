Attribute VB_Name = "Module1"
Sub roman_to_int()

Dim roman As String
Dim lenght As Long
Dim counter As Integer
Dim last As Integer
Dim j As Long
Dim k As Integer


roman = InputBox("Enter number in roman numeral format")
length = Len(roman)

counter = 0
last = 0

For j = 1 To length
    
    numeral = Mid(roman, j, 1)
       
    If numeral = "I" Then k = 1
    If numeral = "V" Then k = 5
    If numeral = "X" Then k = 10
    If numeral = "L" Then k = 50
    If numeral = "C" Then k = 100
    If numeral = "D" Then k = 500
    If numeral = "M" Then k = 1000
    
    If last < k Then
        counter = counter - last * 2
    End If
    
    counter = counter + k
    
    last = k
        
Next j
 
MsgBox counter

End Sub

