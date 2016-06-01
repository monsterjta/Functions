'Capitalize first letter of each word in string, except one letter words.
'Function updates "this is a string" to "This Is a String".
'Author: Jonathan Almquist
'Date: 06-01-2016

Dim sValue
sValue = "this is a string"
sValue = FixString(sValue)
wscript.echo sValue

Function FixString(s)
    Dim aWords, i
    aWords = split(s)
    For i = 0 to UBound(aWords)
        If Len(aWords(i)) > 1 Then
            aWords(i) = ucase(left(aWords(i), 1)) & mid(aWords(i), 2)
        Else
            LCase(aWords(i))
        End If
    Next
    FixString = Join(aWords, " ")
End Function