Function abbreviate(ByVal s As Variant) As Variant

    Dim result As String
    Dim i As Long
    Dim regEx As New RegExp
    
    i = 1
    result = ""
    
    With regEx
        .Global = True
        .MultiLine = True
        .Pattern = "[aeiou\s&]+"
    End With

    result = regEx.Replace(s, "")
    
    If (Len(Trim(s)) - Len(Replace(s, " ", "")) + 1) > 1 Then
        abbreviate = result
    Else
        abbreviate = s
    End If

End Function
