Function getGNumber(name As String)

    Dim AllMatches As Object
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")
    
    RE.Pattern = "(\d{1,3}g)"
    RE.Global = True
    RE.IgnoreCase = True
    
    Set AllMatches = RE.Execute(name)

        If (AllMatches.Count <> 0) Then
            result = AllMatches.Item(0).submatches.Item(0)
        End If

    getGNumber = result
    
End Function

Function getMlNumber(name As String)

    Dim AllMatches As Object
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")
    
    RE.Pattern = "(\s\d{1,4}ml)"
    RE.Global = True
    RE.IgnoreCase = True
    
    
    Set AllMatches = RE.Execute(name)
    
    If (AllMatches.Count <> 0) Then
        result = AllMatches.Item(0).submatches.Item(0)
    End If

    getMlNumber = result

End Function

Function getMgNumber(name As String)

    Dim AllMatches As Object
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")
    
    RE.Pattern = "(\d{1,4}mg)"
    RE.Global = True
    RE.IgnoreCase = True
    
    Set AllMatches = RE.Execute(name)

        If (AllMatches.Count <> 0) Then
            result = AllMatches.Item(0).submatches.Item(0)
        End If

    getMgNumber = result
    
End Function

Function getOnlyNumber(name As String)

    Dim AllMatches As Object
    Dim RE As Object
    Set RE = CreateObject("VBScript.RegExp")
    
    RE.Pattern = "(\s\d{1,3}\s)"
    RE.Global = True
    RE.IgnoreCase = True
    
    
    Set AllMatches = RE.Execute(name)
    
    If (AllMatches.Count <> 0) Then
        result = AllMatches.Item(0).submatches.Item(0)
    End If

    getOnlyNumber = result

End Function

Function RemoveOnlySpace(name As String)
    Set regEx = CreateObject("vbscript.regexp")
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String


    strPattern = "(\s)"

    If strPattern <> "" Then
        strInput = name
        strReplace = ""

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = True
            .Pattern = strPattern
        End With

        If regEx.test(strInput) Then
            RemoveOnlySpace = regEx.Replace(strInput, strReplace)
        Else
            RemoveOnlySpace = name
        End If
    End If
End Function

Function RemoveMlNumber(name As String)
    Set regEx = CreateObject("vbscript.regexp")
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String


    strPattern = "(\s\d{1,4}ml)"

    If strPattern <> "" Then
        strInput = name
        strReplace = ""

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = True
            .Pattern = strPattern
        End With

        If regEx.test(strInput) Then
            RemoveMlNumber = regEx.Replace(strInput, strReplace)
        Else
            RemoveMlNumber = name
        End If
    End If
End Function

Function RemoveMgNumber(name As String)
    Set regEx = CreateObject("vbscript.regexp")
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String


    strPattern = "(\s\d{1,4}mg)"

    If strPattern <> "" Then
        strInput = name
        strReplace = ""

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = True
            .Pattern = strPattern
        End With

        If regEx.test(strInput) Then
            RemoveMgNumber = regEx.Replace(strInput, strReplace)
        Else
            RemoveMgNumber = name
        End If
    End If
End Function



Function RemoveGNumber(name As String)
    Set regEx = CreateObject("vbscript.regexp")
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String


    strPattern = "(\s\d{1,3}g)"

    If strPattern <> "" Then
        strInput = name
        strReplace = ""

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = True
            .Pattern = strPattern
        End With

        If regEx.test(strInput) Then
            RemoveGNumber = regEx.Replace(strInput, strReplace)
        Else
            RemoveGNumber = name
        End If
    End If
End Function

Function RemoveOnlyNumber(name As String)
    Set regEx = CreateObject("vbscript.regexp")
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String


    strPattern = "(\d{1,3}\s)"

    If strPattern <> "" Then
        strInput = name
        strReplace = ""

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = True
            .Pattern = strPattern
        End With

        If regEx.test(strInput) Then
            RemoveOnlyNumber = regEx.Replace(strInput, strReplace)
        Else
            RemoveOnlyNumber = name
        End If
    End If
End Function
