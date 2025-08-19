<%
Function JSONEscape(str)
    If IsNull(str) Then
        JSONEscape = ""
    Else
        str = Replace(str, "\", "\\")
        str = Replace(str, """", "\""")
        str = Replace(str, vbCrLf, "\n")
        str = Replace(str, vbCr, "\n")
        str = Replace(str, vbLf, "\n")
        JSONEscape = str
    End If
End Function


Function ParseSimpleJSON(json)
    Dim data, item, keyValue, dict
    Set dict = Server.CreateObject("Scripting.Dictionary")

    
    json = Replace(json, "{", "")
    json = Replace(json, "}", "")
    json = Replace(json, """", "")
    json = Replace(json, vbCr, "")
    json = Replace(json, vbLf, "")
    json = Replace(json, vbTab, "")

    
    data = Split(json, ",")

    For Each item In data
        keyValue = Split(item, ":")
        If UBound(keyValue) = 1 Then
            dict(Trim(keyValue(0))) = Trim(keyValue(1))
        End If
    Next

    Set ParseSimpleJSON = dict
End Function


Function GenerateRandomToken()
    Dim token, i, chars
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789" 
    Randomize
    token = ""
    For i = 1 To 32 
        token = token & Mid(chars, Int((Len(chars) * Rnd) + 1), 1)
    Next
    GenerateRandomToken = UCase(token) 
End Function

%>
