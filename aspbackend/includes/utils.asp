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
%>
