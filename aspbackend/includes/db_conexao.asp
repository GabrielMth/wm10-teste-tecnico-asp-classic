<%
Dim connect

If IsEmpty(connect) Or IsNull(connect) Then
    Set connect = Server.CreateObject("ADODB.Connection")
    connect.ConnectionString = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=wm10;User ID=sa;Password=Wm10;"
    
    On Error Resume Next
    connect.Open
    If Err.Number <> 0 Then
        Response.Status = "500 Internal Server Error"
        Response.Write "{""erro"":""Falha na conexão com o banco: " & Replace(Err.Description, """", "'") & """}"
        Response.End
    End If
    On Error GoTo 0
End If
%>
