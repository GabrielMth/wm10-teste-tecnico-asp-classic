<%

Dim connect
Set connect = Server.CreateObject("ADODB.Connection")


connect.ConnectionString = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=wm10;User ID=sa;Password=Wm10;"

On Error Resume Next
connect.Open
If Err.Number <> 0 Then
    Response.Write "FALHOU a conexão com o bd!" & Err.Description
    Response.End
End If
On Error GoTo 0

%>
