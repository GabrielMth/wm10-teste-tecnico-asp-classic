<!--#include virtual="/aspbackend/includes/db_conexao.asp" -->
<!--#include virtual="/aspbackend/includes/utils.asp" -->
<!--#include virtual="/aspbackend/includes/validar_tokenapi.asp" -->

<%
Response.CodePage = 65001
Response.Charset = "UTF-8"
Response.ContentType = "application/json"


If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Status = "405 Method Not Allowed"
    Response.AddHeader "Allow", "POST"
    Response.Write "{""erro"":""Metodo HTTP nao permitido""}"
    Response.End
End If

Dim cmdLogout
Set cmdLogout = Server.CreateObject("ADODB.Command")
cmdLogout.ActiveConnection = connect
cmdLogout.CommandText = "sp_logout_usuario"
cmdLogout.CommandType = 4 
cmdLogout.Parameters.Append cmdLogout.CreateParameter("@token", 200, 1, 255, token)
cmdLogout.Execute
Set cmdLogout = Nothing

Response.Status = "200 OK"
Response.Write "{""mensagem"":""Logout realizado com sucesso""}"
%>
