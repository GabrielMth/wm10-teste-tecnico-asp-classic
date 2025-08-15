<%
<!--#include file="../includes/db_conexao.asp"-->
<!--#include file="../includes/utils.asp"-->
<!--#include file="../includes/jwt_utils.asp"-->

If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Status = "405 Method Not Allowed"
    Response.AddHeader "Allow", "POST"
    Response.ContentType = "application/json"
    Response.Write "{""erro"":""Método não permitido. Use POST""}"
    Response.End
End If

Response.ContentType = "application/json"

Dim usuario, senha, rs, cmd, token, payloadJson
usuario = Request.Form("usuario")
senha = Request.Form("senha")

If usuario = "" Or senha = "" Then
    Response.Status = "400 Bad Request"
    Response.Write "{""erro"":""usuario e senha são obrigatórios""}"
    Response.End
End If

Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = connect
cmd.CommandText = "sp_validar_login"
cmd.CommandType = 4 

cmd.Parameters.Append cmd.CreateParameter("@email", 200, 1, 100, usuario)
cmd.Parameters.Append cmd.CreateParameter("@senha", 200, 1, 255, senha)

Set rs = cmd.Execute

If rs.EOF Then
    Response.Status = "401 Unauthorized"
    Response.Write "{""erro"":""Usuário ou senha inválidos""}"
    rs.Close: Set rs = Nothing
    connect.Close: Set connect = Nothing
    Response.End
End If


If rs("senha_hash") <> HashSenha(senha) Then 

    Response.Status = "401 Unauthorized"
    Response.Write "{""erro"":""Usuário ou senha inválidos""}"
    rs.Close: Set rs = Nothing
    connect.Close: Set connect = Nothing
    Response.End
End If


payloadJson = "{""usuario_id"":" & rs("usuario_id") & ",""perfil"":""" & rs("perfil") & """,""exp"":" & DateAdd("s", 3600, Now()) & "}"
token = GenerateJWT(payloadJson)

Response.Status = "200 OK"
Response.Write "{""token"":""" & token & """}"

rs.Close: Set rs = Nothing
connect.Close: Set connect = Nothing
%>
