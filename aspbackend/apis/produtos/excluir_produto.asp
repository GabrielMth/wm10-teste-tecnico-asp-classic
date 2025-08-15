<%
<!--#include file="../includes/db_conexao.asp"-->
<!--#include file="../includes/utils.asp"-->

If Request.ServerVariables("REQUEST_METHOD") <> "DELETE" Then
    Response.Status = "405 Método Não Permitido"
    Response.AddHeader "Allow", "DELETE"
    Response.ContentType = "application/json"
    Response.Write "{""erro"":""Método não permitido""}"
    Response.End
End If

Dim authHeader, token
authHeader = Request.ServerVariables("HTTP_AUTHORIZATION")

If authHeader = "" Or Left(authHeader, 7) <> "Bearer " Then
    Response.Status = "401 Unauthorized"
    Response.ContentType = "application/json"
    Response.Write "{""erro"":""Token JWT ausente ou inválido""}"
    Response.End
End If

token = Mid(authHeader, 8) 

If Not ValidateJWT(token) Then
    Response.Status = "401 Unauthorized"
    Response.ContentType = "application/json"
    Response.Write "{""erro"":""Token JWT inválido ou expirado""}"
    Response.End
End If

Response.ContentType = "application/json"

Dim produto_id
Dim cmd, rs

produto_id = Request.QueryString("produto_id")

If produto_id = "" Then
    Response.Status = "400 Bad Request"
    Response.Write "{""erro"":""produto_id é obrigatório""}"
    Response.End
End If

Set cmd = Server.CreateObject("ADODB.Command")
Set rs = Server.CreateObject("ADODB.Recordset")

On Error Resume Next
Set cmd.ActiveConnection = connect
cmd.CommandText = "sp_deletar_produto"
cmd.CommandType = 4

cmd.Parameters.Append cmd.CreateParameter("@produto_id", 3, 1, , produto_id)

Set rs = cmd.Execute
If Err.Number <> 0 Then
    Response.Status = "500 Internal Server Error"
    Response.Write "{""erro"":""" & JsonEscape(Err.Description) & """}"
    
    If Not rs Is Nothing Then rs.Close : Set rs = Nothing
    If Not cmd Is Nothing Then Set cmd = Nothing
    If Not connect Is Nothing Then connect.Close : Set connect = Nothing
    Response.End
End If
On Error GoTo 0

Response.Status = "200 OK"
Response.Write "{""mensagem"":""Produto deletado com sucesso"",""produto_id"":" & produto_id & "}"

Set rs = Nothing
Set cmd = Nothing
connect.Close
Set connect = Nothing
%>
