<%
<!--#include file="../includes/db_conexao.asp"-->
<!--#include file="../includes/utils.asp"-->

If Request.ServerVariables("REQUEST_METHOD") <> "PUT" Then
    Response.Status = "405 Método Não Permitido"
    Response.AddHeader "Allow", "PUT"
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

Dim produto_id, nome, descricao, preco, quantidade
produto_id = Request.Form("produto_id")
nome = Request.Form("nome")
descricao = Request.Form("descricao")
preco = Request.Form("preco")
quantidade = Request.Form("quantidade")


If produto_id = "" Or nome = "" Or preco = "" Or quantidade = "" Then
    Response.Status = "400 Bad Request"
    Response.Write "{""erro"":""Campos obrigatórios ausentes: produto_id, nome, preco ou quantidade""}"
    Response.End
End If

Dim cmd, rs
Set cmd = Server.CreateObject("ADODB.Command")
Set rs = Server.CreateObject("ADODB.Recordset")

On Error Resume Next
Set cmd.ActiveConnection = connect
cmd.CommandText = "sp_atualizar_produto"
cmd.CommandType = 4 

cmd.Parameters.Append cmd.CreateParameter("@produto_id", 3, 1, , produto_id)
cmd.Parameters.Append cmd.CreateParameter("@nome", 202, 1, 100, nome)
cmd.Parameters.Append cmd.CreateParameter("@descricao", 202, 1, 255, descricao)
cmd.Parameters.Append cmd.CreateParameter("@preco", 5, 1, , preco)
cmd.Parameters.Append cmd.CreateParameter("@quantidade", 3, 1, , quantidade)

cmd.Execute

If Err.Number <> 0 Then
    Response.Status = "500 Internal Server Error"
    Response.Write "{""erro"":""" & JSONEscape(Err.Description) & """}"
    
    
    Set rs = Nothing
    Set cmd = Nothing
    connect.Close
    Set connect = Nothing
    Response.End
End If
On Error GoTo 0

Response.Status = "200 OK"
Response.Write "{""mensagem"":""Produto atualizado com sucesso""," & _
               """produto_id"":" & produto_id & "," & _
               """nome"":""" & JSONEscape(nome) & """," & _
               """descricao"":""" & JSONEscape(descricao) & """," & _
               """preco"":" & preco & "," & _
               """quantidade"":" & quantidade & "}"

Set rs = Nothing
Set cmd = Nothing
connect.Close
Set connect = Nothing
%>
