<%
<!--#include file="../includes/db_conexao.asp"-->

Response.ContentType = "application/json"

Dim nome, descricao, preco, quantidade
Dim cmd, rs


nome = Request.Form("nome")
descricao = Request.Form("descricao")
preco = Request.Form("preco")
quantidade = Request.Form("quantidade")


If nome = "" Or preco = "" Or quantidade = "" Then
    Response.Status = "400 Bad Request"
    Response.Write "{""erro"":""Campos obrigatórios ausentes: nome, preco e quantidade""}"
    Response.End
End If

Set cmd = Server.CreateObject("ADODB.Command")
Set rs = Server.CreateObject("ADODB.Recordset")

On Error Resume Next


Set cmd.ActiveConnection = connect
cmd.CommandText = "sp_inserir_produto"
cmd.CommandType = 4 


cmd.Parameters.Append cmd.CreateParameter("@nome", 202, 1, 100, nome) 
cmd.Parameters.Append cmd.CreateParameter("@descricao", 202, 1, 255, descricao)
cmd.Parameters.Append cmd.CreateParameter("@preco", 5, 1, , preco) 
cmd.Parameters.Append cmd.CreateParameter("@quantidade", 3, 1, , quantidade) 


Set rs = cmd.Execute
If Err.Number <> 0 Then
    Response.Status = "500 Internal Server Error"
    Response.Write "{""erro"":""" & Err.Description & """}"
    
    
    If Not rs Is Nothing Then rs.Close : Set rs = Nothing
    If Not cmd Is Nothing Then Set cmd = Nothing
    If Not connect Is Nothing Then connect.Close : Set connect = Nothing
    Response.End
End If
On Error GoTo 0


Response.Status = "201 Created"
Response.Write "{""mensagem"":""Produto inserido com sucesso""," & _
               """produto"":{""nome"":""" & nome & """," & _
               """descricao"":""" & descricao & """," & _
               """preco"":" & preco & "," & _
               """quantidade"":" & quantidade & "}}"

Set rs = Nothing
Set cmd = Nothing
connect.Close
Set connect = Nothing
%>
