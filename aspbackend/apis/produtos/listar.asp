<!--#include virtual="/aspbackend/includes/db_conexao.asp" -->
<!--#include virtual="/aspbackend/includes/utils.asp" -->
<!--#include virtual="/aspbackend/includes/validar_tokenapi.asp" -->

<%
Response.CodePage = 65001
Response.Charset = "UTF-8"
Response.ContentType = "application/json"


' PERMITE APENAS REQUISIÇÕES GET

If Request.ServerVariables("REQUEST_METHOD") <> "GET" Then
    Response.Status = "405 Method Not Allowed"
    Response.AddHeader "Allow", "GET"
    Response.Write "{""erro"":""Metodo HTTP nao permitido""}"
    Response.End
End If

' EXEC PROCEDURE sp_listar_produtos

Dim cmd, rs, json, first
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connect
cmd.CommandText = "sp_listar_produtos"
cmd.CommandType = 4
Set rs = cmd.Execute()

If rs.EOF Then
    Response.Status = "204 No Content"
    Response.Write "{""mensagem"":""Não há produtos cadastrados"",""produtos"":[]}"
Else
    Response.Status = "200 OK"
    json = "["
    first = True
    Do Until rs.EOF
        If Not first Then json = json & ","
        json = json & "{""produto_id"":" & rs("produto_id") & _
               ",""nome"":""" & JSONEscape(rs("nome")) & _
               """,""descricao"":""" & JSONEscape(rs("descricao")) & _
               """,""preco"":" & Replace(rs("preco"), ",", ".") & _
               ",""quantidade"":" & rs("quantidade") & _
               ",""data_criacao"":""" & rs("data_criacao") & _
               """,""data_atualizacao"":""" & rs("data_atualizacao") & """}"
        first = False
        rs.MoveNext
    Loop
    json = json & "]"
    Response.Write json
End If

If Not rs Is Nothing Then rs.Close : Set rs = Nothing
If Not cmd Is Nothing Then Set cmd = Nothing
%>
