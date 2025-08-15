<!--#include virtual="/aspbackend/includes/db_conexao.asp" -->
<!--#include virtual="/aspbackend/includes/utils.asp"-->

<%
Response.CodePage = 65001
Response.Charset = "UTF-8"
Response.ContentType = "application/json"

If Request.ServerVariables("REQUEST_METHOD") <> "GET" Then
    Response.Status = "405 Method Not Allowed"
    Response.AddHeader "Allow", "GET"
    Response.Write "{""erro"":""Protocolo HTTP não permitido para essa rota""}"
    Response.End
End If

Dim cmd, rs, sql, json, first
sql = "EXEC sp_listar_produtos"

Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connect
cmd.CommandText = sql
cmd.CommandType = 1 
Set rs = cmd.Execute

json = "["
first = True
Do Until rs.EOF
    If Not first Then json = json & ","
    json = json & "{""produto_id"":" & rs("produto_id") & _
           ",""nome"":""" & JSONEscape(rs("nome")) & _ 
           """,""descricao"":""" & JSONEscape(rs("descricao")) & _ 
           """,""preco"":" & rs("preco") & _
           ",""quantidade"":" & rs("quantidade") & _
           ",""data_criacao"":""" & rs("data_criacao") & _ 
           """,""data_atualizacao"":""" & rs("data_atualizacao") & """}"
    first = False
    rs.MoveNext
Loop
json = json & "]"

Response.Write json

rs.Close
Set rs = Nothing
Set cmd = Nothing
%>
