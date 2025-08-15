<%
<!--#include file="../includes/db_conexao.asp"-->
<!--#include file="../includes/utils.asp"-->

Response.ContentType = "application/json"

Dim cmd, rs, json, first
Set cmd = Server.CreateObject("ADODB.Command")
Set rs = Server.CreateObject("ADODB.Recordset")

Dim filtro_nome, preco_min, preco_max, quantidade_min, quantidade_max
filtro_nome = Request.QueryString("nome")
preco_min = Request.QueryString("preco_min")
preco_max = Request.QueryString("preco_max")
quantidade_min = Request.QueryString("quantidade_min")
quantidade_max = Request.QueryString("quantidade_max")

On Error Resume Next

Set cmd.ActiveConnection = connect
cmd.CommandText = "sp_listar_produtos_filtrados"
cmd.CommandType = 4 

cmd.Parameters.Append cmd.CreateParameter("@filtro_nome", 202, 1, 100, IIf(filtro_nome="", Null, filtro_nome))
cmd.Parameters.Append cmd.CreateParameter("@preco_min", 5, 1, , IIf(preco_min="", Null, preco_min))
cmd.Parameters.Append cmd.CreateParameter("@preco_max", 5, 1, , IIf(preco_max="", Null, preco_max))
cmd.Parameters.Append cmd.CreateParameter("@quantidade_min", 3, 1, , IIf(quantidade_min="", Null, quantidade_min))
cmd.Parameters.Append cmd.CreateParameter("@quantidade_max", 3, 1, , IIf(quantidade_max="", Null, quantidade_max))

Set rs = cmd.Execute
If Err.Number <> 0 Then
    Response.Status = "500 Internal Server Error"
    Response.Write "{""erro"":""" & JSONEscape(Err.Description) & """}"

    If Not rs Is Nothing Then rs.Close : Set rs = Nothing
    If Not cmd Is Nothing Then Set cmd = Nothing
    If Not connect Is Nothing Then connect.Close : Set connect = Nothing

    Response.End
End If
On Error GoTo 0

If rs.EOF Then
    Response.Status = "404 Not Found"
    Response.Write "{""mensagem"":""Nenhum produto encontrado com os filtros fornecidos"",""produtos"":[]}"
Else
    Response.Status = "200 OK"
    json = "["
    first = True

    Do Until rs.EOF
        If Not first Then json = json & "," Else first = False
        json = json & "{"
        json = json & """produto_id"":" & rs("produto_id") & ","
        json = json & """nome"":""" & JSONEscape(rs("nome")) & ""","
        json = json & """descricao"":""" & JSONEscape(rs("descricao")) & ""","
        json = json & """preco"":" & rs("preco") & ","
        json = json & """quantidade"":" & rs("quantidade") & ","
        json = json & """data_criacao"":""" & rs("data_criacao") & ""","
        json = json & """data_atualizacao"":""" & rs("data_atualizacao") & """"
        json = json & "}"
        rs.MoveNext
    Loop

    json = json & "]"
    Response.Write "{""produtos"":" & json & "}"
End If

rs.Close
Set rs = Nothing
Set cmd = Nothing
connect.Close
Set connect = Nothing
%>
