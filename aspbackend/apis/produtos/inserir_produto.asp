<%
<!--#include file="../includes/db_conexao.asp"-->

Response.ContentType = "application/json"

Dim jsonRequest, nome, descricao, preco, quantidade
Dim rs, sql


jsonRequest = Request.Form
If jsonRequest = "" Then
    Response.Status = "400 Bad Request"
    Response.Write "{""erro"":""Nenhum dado recebido""}"
    Response.End
End If

nome = Request.Form("nome")
descricao = Request.Form("descricao")
preco = Request.Form("preco")
quantidade = Request.Form("quantidade")


If nome = "" Or preco = "" Or quantidade = "" Then
    Response.Status = "400 Bad Request"
    Response.Write "{""erro"":""Campos obrigatórios ausentes""}"
    Response.End
End If

Set rs = Server.CreateObject("ADODB.Recordset")

On Error Resume Next
sql = "EXEC sp_inserir_produto @nome = '" & Replace(nome,"'","''") & "', " & _
      "@descricao = '" & Replace(descricao,"'","''") & "', " & _
      "@preco = " & preco & ", " & _
      "@quantidade = " & quantidade

rs.Open sql, connect, 1, 3

If Err.Number <> 0 Then
    Response.Status = "500 Internal Server Error"
    Response.Write "{""erro"":""" & Err.Description & """}"
    
    rs.Close
    Set rs = Nothing
    connect.Close
    Set connect = Nothing
    Response.End
End If
On Error GoTo 0

Response.Status = "201 Created"
Response.Write "{""mensagem"":""Produto inserido com sucesso""," & _
               """nome"":""" & nome & """," & _
               """descricao"":""" & descricao & """," & _
               """preco"":" & preco & "," & _
               """quantidade"":" & quantidade & "}"

rs.Close
Set rs = Nothing
connect.Close
Set connect = Nothing
%>
