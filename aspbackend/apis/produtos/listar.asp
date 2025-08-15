<%
<!--#include file="../includes/db_conexao.asp"-->
<!--#include file="../includes/utils.asp"-->

Response.ContentType = "application/json"

Dim rs, json, first
Set rs = Server.CreateObject("ADODB.Recordset")

' Tratamento de erro na execução da procedure
On Error Resume Next
rs.Open "sp_listar_produtos", connect, 1, 3
If Err.Number <> 0 Then
    Response.Status = "500 Internal Server Error"
    Response.Write "{""erro"":""" & JSONEscape(Err.Description) & """}"

    ' Fechando objetos antes de encerrar
    If Not rs Is Nothing Then rs.Close : Set rs = Nothing
    If Not connect Is Nothing Then connect.Close : Set connect = Nothing

    Response.End
End If
On Error GoTo 0

If rs.EOF Then
    Response.Status = "200 OK"
    Response.Write "{""mensagem"":""Não contém produtos cadastrados"",""produtos"":[]}"
Else
    Response.Status = "200 OK"
    json = "["
    first = True

    Do Until rs.EOF
        If Not first Then
            json = json & ","
        Else
            first = False
        End If

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
connect.Close
Set connect = Nothing
%>
