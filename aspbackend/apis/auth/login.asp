<!--#include virtual="/aspbackend/includes/db_conexao.asp" -->
<!--#include virtual="/aspbackend/includes/utils.asp" -->

<%
Response.AddHeader "Content-Type", "application/json; charset=utf-8"
Response.CodePage = 65001
Response.Charset = "UTF-8"

If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Status = "405 Method Not Allowed"
    Response.AddHeader "Allow", "POST"
    Response.Write "{""erro"":""Método HTTP não permitido""}"
    Response.End
End If

Dim body
If Request.TotalBytes > 0 Then
    Dim binData, stream
    binData = Request.BinaryRead(Request.TotalBytes)
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write binData
    stream.Position = 0
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    body = stream.ReadText
    stream.Close
    Set stream = Nothing
Else
    body = ""
End If

Dim data, email, senha
Set data = ParseSimpleJSON(body)

If Not IsObject(data) Then
    Response.Status = "400 Bad Request"
    Response.Write "{""erro"":""JSON inválido""}"
    Response.End
End If

email = ""
senha = ""

If data.Exists("email") Then email = Trim(data("email"))
If data.Exists("senha") Then senha = Trim(data("senha"))

If email = "" Or senha = "" Then
    Response.Status = "400 Bad Request"
    Response.Write "{""erro"":""Email e senha são obrigatórios""}"
    Response.End
End If

Dim cmdLogin, rsLogin
Set cmdLogin = Server.CreateObject("ADODB.Command")
cmdLogin.ActiveConnection = connect
cmdLogin.CommandText = "sp_validar_login"
cmdLogin.CommandType = 4
cmdLogin.Parameters.Append cmdLogin.CreateParameter("@email", 200, 1, 100, email)
cmdLogin.Parameters.Append cmdLogin.CreateParameter("@senha", 200, 1, 255, senha)
Set rsLogin = cmdLogin.Execute()

If rsLogin.EOF Then
    Response.Status = "401 Unauthorized"
    Response.Write "{""erro"":""Credenciais inválidas""}"
    Response.End
End If

Dim usuario_id, nome_usuario, perfil_usuario, criado_em
usuario_id = rsLogin("usuario_id")
nome_usuario = rsLogin("nome_usuario")
perfil_usuario = rsLogin("perfil")


Dim token, expire_date
token = GenerateRandomToken()
expire_date = DateAdd("h", 2, Now()) 


Dim cmdToken
Set cmdToken = Server.CreateObject("ADODB.Command")
cmdToken.ActiveConnection = connect
cmdToken.CommandText = "sp_inserir_token"
cmdToken.CommandType = 4
cmdToken.Parameters.Append cmdToken.CreateParameter("@usuario_id", 3, 1, , usuario_id)
cmdToken.Parameters.Append cmdToken.CreateParameter("@token", 200, 1, 255, token)
cmdToken.Parameters.Append cmdToken.CreateParameter("@expire_date", 135, 1, , expire_date)
cmdToken.Execute
Set cmdToken = Nothing

Response.Status = "200 OK"
jsonResponse = "{""perfil_usuario"":""" & perfil_usuario & """," & _
               """usuario"":{" & _
                   """id"":" & usuario_id & "," & _
                   """nome"":""" & JsonEscape(nome_usuario) & """," & _
                   """token"":""" & (token) & """" & _
               "}" & _
               "}"

Set rsLogin = Nothing
Response.Write jsonResponse
Response.End
%>
