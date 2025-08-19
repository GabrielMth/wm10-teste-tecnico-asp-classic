<!--#include virtual="/aspbackend/includes/utils.asp" -->
    
<%
' VERIFICA SE O TOKEN ESTÁ NO HEADER
Dim authHeader, token, usuario_id, is_valid
authHeader = Request.ServerVariables("HTTP_AUTHORIZATION")

If authHeader = "" Or InStr(authHeader, "Bearer ") = 0 Then
    Response.Status = "401 Unauthorized"
    Response.Write "{""erro"":""Token nao informado""}"
    Response.End
End If

token = Trim(Replace(authHeader, "Bearer", ""))

' Valida o token com procedure sp_validar_token
Dim cmdValidar
Set cmdValidar = Server.CreateObject("ADODB.Command")
cmdValidar.ActiveConnection = connect
cmdValidar.CommandText = "sp_validar_token"
cmdValidar.CommandType = 4 

cmdValidar.Parameters.Append cmdValidar.CreateParameter("@token", 200, 1, 255, token)
cmdValidar.Parameters.Append cmdValidar.CreateParameter("@usuario_id", 3, 2) ' Output
cmdValidar.Parameters.Append cmdValidar.CreateParameter("@is_valid", 11, 2)  ' Output

cmdValidar.Execute

usuario_id = cmdValidar.Parameters("@usuario_id").Value
is_valid = cmdValidar.Parameters("@is_valid").Value

Set cmdValidar = Nothing

If Not is_valid Then
    Response.Status = "401 Unauthorized"
    Response.Write "{""erro"":""Token expirado ou inválido""}"
    Response.End
End If
%>
