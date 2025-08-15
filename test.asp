<%
Response.ContentType = "text/plain"
Response.Write "Versão do ASP: " & Server.ScriptTimeout
Response.Write vbCrLf
Response.Write "Versão do Script Engine: " & ScriptEngineMajorVersion() & "." & ScriptEngineMinorVersion()




%>