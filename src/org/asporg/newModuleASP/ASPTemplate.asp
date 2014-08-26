<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Expires = -1

Dim objConn, objRS

Set objRS = Server.CreateObject("ADODB.Recordset")
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Provider="SQLOLEDB"
objConn.Open CONN_DSNLess


Response.Write("HELLO WORLD")

%>