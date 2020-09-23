<div align="center">

## Login Authentication with 2 Files\!


</div>

### Description

This simple file (2 files including the text file

of usernames/passwords) allows password

protection of web pages. It was created with 2

thoughts in mind: 1. User does not need access to

the web server the script resides on (NT

authentication is impossible unless you own the

Web Server) 2. Needs no database access.
 
### More Info
 
If you want to protect a page called secured.asp (must be .asp because the code in the include is asp code), then place this include in the FIRST line of the HTML, before the <html> tag: "<!--include file="includelogin.asp"-->" The path is relative to wherever secured.asp exists. Also assumed is a file that holds your usernames / passwords (by default passwords.txt). PLEASE rename this file to something I would never guess!!!

If sucessfull, user is passed through to the original web page requested (along with any query strings applicable to the original page). So, if your original request looked like: http:\\www.somesite.com\sendmail.asp?id=1zx Then once being validated, you will still be passed to the same address, with the "id=1zx" query_string intact. If unsuccessfull, the user stares at a login page forever.

Hairy palms, sweaty feet, & a craving for bologna.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Nik Martin](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nik-martin.md)
**Level**          |Beginner
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Security](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/security__4-14.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/nik-martin-login-authentication-with-2-files__4-6385/archive/master.zip)

### API Declarations

```
This is a compilation of all the
login schemes I found on the
'net that for some reason were
missing some functionality I needed.
```


### Source Code

```
includelogin.asp:
<%
'include this on pages to protect
' (put it before the <html> tag):
'<!--#include file="includelogin.asp"-->
Response.Buffer = True
Function ValidateLogin(sId,sPwd)
	dim FSObject
	dim LoginFile
	Set FSObject = Server.CreateObject("Scripting.FileSystemObject")
	Set LoginFile = FSObject.OpenTextFile(Server.MapPath("passwords.txt"))
	' change this to the path\name
	' of the file that holds your passwords
	'DATA FORMAT IN TEXT FILE: "username<SPACE>password"
	ValidateLogin = False
	WHILE NOT LoginFile.AtEndOfStream 'Scan the text file to determine if the user is legal
		IF LoginFile.ReadLine = sID & " " & sPwd THEN 	'If username AND password are found,
			ValidateLogin = True 			' You passed!
		End If
	WEND
	LoginFile.Close 'Close the text file
	Set LoginFile = Nothing 'free up objects
	Set FSObject = Nothing
End Function
Dim sText
Dim fBack
fBack = False
If Request.Form("dologin") = "yes" Then
	'Try to login
	If ValidateLogin( Request.Form("id"),Request.Form("pwd") ) = True Then
		'It is OK!!!
		'We are logged in so lets go back to the file that included us
		fBack = True
		Session("logonid") = Request.Form("id")
	Else
		sText = "Wrong User ID or Password"
	End If
Else
	'We are not trying to login...
	If Session("logonid") <> "" Then
		'
		fBack = True
		'We are logged in so lets go back to the file that included us
	Else
		sText = "Please login"
	End If
End If
If fBack = False Then %>
	<html>
	<head>
		<meta http-equiv="Content-Language" content="en">
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
		<title>You need to login</title>
	</head>
	<body>
	<%=sText%>
	<%
	Dim sURL
	sURL = Request.ServerVariables("SCRIPT_NAME")
	If Request.ServerVariables("QUERY_STRING") <> "" Then
		sURL = sURL & "?" & Request.ServerVariables("QUERY_STRING")
	End If
	%>
	<form method="POST" action="<%=sURL%>">
	<input type="hidden" name="dologin" value="yes">
 	<table border="0">
 		<tr>
 			<td>User ID:</td>
 			<td><input name="id" size="30"></td>
 		</tr>
 		<tr>
 			<td>Password:</td>
 			<td><input type="password" name="pwd" size="30"></td>
 		</tr>
 	</table>
 	<p><input type="submit" value="Login" name="B1"></p>
	</form>
	</body>
	</html>
<%
Response.End
End If
%>
```

