<%@ Language="VBScript" %>
<!-- #include file="config.inc" -->

<script language="JScript" runat="server" src="json2.js"></script>

<%

AUTHORIZATION_CODE = Request.querystring( "code" )

access_token = GetAccessToken(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI, AUTHORIZATION_CODE)

set profile = GetUserProfile(access_token)

'Here, you should save the profile in the session or somewhere'

Response.Write "UserID = " & profile.user_id

Function GetUserProfile(access_token)

  Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")

  http.open "GET", "https://tutorials.auth0.com/userinfo?access_token=" & access_token, False

  http.send

  profile = http.responseText

  Set GetUserProfile = JSON.parse(profile)

  Set http = Nothing

End Function


Function GetAccessToken(client_id, client_secret, redirect_uri, authorization_code)

  Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")

  http.open "POST", "https://tutorials.auth0.com/oauth/token", False

  http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

  http.send "client_id=" & client_id & "&client_secret=" & client_secret & "&redirect_uri=" & server.UrlEncode(redirect_uri) & "&code=" & authorization_code & "&grant_type=authorization_code"

  result = http.responseText

  Set http = Nothing

  set jsonResult = JSON.parse(result)

  GetAccessToken = jsonresult.access_token

End Function

%>
