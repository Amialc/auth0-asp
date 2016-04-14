<%@ Language="VBScript" %>
<% Option Explicit %>
<!-- #include file="config.inc" -->
<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Testing Auth0 with Classic ASP</title>
</head>
<body>
  <script src="https://cdn.auth0.com/js/lock-9.0.min.js"></script>
  <script type="text/javascript">
    var lock = new Auth0Lock('<% response.write CLIENT_ID %>', '<% response.write DOMAIN %>');

    function signin() {
      lock.show({
        callbackURL: 'http://localhost/auth0/callback.asp'
      });
    }
  </script>
  <button onclick="signin()">Login</button>
</body>
</html>
