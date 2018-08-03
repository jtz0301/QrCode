<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no" name="viewport">
</head>
<body>
<img id="authCode" onclick="changeCode()" src="/QrCode/code/code" alt="验证码" title="点击更换" />

<script>
function changeCode() {
    document.getElementById("authCode").src = "/QrCode/code/code?t=" + genTimestamp();
}
function genTimestamp() {
    var time = new Date();
    return time.getTime();
}
</script>
</body>
</html>