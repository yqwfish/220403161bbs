<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/luntan.asp" -->
<%
var zhuti__MMColParam = "交流";
if (String(Request("MM_EmptyValue")) != "undefined" && 
    String(Request("MM_EmptyValue")) != "") { 
  zhuti__MMColParam = String(Request("MM_EmptyValue"));
}
%>
<%
var zhuti = Server.CreateObject("ADODB.Recordset");
zhuti.ActiveConnection = MM_luntan_STRING;
zhuti.Source = "SELECT state FROM zhuti WHERE state = '"+ zhuti__MMColParam.replace(/'/g, "''") + "'";
zhuti.CursorType = 0;
zhuti.CursorLocation = 2;
zhuti.LockType = 1;
zhuti.Open();
var zhuti_numRows = 0;
%>
<%
// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var zhuti_total = zhuti.RecordCount;

// set the number of rows displayed on this page
if (zhuti_numRows < 0) {            // if repeat region set to all records
  zhuti_numRows = zhuti_total;
} else if (zhuti_numRows == 0) {    // if no repeat regions
  zhuti_numRows = 1;
}

// set the first and last displayed record
var zhuti_first = 1;
var zhuti_last  = zhuti_first + zhuti_numRows - 1;

// if we have the correct record count, check the other stats
if (zhuti_total != -1) {
  zhuti_numRows = Math.min(zhuti_numRows, zhuti_total);
  zhuti_first   = Math.min(zhuti_first, zhuti_total);
  zhuti_last    = Math.min(zhuti_last, zhuti_total);
}
%>

<%
// *** Recordset Stats: if we don't know the record count, manually count them

if (zhuti_total == -1) {

  // count the total records by iterating through the recordset
  for (zhuti_total=0; !zhuti.EOF; zhuti.MoveNext()) {
    zhuti_total++;
  }

  // reset the cursor to the beginning
  if (zhuti.CursorType > 0) {
    if (!zhuti.BOF) zhuti.MoveFirst();
  } else {
    zhuti.Requery();
  }

  // set the number of rows displayed on this page
  if (zhuti_numRows < 0 || zhuti_numRows > zhuti_total) {
    zhuti_numRows = zhuti_total;
  }

  // set the first and last displayed record
  zhuti_last  = Math.min(zhuti_first + zhuti_numRows - 1, zhuti_total);
  zhuti_first = Math.min(zhuti_first, zhuti_total);
}
%>
<%
// *** Validate request to log in to this site.
var MM_LoginAction = Request.ServerVariables("URL");
if (Request.QueryString!="") MM_LoginAction += "?" + Server.HTMLEncode(Request.QueryString);
var MM_valUsername=String(Request.Form("user"));
if (MM_valUsername != "undefined") {
  var MM_fldUserAuthorization="admin";
  var MM_redirectLoginSuccess="/ok.asp";
  var MM_redirectLoginFailed="/post_defeat.asp";
  var MM_flag="ADODB.Recordset";
  var MM_rsUser = Server.CreateObject(MM_flag);
  MM_rsUser.ActiveConnection = MM_luntan_STRING;
  MM_rsUser.Source = "SELECT user, password";
  if (MM_fldUserAuthorization != "") MM_rsUser.Source += "," + MM_fldUserAuthorization;
  MM_rsUser.Source += " FROM user WHERE user='" + MM_valUsername.replace(/'/g, "''") + "' AND password='" + String(Request.Form("password")).replace(/'/g, "''") + "'";
  MM_rsUser.CursorType = 0;
  MM_rsUser.CursorLocation = 2;
  MM_rsUser.LockType = 3;
  MM_rsUser.Open();
  if (!MM_rsUser.EOF || !MM_rsUser.BOF) {
    // username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername;
    if (MM_fldUserAuthorization != "") {
      Session("MM_UserAuthorization") = String(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value);
    } else {
      Session("MM_UserAuthorization") = "";
    }
    if (String(Request.QueryString("accessdenied")) != "undefined" && false) {
      MM_redirectLoginSuccess = Request.QueryString("accessdenied");
    }
    MM_rsUser.Close();
    Response.Redirect(MM_redirectLoginSuccess);
  }
  MM_rsUser.Close();
  Response.Redirect(MM_redirectLoginFailed);
}
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>用户登入</title>
<style type="text/css">
<!--
body {
	background-image: url(/images/bg.gif);
	margin-top: 0px;
}
.style2 {color: #FF0000}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>

<body>
<form name="form1" method="POST" action="<%=MM_LoginAction%>">
  <div align="center">
    <div align="center">
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="/images/luntan.jpg" width="320" height="142"></td>
        </tr>
      </table>
      <hr width="800">
      <table width="800" height="30" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <th><div align="left">| 主题留言数:<span class="style2"><%=(zhuti_total)%></span></div></th>
          <td><div align="right">| <a href="/guest_register.asp">注册用户</a> | <a href="/login.asp">用户登入</a>  | <a href="lianxi.htm">联系我们</a> | </div></td>
        </tr>
      </table>
      <hr width="800">
      <table width="800" height="30" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td nowrap><div align="left"><a href="/default.asp">欢迎来到&quot;王国之心:记忆之链&quot;<strong>游戏论坛</strong>!</a> → <a href="/login.asp">用户登入</a></div></td>
          <td nowrap><div align="right">| <a href="/index2.htm">进入&quot;王国之心:记忆之链&quot;<strong>游戏主页</strong>!</a></div></td>
        </tr>
      </table>
      <hr width="800">
      <table width="800" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><div align="center">
              <table width="800" height="198" border="1" cellpadding="10" cellspacing="0">
                <tr bgcolor="#CCFFCC">
                  <th colspan="2"><div align="center">用户登入</div></th>
                </tr>
                <tr bgcolor="#FFFFCC">
                  <td width="291" height="50"><div align="left">请输入您的用户名:</div></td>
                  <td width="463"><div align="left">
                      <input name="user" type="text" id="user" size="30" maxlength="30">
                  </div></td>
                </tr>
                <tr bgcolor="#FFFFCC">
                  <td height="50"><div align="left">请输入您的密码:</div></td>
                  <td><div align="left">
                      <input name="password" type="password" id="password" size="33" maxlength="30">
                  </div></td>
                </tr>
                <tr bgcolor="#FFFFCC">
                  <td colspan="2"><div align="center">
                      <table width="600" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="150"><div align="center"></div></td>
                          <td width="150"><div align="center">
                              <input name="Submit" type="submit" onClick="MM_validateForm('user','','R','password','','R');return document.MM_returnValue" value="登入">
                          </div></td>
                          <td width="150"><div align="center">
                              <input type="reset" name="Submit2" value="重置">
                          </div></td>
                          <td width="150"><div align="center"></div></td>
                        </tr>
                      </table>
                  </div></td>
                </tr>
              </table>
          </div></td>
        </tr>
      </table>
      <hr width="800">
    </div>
  </div>
</form>
</body>
</html>
<%
zhuti.Close();
%>
