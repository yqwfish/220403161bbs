<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/luntan.asp" -->
<%
// *** Edit Operations: declare variables

// set the form action variable
var MM_editAction = Request.ServerVariables("SCRIPT_NAME");
if (Request.QueryString) {
  MM_editAction += "?" + Server.HTMLEncode(Request.QueryString);
}

// boolean to abort record edit
var MM_abortEdit = false;

// query string to execute
var MM_editQuery = "";
%>
<%
// *** Redirect if username exists
var MM_flag="MM_insert";
if (String(Request(MM_flag)) != "undefined") {
  var MM_dupKeyRedirect="/error.asp";
  var MM_rsKeyConnection=MM_luntan_STRING;
  var MM_dupKeyUsernameValue = String(Request.Form("user"));
  var MM_dupKeySQL = "SELECT user FROM user WHERE user='" + MM_dupKeyUsernameValue + "'"
  var MM_adodbRecordset = "ADODB.Recordset";
  var MM_rsKey = Server.CreateObject(MM_adodbRecordset);
  MM_rsKey.ActiveConnection = MM_rsKeyConnection;
  MM_rsKey.Source = MM_dupKeySQL;
  MM_rsKey.CursorType=0;
  MM_rsKey.CursorLocation=2;
  MM_rsKey.LockType=3;
  MM_rsKey.Open();
  if (!MM_rsKey.EOF || !MM_rsKey.BOF) {
    // the username was found - can not add the requested username
    var MM_qsChar = "?";
    if (MM_dupKeyRedirect.indexOf("?") >= 0) MM_qsChar = "&";
    MM_dupKeyRedirect = MM_dupKeyRedirect + MM_qsChar + "requsername=" + MM_dupKeyUsernameValue;
    Response.Redirect(MM_dupKeyRedirect);
  }
  MM_rsKey.Close();
}
%>
<%
// *** Insert Record: set variables

if (String(Request("MM_insert")) == "form1") {

  var MM_editConnection = MM_luntan_STRING;
  var MM_editTable  = "user";
  var MM_editRedirectUrl = "/guest_register_ok.asp";
  var MM_fieldsStr = "user|value|password|value|sex|value|Email|value|signature|value";
  var MM_columnsStr = "user|',none,''|password|',none,''|sex|',none,''|Email|',none,''|signature|',none,''";

  // create the MM_fields and MM_columns arrays
  var MM_fields = MM_fieldsStr.split("|");
  var MM_columns = MM_columnsStr.split("|");
  
  // set the form values
  for (var i=0; i+1 < MM_fields.length; i+=2) {
    MM_fields[i+1] = String(Request.Form(MM_fields[i]));
  }

  // append the query string to the redirect URL
  if (MM_editRedirectUrl && Request.QueryString && Request.QueryString.Count > 0) {
    MM_editRedirectUrl += ((MM_editRedirectUrl.indexOf('?') == -1)?"?":"&") + Request.QueryString;
  }
}
%>
<%
// *** Insert Record: construct a sql insert statement and execute it

if (String(Request("MM_insert")) != "undefined") {

  // create the sql insert statement
  var MM_tableValues = "", MM_dbValues = "";
  for (var i=0; i+1 < MM_fields.length; i+=2) {
    var formVal = MM_fields[i+1];
    var MM_typesArray = MM_columns[i+1].split(",");
    var delim =    (MM_typesArray[0] != "none") ? MM_typesArray[0] : "";
    var altVal =   (MM_typesArray[1] != "none") ? MM_typesArray[1] : "";
    var emptyVal = (MM_typesArray[2] != "none") ? MM_typesArray[2] : "";
    if (formVal == "" || formVal == "undefined") {
      formVal = emptyVal;
    } else {
      if (altVal != "") {
        formVal = altVal;
      } else if (delim == "'") { // escape quotes
        formVal = "'" + formVal.replace(/'/g,"''") + "'";
      } else {
        formVal = delim + formVal + delim;
      }
    }
    MM_tableValues += ((i != 0) ? "," : "") + MM_columns[i];
    MM_dbValues += ((i != 0) ? "," : "") + formVal;
  }
  MM_editQuery = "insert into " + MM_editTable + " (" + MM_tableValues + ") values (" + MM_dbValues + ")";

  if (!MM_abortEdit) {
    // execute the insert
    var MM_editCmd = Server.CreateObject('ADODB.Command');
    MM_editCmd.ActiveConnection = MM_editConnection;
    MM_editCmd.CommandText = MM_editQuery;
    MM_editCmd.Execute();
    MM_editCmd.ActiveConnection.Close();

    if (MM_editRedirectUrl) {
      Response.Redirect(MM_editRedirectUrl);
    }
  }

}
%>
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>注册新用户</title>
<style type="text/css">
<!--
body {
	background-image: url(/images/bg.gif);
	margin-top: 0px;
}
.style1 {color: #FF0000}
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
  } if (MM_findObj('password').value!=MM_findObj('password2').value)errors +='\请确认你两次输入的密码一致! \n'; 
  if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
</head>

<body>
<form name="form1" method="POST" action="<%=MM_editAction%>">
  <div align="center">
    <table border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="/images/luntan.jpg" width="320" height="142"></td>
      </tr>
    </table>
    <hr width="800">
    <table width="800" height="30" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <th><div align="left">| 主题留言数:<span class="style1"><%=(zhuti_total)%></span></div></th>
        <td><div align="right">| <a href="/guest_register.asp">注册用户</a> | <a href="/login.asp">用户登入</a> | <a href="lianxi.htm">联系我们</a> |  </div></td>
      </tr>
    </table>
    <hr width="800">
    <table width="800" height="30" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td nowrap><div align="left"><a href="/default.asp">欢迎来到&quot;王国之心:记忆之链&quot;<strong>游戏论坛</strong>!</a> → <a href="/guest_register.asp">注册新用户</a></div></td>
        <td nowrap><div align="right">| <a href="/index2.htm">进入&quot;王国之心:记忆之链&quot;<strong>游戏主页</strong>!</a></div></td>
      </tr>
    </table>
    <hr width="800">
    <table width="800" height="297" border="1" cellpadding="10">
      <tr bgcolor="#CCFFCC">
        <th colspan="2"><div align="center">注册新用户</div></th>
      </tr>
      <tr bgcolor="#FFFFCC">
        <td><div align="left">用户名(*):<br>
  注册用户名，长度限制为1~30字符</div></td>
        <td><div align="left">
          <input name="user" type="text" id="user" size="30" maxlength="30">
        </div></td>
      </tr>
      <tr bgcolor="#FFFFCC">
        <td><div align="left">密码(*):<br>
  密码区分大小写，长度限制为1~30字符</div></td>
        <td><div align="left">
          <input name="password" type="password" id="password" size="33" maxlength="30">
        </div></td>
      </tr>
      <tr bgcolor="#FFFFCC">
        <td><div align="left">确认密码(*):<br>
            重新输入密码          <br>
        </div></td>
        <td><div align="left">
          <input name="password2" type="password" id="password2" size="33" maxlength="30">
        </div></td>
      </tr>
      <tr bgcolor="#FFFFCC">
        <td><div align="left">性别:<br>
  请选择您的性别</div></td>
        <td><div align="left">
          <select name="sex" id="sex">
              <option value="保密" selected>保密</option>
              <option value="男">男</option>
              <option value="女">女</option>
          </select>
        </div></td>
      </tr>
      <tr bgcolor="#FFFFCC">
        <td><div align="left">电子邮箱:<br>
  请输入E-mail地址，以方便我们联系您</div></td>
        <td><div align="left">
          <input name="Email" type="text" id="Email" size="40" maxlength="255">
        </div></td>
      </tr>
      <tr bgcolor="#FFFFCC">
        <td><div align="left">个性签名:<br>
  请输入您的个性签名</div></td>
        <td><div align="left">
          <textarea name="signature" cols="50" rows="4" id="signature"></textarea>
        </div></td>
      </tr>
      <tr bgcolor="#FFFFCC">
        <td colspan="2"><div align="left">有加(*)号的为必填项</div></td>
      </tr>
      <tr bgcolor="#FFFFCC">
        <td colspan="2"><div align="center">
          <table width="600" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="150"><div align="center"></div></td>
              <td width="150"><div align="center">
                <input name="Submit" type="submit" onClick="MM_validateForm('user','','R','Email','','NisEmail','password','','R');return document.MM_returnValue" value="提交">
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
  </div>

  <input type="hidden" name="MM_insert" value="form1">
</form>
</body>
</html>
<%
zhuti.Close();
%>
