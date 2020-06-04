<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<%
// *** Logout the current user.
MM_Logout = String(Request.ServerVariables("URL")) + "?MM_Logoutnow=1";
if (String(Request("MM_Logoutnow"))=="1") {
  Session.Contents.Remove("MM_Username");
  Session.Contents.Remove("MM_UserAuthorization");
  var MM_logoutRedirectPage = "/default.asp";
  // redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage == "") MM_logoutRedirectPage = String(Request.ServerVariables("URL"));
  if (String(MM_logoutRedirectPage).indexOf("?") == -1 && Request.QueryString != "") {
    var MM_newQS = "?";
    for (var items=new Enumerator(Request.QueryString); !items.atEnd(); items.moveNext()) {
      if (String(items.item()) != "MM_Logoutnow") {
        if (MM_newQS.length > 1) MM_newQS += "&";
        MM_newQS += items.item() + "=" + Server.URLencode(Request.QueryString(items.item()));
      }
    }
    if (MM_newQS.length > 1) MM_logoutRedirectPage += MM_newQS;
  }
  Response.Redirect(MM_logoutRedirectPage);
}
%>
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
// *** Update Record: set variables

if (String(Request("MM_update")) == "form1" &&
    String(Request("MM_recordId")) != "undefined") {

  var MM_editConnection = MM_luntan_STRING;
  var MM_editTable  = "user";
  var MM_editColumn = "user";
  var MM_recordId = "'" + Request.Form("MM_recordId") + "'";
  var MM_editRedirectUrl = "/user_manage.asp";
  var MM_fieldsStr = "admin|value";
  var MM_columnsStr = "admin|none,none,NULL";

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
// *** Update Record: construct a sql update statement and execute it

if (String(Request("MM_update")) != "undefined" &&
    String(Request("MM_recordId")) != "undefined") {

  // create the sql update statement
  MM_editQuery = "update " + MM_editTable + " set ";
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
    MM_editQuery += ((i != 0) ? "," : "") + MM_columns[i] + " = " + formVal;
  }
  MM_editQuery += " where " + MM_editColumn + " = " + MM_recordId;

  if (!MM_abortEdit) {
    // execute the update
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
// *** Restrict Access To Page: Grant or deny access to this page
var MM_authorizedUsers="";
var MM_authFailedURL="/login.asp";
var MM_grantAccess=false;
if (String(Session("MM_Username")) != "undefined") {
  if (true || (String(Session("MM_UserAuthorization"))=="") || (MM_authorizedUsers.indexOf(String(Session("MM_UserAuthorization"))) >=0)) {
    MM_grantAccess = true;
  }
}
if (!MM_grantAccess) {
  var MM_qsChar = "?";
  if (MM_authFailedURL.indexOf("?") >= 0) MM_qsChar = "&";
  var MM_referrer = Request.ServerVariables("URL");
  if (String(Request.QueryString()).length > 0) MM_referrer = MM_referrer + "?" + String(Request.QueryString());
  MM_authFailedURL = MM_authFailedURL + MM_qsChar + "accessdenied=" + Server.URLEncode(MM_referrer);
  Response.Redirect(MM_authFailedURL);
}
%>
<%
var user__MMColParam = "1";
if (String(Session("MM_Username")) != "undefined" && 
    String(Session("MM_Username")) != "") { 
  user__MMColParam = String(Session("MM_Username"));
}
%>
<%
var user = Server.CreateObject("ADODB.Recordset");
user.ActiveConnection = MM_luntan_STRING;
user.Source = "SELECT user, admin FROM user WHERE user = '"+ user__MMColParam.replace(/'/g, "''") + "'";
user.CursorType = 0;
user.CursorLocation = 2;
user.LockType = 1;
user.Open();
var user_numRows = 0;
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
var manage__MMColParam = "1";
if (String(Request.QueryString("user")) != "undefined" && 
    String(Request.QueryString("user")) != "") { 
  manage__MMColParam = String(Request.QueryString("user"));
}
%>
<%
var manage = Server.CreateObject("ADODB.Recordset");
manage.ActiveConnection = MM_luntan_STRING;
manage.Source = "SELECT * FROM user WHERE user = '"+ manage__MMColParam.replace(/'/g, "''") + "'";
manage.CursorType = 0;
manage.CursorLocation = 2;
manage.LockType = 1;
manage.Open();
var manage_numRows = 0;
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
<title>修改管理员权限</title>
<style type="text/css">
<!--
.style1 {color: #FF0000}
body {
	background-image: url(/images/bg.gif);
	margin-top: 0px;
}
-->
</style>
</head>

<body>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
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
        <td><div align="right"> <%=String(Session("MM_Username"))%>:
                <% if (user.Fields.Item("admin").Value == "1"){ %>
            | <a href="/user_manage.asp">用户管理</a>
            <% } %>
            | <a href="/user.asp">个人信息</a> | <a href="/guest_input.asp">发表主题</a> | <a href="lianxi.htm">联系我们</a> | <a href="<%= MM_Logout %>">注销用户</a> | </div></td>
      </tr>
    </table>
    <hr width="800">
    <table width="800" height="30" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><div align="left"><a href="/default1.asp">返回&quot;王国之心:记忆之链&quot;游戏<strong>论坛首页</strong>!</a> </div></td>
        <td><div align="right"><a href="/index2.htm">进入&quot;王国之心:记忆之链&quot;<strong>游戏主页</strong>!</a></div></td>
      </tr>
    </table>
  <hr width="800">
  <table width="800" border="1" cellpadding="10" cellspacing="0">
    <tr>
      <th colspan="4">修改管理员权限</th>
    </tr>
    <tr>
      <th>用户名</th>
      <th>性别</th>
      <th>E-mail</th>
      <th>管理权限</th>
      </tr>
    <tr>
      <td><div align="left"><%=(manage.Fields.Item("user").Value)%></div></td>
      <td><div align="left"><%=(manage.Fields.Item("sex").Value)%></div></td>
      <td><div align="left"><a href="mailto:<%=(manage.Fields.Item("Email").Value)%>?subject=<%=(manage.Fields.Item("user").Value)%>"><%=(manage.Fields.Item("Email").Value)%></a></div></td>
      <td>
            <div align="center">
    <select name="admin" id="admin">
      <option value="1" <%=((1 == (manage.Fields.Item("admin").Value))?"SELECTED":"")%>>管理员</option>
      <option value="0" <%=((0 == (manage.Fields.Item("admin").Value))?"SELECTED":"")%>>普通会员</option>
    </select>
        </div></td>
      </tr>
<tr>
      <td colspan="4"><div align="center">        <table width="600" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="150"><div align="center"></div></td>
            <td width="150"><div align="center">
              <input type="submit" name="Submit" value="确定修改">
            </div></td>
            <td width="150"><div align="center"><a href="/user_manage.asp">取消</a></div></td>
            <td width="150"><div align="center"></div></td>
          </tr>
        </table>
      </div></td>
    </tr>
  </table>
  </div>

  
  

  
  

    <input type="hidden" name="MM_update" value="form1">
  <input type="hidden" name="MM_recordId" value="<%= manage.Fields.Item("user").Value %>">
</form>
</body>
</html>
<%
user.Close();
%>
<%
zhuti.Close();
%>
<%
manage.Close();
%>
