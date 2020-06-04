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
var manage__MMColParam = "admin";
if (String(Request("MM_EmptyValue")) != "undefined" && 
    String(Request("MM_EmptyValue")) != "") { 
  manage__MMColParam = String(Request("MM_EmptyValue"));
}
%>
<%
var manage = Server.CreateObject("ADODB.Recordset");
manage.ActiveConnection = MM_luntan_STRING;
manage.Source = "SELECT * FROM user WHERE user <> '"+ manage__MMColParam.replace(/'/g, "''") + "' ORDER BY admin DESC";
manage.CursorType = 0;
manage.CursorLocation = 2;
manage.LockType = 1;
manage.Open();
var manage_numRows = 0;
%>
<%
var Repeat1__numRows = 10;
var Repeat1__index = 0;
manage_numRows += Repeat1__numRows;
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
// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var manage_total = manage.RecordCount;

// set the number of rows displayed on this page
if (manage_numRows < 0) {            // if repeat region set to all records
  manage_numRows = manage_total;
} else if (manage_numRows == 0) {    // if no repeat regions
  manage_numRows = 1;
}

// set the first and last displayed record
var manage_first = 1;
var manage_last  = manage_first + manage_numRows - 1;

// if we have the correct record count, check the other stats
if (manage_total != -1) {
  manage_numRows = Math.min(manage_numRows, manage_total);
  manage_first   = Math.min(manage_first, manage_total);
  manage_last    = Math.min(manage_last, manage_total);
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
<% var MM_paramName = ""; %>
<%
// *** Move To Record and Go To Record: declare variables

var MM_rs        = manage;
var MM_rsCount   = manage_total;
var MM_size      = manage_numRows;
var MM_uniqueCol = "";
    MM_paramName = "";
var MM_offset = 0;
var MM_atTotal = false;
var MM_paramIsDefined = (MM_paramName != "" && String(Request(MM_paramName)) != "undefined");
%>
<%
// *** Move To Record: handle 'index' or 'offset' parameter

if (!MM_paramIsDefined && MM_rsCount != 0) {

  // use index parameter if defined, otherwise use offset parameter
  r = String(Request("index"));
  if (r == "undefined") r = String(Request("offset"));
  if (r && r != "undefined") MM_offset = parseInt(r);

  // if we have a record count, check if we are past the end of the recordset
  if (MM_rsCount != -1) {
    if (MM_offset >= MM_rsCount || MM_offset == -1) {  // past end or move last
      if ((MM_rsCount % MM_size) != 0) {  // last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount % MM_size);
      } else {
        MM_offset = MM_rsCount - MM_size;
      }
    }
  }

  // move the cursor to the selected record
  for (var i=0; !MM_rs.EOF && (i < MM_offset || MM_offset == -1); i++) {
    MM_rs.MoveNext();
  }
  if (MM_rs.EOF) MM_offset = i;  // set MM_offset to the last possible record
}
%>
<%
// *** Move To Record: if we dont know the record count, check the display range

if (MM_rsCount == -1) {

  // walk to the end of the display range for this page
  for (var i=MM_offset; !MM_rs.EOF && (MM_size < 0 || i < MM_offset + MM_size); i++) {
    MM_rs.MoveNext();
  }

  // if we walked off the end of the recordset, set MM_rsCount and MM_size
  if (MM_rs.EOF) {
    MM_rsCount = i;
    if (MM_size < 0 || MM_size > MM_rsCount) MM_size = MM_rsCount;
  }

  // if we walked off the end, set the offset based on page size
  if (MM_rs.EOF && !MM_paramIsDefined) {
    if ((MM_rsCount % MM_size) != 0) {  // last page not a full repeat region
      MM_offset = MM_rsCount - (MM_rsCount % MM_size);
    } else {
      MM_offset = MM_rsCount - MM_size;
    }
  }

  // reset the cursor to the beginning
  if (MM_rs.CursorType > 0) {
    if (!MM_rs.BOF) MM_rs.MoveFirst();
  } else {
    MM_rs.Requery();
  }

  // move the cursor to the selected record
  for (var i=0; !MM_rs.EOF && i < MM_offset; i++) {
    MM_rs.MoveNext();
  }
}
%>
<%
// *** Move To Record: update recordset stats

// set the first and last displayed record
manage_first = MM_offset + 1;
manage_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  manage_first = Math.min(manage_first, MM_rsCount);
  manage_last  = Math.min(manage_last, MM_rsCount);
}

// set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount != -1 && MM_offset + MM_size >= MM_rsCount);
%>
<%
// *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

// create the list of parameters which should not be maintained
var MM_removeList = "&index=";
if (MM_paramName != "") MM_removeList += "&" + MM_paramName.toLowerCase() + "=";
var MM_keepURL="",MM_keepForm="",MM_keepBoth="",MM_keepNone="";

// add the URL parameters to the MM_keepURL string
for (var items=new Enumerator(Request.QueryString); !items.atEnd(); items.moveNext()) {
  var nextItem = "&" + items.item().toLowerCase() + "=";
  if (MM_removeList.indexOf(nextItem) == -1) {
    MM_keepURL += "&" + items.item() + "=" + Server.URLencode(Request.QueryString(items.item()));
  }
}

// add the Form variables to the MM_keepForm string
for (var items=new Enumerator(Request.Form); !items.atEnd(); items.moveNext()) {
  var nextItem = "&" + items.item().toLowerCase() + "=";
  if (MM_removeList.indexOf(nextItem) == -1) {
    MM_keepForm += "&" + items.item() + "=" + Server.URLencode(Request.Form(items.item()));
  }
}

// create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL + MM_keepForm;
if (MM_keepBoth.length > 0) MM_keepBoth = MM_keepBoth.substring(1);
if (MM_keepURL.length > 0)  MM_keepURL = MM_keepURL.substring(1);
if (MM_keepForm.length > 0) MM_keepForm = MM_keepForm.substring(1);
%>
<%
// *** Move To Record: set the strings for the first, last, next, and previous links

var MM_moveFirst="",MM_moveLast="",MM_moveNext="",MM_movePrev="";
var MM_keepMove = MM_keepBoth;  // keep both Form and URL parameters for moves
var MM_moveParam = "index";

// if the page has a repeated region, remove 'offset' from the maintained parameters
if (MM_size > 1) {
  MM_moveParam = "offset";
  if (MM_keepMove.length > 0) {
    params = MM_keepMove.split("&");
    MM_keepMove = "";
    for (var i=0; i < params.length; i++) {
      var nextItem = params[i].substring(0,params[i].indexOf("="));
      if (nextItem.toLowerCase() != MM_moveParam) {
        MM_keepMove += "&" + params[i];
      }
    }
    if (MM_keepMove.length > 0) MM_keepMove = MM_keepMove.substring(1);
  }
}

// set the strings for the move to links
if (MM_keepMove.length > 0) MM_keepMove = Server.HTMLEncode(MM_keepMove) + "&";
var urlStr = Request.ServerVariables("URL") + "?" + MM_keepMove + MM_moveParam + "=";
MM_moveFirst = urlStr + "0";
MM_moveLast  = urlStr + "-1";
MM_moveNext  = urlStr + (MM_offset + MM_size);
MM_movePrev  = urlStr + Math.max(MM_offset - MM_size,0);
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>用户管理</title>
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
<form name="form1" method="post" action="user.asp">
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
        <td nowrap><div align="left"><a href="/default1.asp">返回&quot;王国之心:记忆之链&quot;游戏<strong>论坛首页</strong>!</a> </div></td>
        <td nowrap><div align="right">| <a href="/index2.htm">进入&quot;王国之心:记忆之链&quot;<strong>游戏主页</strong>!</a></div></td>
      </tr>
    </table>
    <hr width="800">
    <table width="800" border="1" cellpadding="10" cellspacing="0">
      <tr>
        <th colspan="6">用户管理</th>
      </tr>
      <tr>
        <th>用户名</th>
        <th>性别</th>
        <th>E-mail</th>
        <th>管理权限</th>
        <th>&nbsp;</th>
        <th>&nbsp;</th>
      </tr>
      <% while ((Repeat1__numRows-- != 0) && (!manage.EOF)) { %>
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
          </div></td><td><div align="center"><A HREF="/user_manage_updata.asp?user=<%=(manage.Fields.Item("user").Value)%>">修改</A></div></td>
        <td><div align="center"><a href="/user_manage_delete.asp?user=<%=(manage.Fields.Item("user").Value)%>">删除</a></div></td>
      </tr>
      <%
  Repeat1__index++;
  manage.MoveNext();
}
%>
      <tr>
        <td colspan="6"><div align="center">
          <% if (!manage.EOF || !manage.BOF) { %>
          <table border="0" width="50%" align="center">
            <tr>
              <td width="23%" align="center">
                <% if (MM_offset != 0) { %>
                <a href="<%=MM_moveFirst%>">第一页</a>
                <% } // end MM_offset != 0 %>
              </td>
              <td width="31%" align="center">
                <% if (MM_offset != 0) { %>
                <a href="<%=MM_movePrev%>">前一页</a>
                <% } // end MM_offset != 0 %>
              </td>
              <td width="23%" align="center">
                <% if (!MM_atTotal) { %>
                <a href="<%=MM_moveNext%>">下一页</a>
                <% } // end !MM_atTotal %>
              </td>
              <td width="23%" align="center">
                <% if (!MM_atTotal) { %>
                <a href="<%=MM_moveLast%>">最后一页</a>
                <% } // end !MM_atTotal %>
              </td>
            </tr>
          </table>
          <% } // end !manage.EOF || !manage.BOF %>
</div></td>
      </tr>
    </table>
  </div>
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
