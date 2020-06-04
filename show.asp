<%@LANGUAGE="JAVASCRIPT"%>
<!--#include file="Connections/luntan.asp" -->
<%
var Recordset1__MMColParam = "1";
if (String(Request.QueryString("ID")) != "undefined" && 
    String(Request.QueryString("ID")) != "") { 
  Recordset1__MMColParam = String(Request.QueryString("ID"));
}
%>
<%
var Recordset1 = Server.CreateObject("ADODB.Recordset");
Recordset1.ActiveConnection = MM_luntan_STRING;
Recordset1.Source = "SELECT * FROM zhuti WHERE ID = "+ Recordset1__MMColParam.replace(/'/g, "''") + "";
Recordset1.CursorType = 0;
Recordset1.CursorLocation = 2;
Recordset1.LockType = 1;
Recordset1.Open();
var Recordset1_numRows = 0;
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
var huifu__MMColParam = "1";
if (String(Request.QueryString("ID")) != "undefined" && 
    String(Request.QueryString("ID")) != "") { 
  huifu__MMColParam = String(Request.QueryString("ID"));
}
%>
<%
var huifu = Server.CreateObject("ADODB.Recordset");
huifu.ActiveConnection = MM_luntan_STRING;
huifu.Source = "SELECT * FROM huifu WHERE ID = "+ huifu__MMColParam.replace(/'/g, "''") + "";
huifu.CursorType = 0;
huifu.CursorLocation = 2;
huifu.LockType = 1;
huifu.Open();
var huifu_numRows = 0;
%>
<%
var Repeat1__numRows = 10;
var Repeat1__index = 0;
huifu_numRows += Repeat1__numRows;
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
var huifu_total = huifu.RecordCount;

// set the number of rows displayed on this page
if (huifu_numRows < 0) {            // if repeat region set to all records
  huifu_numRows = huifu_total;
} else if (huifu_numRows == 0) {    // if no repeat regions
  huifu_numRows = 1;
}

// set the first and last displayed record
var huifu_first = 1;
var huifu_last  = huifu_first + huifu_numRows - 1;

// if we have the correct record count, check the other stats
if (huifu_total != -1) {
  huifu_numRows = Math.min(huifu_numRows, huifu_total);
  huifu_first   = Math.min(huifu_first, huifu_total);
  huifu_last    = Math.min(huifu_last, huifu_total);
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

var MM_rs        = huifu;
var MM_rsCount   = huifu_total;
var MM_size      = huifu_numRows;
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
huifu_first = MM_offset + 1;
huifu_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  huifu_first = Math.min(huifu_first, MM_rsCount);
  huifu_last  = Math.min(huifu_last, MM_rsCount);
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>欢迎进入&quot;王国之心:记忆之链&quot;游戏论坛!</title>
<style type="text/css">
<!--
body {
	background-image: url(/images/bg.gif);
	margin-top: 0px;
}
.style1 {color: #FF0000}
-->
</style></head>

<body>
<form name="form1" method="post" action="show.asp">
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
          <th><div align="left">| 主题留言数:<span class="style1"><%=(zhuti_total)%></span></div></th>
          <td><div align="right">| <a href="/guest_register.asp">注册用户</a> | <a href="/login.asp">用户登入</a> | <a href="lianxi.htm">联系我们</a> |  </div></td>
        </tr>
      </table>
      <hr width="800">
      <table width="800" height="30" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td nowrap><div align="left"><a href="/default.asp">返回<strong>论坛首页</strong>!</a> → <a href="/show.asp?ID=<%=(Recordset1.Fields.Item("ID").Value)%>"><%=(Recordset1.Fields.Item("title").Value)%></a></div></td>
          <td nowrap><div align="right">| <a href="/index2.htm">进入&quot;王国之心:记忆之链&quot;<strong>游戏主页</strong>!</a></div></td>
        </tr>
      </table>
      <hr width="800">
      <table width="800" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><div align="center">
            <table width="800" border="1" cellpadding="10" cellspacing="0">
              <tr bgcolor="#CCFFCC">
                <th colspan="2">(<%=(Recordset1.Fields.Item("state").Value)%>)留言主题:<%=(Recordset1.Fields.Item("title").Value)%></th>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td width="119" rowspan="2"><div align="center"><%=(Recordset1.Fields.Item("user").Value)%></div></td>
                <td width="640"><div align="left">
                  <table width="640" height="20" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td><div align="left">发表时间:<%=(Recordset1.Fields.Item("date").Value)%> | <%=(Recordset1.Fields.Item("time").Value)%></div></td>
                      <td><div align="right">| <a href="mailto:<%=(Recordset1.Fields.Item("Email").Value)%>?subject=<%=(Recordset1.Fields.Item("user").Value)%>">
                        Email</a> | <a href="/error.asp">回复</a> |</div></td>
                    </tr>
                  </table>
                </div></td>
              </tr>
              <tr>
                <td height="150" bgcolor="#FFFFFF">
                  <table width="640" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="80"><div align="left"><%=(Recordset1.Fields.Item("content").Value)%></div></td>
                    </tr>
                    <tr>
                      <td><div align="left">
                            <% if ((Recordset1.Fields.Item("signature").Value) != "Null"){ %>
                        </div>
                      <hr align="left">
                        <div align="left"><%=(Recordset1.Fields.Item("signature").Value)%>
                            <% } %>
                        </div></td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            
            <% while ((Repeat1__numRows-- != 0) && (!huifu.EOF)) { %>
            <table width="800" border="1" cellpadding="10" cellspacing="0">
                <tr bgcolor="#FFFFCC">
                  <td width="114" rowspan="2"><div align="center"><%=(huifu.Fields.Item("user").Value)%></div></td>
                  <td width="640"><div align="left">
                    <table width="640" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><div align="left">发表时间:<%=(huifu.Fields.Item("date").Value)%> | <%=(huifu.Fields.Item("time").Value)%></div></td>
                        <td><div align="right">| <a href="mailto:<%=(huifu.Fields.Item("Email").Value)%>?subject=<%=(huifu.Fields.Item("user").Value)%>">Email</a> | <a href="/post_error.asp">回复</a> |</div></td>
                      </tr>
                    </table>
                  </div></td>
                </tr>
                <tr>
                  <td height="100" bgcolor="#FFFFCC">
                    <table width="640" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td height="80"><div align="left"><%=(huifu.Fields.Item("reple").Value)%></div></td>
                      </tr>
                      <tr>
                        <td><div align="left">
                            <% if ((huifu.Fields.Item("signature").Value)!="Null"){ %>
                        </div>
                          <hr align="left">
                          <div align="left"><%=(huifu.Fields.Item("signature").Value)%>
                            <% } %>
                        </div></td>
                      </tr>
                    </table></td>
                </tr>
            </table>
            
            <%
  Repeat1__index++;
  huifu.MoveNext();
}
%>
</div></td>
        </tr>
      </table>
      ,
      <% if (!huifu.EOF || !huifu.BOF) { %>
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
      <% } // end !huifu.EOF || !huifu.BOF %>

      <hr width="800">
</div>
  </div>
</form>
</body>
</html>
<%
Recordset1.Close();
%>
<%
zhuti.Close();
%>
<%
huifu.Close();
%>
