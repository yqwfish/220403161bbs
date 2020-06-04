<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/luntan.asp" -->
<%
var gonggao__MMColParam = "公告";
if (String(Request("MM_EmptyValue")) != "undefined" && 
    String(Request("MM_EmptyValue")) != "") { 
  gonggao__MMColParam = String(Request("MM_EmptyValue"));
}
%>
<%
var gonggao = Server.CreateObject("ADODB.Recordset");
gonggao.ActiveConnection = MM_luntan_STRING;
gonggao.Source = "SELECT * FROM zhuti WHERE state = '"+ gonggao__MMColParam.replace(/'/g, "''") + "' ORDER BY ID DESC";
gonggao.CursorType = 0;
gonggao.CursorLocation = 2;
gonggao.LockType = 1;
gonggao.Open();
var gonggao_numRows = 0;
%>
<%
var jiaoliou__MMColParam = "交流";
if (String(Request("MM_EmptyValue")) != "undefined" && 
    String(Request("MM_EmptyValue")) != "") { 
  jiaoliou__MMColParam = String(Request("MM_EmptyValue"));
}
%>
<%
var jiaoliou = Server.CreateObject("ADODB.Recordset");
jiaoliou.ActiveConnection = MM_luntan_STRING;
jiaoliou.Source = "SELECT * FROM zhuti WHERE state = '"+ jiaoliou__MMColParam.replace(/'/g, "''") + "' ORDER BY ID DESC";
jiaoliou.CursorType = 0;
jiaoliou.CursorLocation = 2;
jiaoliou.LockType = 1;
jiaoliou.Open();
var jiaoliou_numRows = 0;
%>
<%
var Repeat1__numRows = -1;
var Repeat1__index = 0;
gonggao_numRows += Repeat1__numRows;
%>
<%
var Repeat2__numRows = 10;
var Repeat2__index = 0;
jiaoliou_numRows += Repeat2__numRows;
%>
<%
// *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

// set the record count
var jiaoliou_total = jiaoliou.RecordCount;

// set the number of rows displayed on this page
if (jiaoliou_numRows < 0) {            // if repeat region set to all records
  jiaoliou_numRows = jiaoliou_total;
} else if (jiaoliou_numRows == 0) {    // if no repeat regions
  jiaoliou_numRows = 1;
}

// set the first and last displayed record
var jiaoliou_first = 1;
var jiaoliou_last  = jiaoliou_first + jiaoliou_numRows - 1;

// if we have the correct record count, check the other stats
if (jiaoliou_total != -1) {
  jiaoliou_numRows = Math.min(jiaoliou_numRows, jiaoliou_total);
  jiaoliou_first   = Math.min(jiaoliou_first, jiaoliou_total);
  jiaoliou_last    = Math.min(jiaoliou_last, jiaoliou_total);
}
%>
<%
// *** Recordset Stats: if we don't know the record count, manually count them

if (jiaoliou_total == -1) {

  // count the total records by iterating through the recordset
  for (jiaoliou_total=0; !jiaoliou.EOF; jiaoliou.MoveNext()) {
    jiaoliou_total++;
  }

  // reset the cursor to the beginning
  if (jiaoliou.CursorType > 0) {
    if (!jiaoliou.BOF) jiaoliou.MoveFirst();
  } else {
    jiaoliou.Requery();
  }

  // set the number of rows displayed on this page
  if (jiaoliou_numRows < 0 || jiaoliou_numRows > jiaoliou_total) {
    jiaoliou_numRows = jiaoliou_total;
  }

  // set the first and last displayed record
  jiaoliou_last  = Math.min(jiaoliou_first + jiaoliou_numRows - 1, jiaoliou_total);
  jiaoliou_first = Math.min(jiaoliou_first, jiaoliou_total);
}
%>
<% var MM_paramName = ""; %>
<%
// *** Move To Record and Go To Record: declare variables

var MM_rs        = jiaoliou;
var MM_rsCount   = jiaoliou_total;
var MM_size      = jiaoliou_numRows;
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
jiaoliou_first = MM_offset + 1;
jiaoliou_last  = MM_offset + MM_size;
if (MM_rsCount != -1) {
  jiaoliou_first = Math.min(jiaoliou_first, MM_rsCount);
  jiaoliou_last  = Math.min(jiaoliou_last, MM_rsCount);
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
<title>欢迎来到&quot;王国之心:记忆之链&quot;游戏论坛!</title>
<style type="text/css">
<!--
body {
	background-image: url(/images/bg.gif);
	margin-top: 0px;
}
.style2 {color: #FF0000}
-->
</style>
</head>

<body>
<form name="form1" method="post" action="default.asp">
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
          <th height="13"><div align="left">| 主题留言数:<span class="style2"><%=(jiaoliou_total)%></span></div></th>
          <td><div align="right">| <a href="/guest_register.asp">注册用户</a> | <a href="/login.asp">用户登入</a>  | <a href="lianxi.htm">联系我们</a> |   </div></td>
        </tr>
      </table>
      <hr width="800">
      <table width="800" height="30" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td nowrap><div align="left"><a href="/default1.asp">欢迎来到&quot;王国之心:记忆之链&quot;<strong>游戏论坛</strong>!</a></div></td>
          <td nowrap><div align="right">| <a href="/index2.htm">进入&quot;王国之心:记忆之链&quot;<strong>游戏主页</strong>!</a></div></td>
        </tr>
      </table>
      <hr width="800">
      <table width="800" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><div align="center">
              <table width="800" height="85" border="1" cellpadding="10" cellspacing="0" bordercolor="#999999">
                <tr>
                  <th width="70" height="20" bgcolor="#FFFFFF"><div align="center">状态</div></th>
                  <th width="400" height="20" bgcolor="#CCFFFF"><div align="center">发言主题</div></th>
                  <th width="100" height="20" bgcolor="#FFFFFF"><div align="center">作者</div></th>
                  <th width="200" height="20" bgcolor="#CCFFFF"><div align="center">发言时间</div></th>
                </tr>
                <% while ((Repeat1__numRows-- != 0) && (!gonggao.EOF)) { %>
                <tr>
                  <td width="70" height="30" bgcolor="#FFFFFF"><div align="center"><%=(gonggao.Fields.Item("state").Value)%></div></td>
                  <td width="400" height="30" nowrap bgcolor="#CCFFFF"><div align="left"><a href="/show.asp?ID=<%=(gonggao.Fields.Item("ID").Value)%>"><%=(gonggao.Fields.Item("title").Value)%></a></div></td>
                  <td width="100" height="30" bgcolor="#FFFFFF"><div align="center"><%=(gonggao.Fields.Item("user").Value)%></div></td>
                  <td width="200" height="30" bgcolor="#CCFFFF"><div align="center"><%=(gonggao.Fields.Item("date").Value)%> |  <%=(gonggao.Fields.Item("time").Value)%></div></td>
                </tr>
                <%
  Repeat1__index++;
  gonggao.MoveNext();
}
%>
</table>
          </div></td>
        </tr>
      </table>
      <hr width="800">
      <table width="800" border="1" cellpadding="10" cellspacing="0" bordercolor="#999999">
        <% while ((Repeat1__numRows-- != 0) && (!gonggao.EOF)) { %>
        <%
  Repeat1__index++;
  gonggao.MoveNext();
}
%>
        <% while ((Repeat2__numRows-- != 0) && (!jiaoliou.EOF)) { %>
        <tr>
          <td width="70" height="30" bgcolor="#FFFFFF">
          <div align="center"><%=(jiaoliou.Fields.Item("state").Value)%></div></td>
          <td width="400" height="30" nowrap bgcolor="#CCFFFF"><div align="left"><a href="/show.asp?ID=<%=(jiaoliou.Fields.Item("ID").Value)%>"><%=(jiaoliou.Fields.Item("title").Value)%></a></div></td>
          <td width="100" height="30" bgcolor="#FFFFFF"><div align="center"><%=(jiaoliou.Fields.Item("user").Value)%> </div></td>
          <td width="200" height="30" bgcolor="#CCFFFF"><div align="center"><%=(jiaoliou.Fields.Item("date").Value)%> | <%=(jiaoliou.Fields.Item("time").Value)%> </div></td>
        </tr>
        <%
  Repeat2__index++;
  jiaoliou.MoveNext();
}
%>
      </table>
      
      
      <% if (!jiaoliou.EOF || !jiaoliou.BOF) { %>
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
      <% } // end !jiaoliou.EOF || !jiaoliou.BOF %>

      <hr width="800">
    </div>
  </div>
</form>
</body>
</html>
<%
gonggao.Close();
%>
<%
jiaoliou.Close();
%>
