<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/luntan.asp" -->
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
<!-- TemplateBeginEditable name="doctitle" -->
<title>欢迎进入&quot;王国之心:记忆之链&quot;游戏论坛!</title>
<!-- TemplateEndEditable --><style type="text/css">
<!--
body {
	background-image: url(/images/bg.gif);
	margin-top: 0px;
}
.style1 {color: #FF0000}
-->
</style>
<!-- TemplateBeginEditable name="head" --><!-- TemplateEndEditable -->
</head>

<body>

<div align="center">
  <form name="form1" method="post" action="">
    <table border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="/images/luntan.jpg" width="320" height="142"></td>
      </tr>
    </table>
    <hr width="800">
    <table width="800" height="30" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <th><div align="left">| 主题留言数:<span class="style1"><%=(zhuti_total)%></span><!-- TemplateBeginEditable name="sum" --><!-- TemplateEndEditable --></div></th>
        <td><div align="right"><!-- TemplateBeginEditable name="lianjie" --><!-- TemplateEndEditable --> | <a href="lianxi.htm">联系我们</a> | </div></td>
      </tr>
    </table>
    <hr width="800">
    <table width="800" height="30" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td nowrap><div align="left"><a href="/default1.asp">欢迎来到&quot;王国之心:记忆之链&quot;<strong>游戏论坛</strong>!</a><!-- TemplateBeginEditable name="zhuti" --><!-- TemplateEndEditable --></div></td>
        <td nowrap><div align="right"> | <a href="/index2.htm">进入&quot;王国之心:记忆之链&quot;<strong>游戏主页</strong>!</a></div></td>
      </tr>
    </table>
    <hr width="800">
    <table width="800" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center"><!-- TemplateBeginEditable name="table" --><!-- TemplateEndEditable --></div></td>
      </tr>
    </table>
    <hr width="800">
    <!-- TemplateBeginEditable name="bottom" --><!-- TemplateEndEditable -->
  </form>
</div>
</body>
</html>
<%
zhuti.Close();
%>
