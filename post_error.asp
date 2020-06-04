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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/luntan.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- InstanceBeginEditable name="doctitle" -->
<title>操作失败信息</title>
<!-- InstanceEndEditable --><style type="text/css">
<!--
body {
	background-image: url(/images/bg.gif);
	margin-top: 0px;
}
.style1 {color: #FF0000}
-->
</style>
<!-- InstanceBeginEditable name="head" -->
<style type="text/css">
<!--
.style1 {color: #FF0000;
	font-weight: bold;
}
-->
</style>
<meta http-equiv="refresh" content="3;URL=/login.asp">
<!-- InstanceEndEditable -->
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
        <th><div align="left">| 主题留言数:<span class="style1"><%=(zhuti_total)%></span><!-- InstanceBeginEditable name="sum" --><!-- InstanceEndEditable --></div></th>
        <td><div align="right"><!-- InstanceBeginEditable name="lianjie" -->| <a href="/guest_register.asp">注册用户</a> | <a href="/login.asp">用户登入</a> |<!-- InstanceEndEditable --> | <a href="lianxi.htm">联系我们</a> | </div></td>
      </tr>
    </table>
    <hr width="800">
    <table width="800" height="30" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><div align="left"><a href="/default1.asp">欢迎来到&quot;王国之心:记忆之链&quot;<strong>游戏论坛</strong>!</a><!-- InstanceBeginEditable name="zhuti" --><!-- InstanceEndEditable --></div></td>
        <td><div align="right"><a href="/index2.htm">进入&quot;王国之心:记忆之链&quot;<strong>游戏主页</strong>!</a></div></td>
      </tr>
    </table>
    <hr width="800">
    <table width="800" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center"><!-- InstanceBeginEditable name="table" -->
          <table width="800" border="1" cellspacing="0" cellpadding="10">
            <tr>
              <th bgcolor="#CCFFCC"><div align="center">操作失败信息</div></th>
            </tr>
            <tr>
              <td bgcolor="#FFFFFF"><div align="center">
                  <p><strong>对不起</strong>,您需要<strong>登入用户</strong>才能发布消息!</p>
                  <p> <span class="style1">3</span> 秒钟后将自动转到 <strong><a href="/login.asp">用户登入</a></strong> 页面!</p>
                  <p><a href="/default1.asp">返回论坛首页</a><br>
                  </p>
              </div></td>
            </tr>
          </table>
        <!-- InstanceEndEditable --></div></td>
      </tr>
    </table>
    <hr width="800">
    <!-- InstanceBeginEditable name="bottom" --><!-- InstanceEndEditable -->
  </form>
</div>
</body>
<!-- InstanceEnd --></html>
<%
zhuti.Close();
%>
