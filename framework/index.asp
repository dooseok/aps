<%@ Language=VBScript %>
<%
option explicit
Server.ScriptTimeOut = 99999
Response.Buffer = true
dim sys_DBcon
dim sys_DBconString

sys_DBConString = "Driver={SQL Server}; Server=localhost,1011; Database=MSEKOREA; Uid=sa; Pwd=mse7750?!;"

set sys_DBcon = server.CreateObject("adodb.connection")
sys_DBcon.ConnectionTimeout	= 300
sys_DBcon.CommandTimeout 	= 300

sys_DBcon.Open sys_DBConString
sys_DBcon.close
set sys_DBcon = nothing
%>

<html>
<head>

<meta http-equiv="Cache-Control" content="No-Cache">
<meta http-equiv="Pragma" content="No-Cache">
<meta http-equiv="Content-type" content="text/html;" charset="euc-kr">
<meta http-equiv="expires" content="now">

<title>MSERP</title>

<STYLE TYPE="text/css">
@import url(main.css);
</STYLE>

<script language="javascript" src="inc_share_function.js">
</script>
</head>

<body topmargin=0 leftmargin=0>

<div id="idLoading" width=100% height=100%>
<iframe src="inc_loading.asp" width="100%" height="100%" scrolling="no" border="0" frameborder="0"></iframe>
</div>
<table width=100% border=0 cellpadding=0 cellspacing=0>
<tr>
	<td width=15px><img src="/img/blank.gif" width=15px height=1px></td>
	<td align=center valign=top>
		<img src="/img/blank.gif" width=1px height=7px><br>	
		<div id="Progress" style="display:none;position:relative;">
		<table width=100% cellpadding=0 cellspacing=0 border=0>
		<tr>
			<td align=center bgcolor=white>
				<img src="/img/blank.gif" width=1 height=250>
				<img src="/img/ban_loading_ani.gif">
				<br><font face="돋움" style="font-size:10pt">데이터크기에 따라 로딩시간이 길어질 수 있습니다.<br>5분 이상 이 화면이 지속 되는 경우, 관리자에게 연락 바랍니다.<br>창을 닫지 마시고 잠시만 기다려주세요.</font>
			</td>
		</tr>
		</table>
		</div>
		<div id="Contents">
		
		</div>
	</td>
	<td width=15px><img src="/img/blank.gif" width=15px height=1px></td>
</tr>
</table>
</body>
</html>

<script language="javascript">
idLoading.style.display="none";
</script>


