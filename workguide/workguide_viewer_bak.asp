<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<%
dim s_PRD_Line
dim WG_SlideDelay
dim WG_Auto_YN
dim WG_Pos
dim WG_ResX
dim WG_ResY
dim strCurrentPartNo
dim bReload_YN

s_PRD_Line			= request("s_PRD_Line")
WG_SlideDelay		= request("WG_SlideDelay")
WG_Auto_YN			= request("WG_Auto_YN")
WG_Pos				= request("WG_Pos")
WG_ResX				= request("WG_ResX")
WG_ResY				= request("WG_ResY")
strCurrentPartNo	= request("strCurrentPartNo")
bReload_YN			= request("bReload_YN")

if strCurrentPartNo = "" then
	strCurrentPartNo = "select"
end if

dim SQL
dim RS1
dim CNT1
dim accCNT1
dim PRD_PartNo
dim oldPRD_PartNo
%>
<html>
<head>

<meta http-equiv="Cache-Control" content="No-Cache">
<meta http-equiv="Pragma" content="No-Cache">
<meta http-equiv="Content-type" content="text/html;" charset="euc-kr">
<meta http-equiv="Expires" content="0">
<meta Http-Equiv="Pragma-directive: no-cache">
<meta Http-Equiv="Cache-directive: no-cache">

<title>작업지도서</title>

<STYLE TYPE="text/css">
@import url(/header/main.css);
</STYLE>

<script language="javascript" src="/function/inc_share_function.js">
</script>
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<script type="text/javascript">
window.onload = maxWindow;

function maxWindow() {
    window.moveTo(0, 0);


    if (document.all) {
        top.window.resizeTo(screen.availWidth, screen.availHeight);
        top.window.resizeTo(screen.width, screen.height);
        //top.window.resizeTo(1920,1080);
    }

    else if (document.layers || document.getElementById) {
        if (top.window.outerHeight < screen.height || top.window.outerWidth < screen.width) {
            top.window.outerHeight = screen.height;
            top.window.outerWidth = screen.width;
        }
    }
}
</script>
<%
if Request("WG_Auto_YN") = "Y" then
%>
<table width=100% cellpadding=0 cellspacing=0 border=0>
<tr>
	<td align=center><img id=imgWorkGuide src="/img/blank.gif" style="border:solid 0px;"></td>
</tr>
</table>
<%
else
%>
<table width=100% cellpadding=0 cellspacing=0 border=0>
<tr>
	<td align=center><img id=imgWorkGuide src="/img/blank.gif" style="border:solid 0px;"></td>
</tr>
</table>
	<%
	if gM_ID = "shindk" then
	%>
<iframe frameborder=1 width="1000px" height="300px" name="ifrmViewer_ifrm" src="about:blank"></iframe>
<iframe frameborder=1 width="1000px" height="300px" name="ifrmViewer_ifrm_reload" src="workguide_viewer_ifrm_reloader.asp"></iframe>
	<%
	else
	%>
<iframe frameborder=0 width="0px" height="0px" name="ifrmViewer_ifrm" src="about:blank"></iframe>
<iframe frameborder=0 width="0px" height="0px" name="ifrmViewer_ifrm_reload" src="workguide_viewer_ifrm_reloader.asp"></iframe>
	<%
	end if
	%>
<div id="divManualModel" style="background:white;filter:alpha(opacity:'70');cursor:hand;top:950px;left:1670px;width:210px;height:40px;position:absolute;display:block;border:4px solid navy;" onclick="javascript:switchSelector();">
<table width=215px height=40px>
<tr>
	<td align=left>&nbsp;&nbsp;<span style="color:navy;font-size:16px"><b>PARTNO : <span id=spnCurrentPartNo><%=strCurrentPartNo%></span></b></span></td>
</tr>
</table>
</div>

<div id="divManualModelSelector" style="background:white;filter:alpha(opacity:'70');top:650px;left:1670px;width:210px;height:40px;position:absolute;display:none;border:4px solid navy;">
<table width=210px height=40px>
<tr height=38px>
	<td align=center>
		<span style="color:navy;font-size:16px"><b>Select PARTNO</b></span>
	</td>
</tr>
<%
	set RS1 = server.CreateObject("ADODB.RecordSet")
	SQL = "select PRD_PartNo from tbPWS_Raw_Data where PRD_Line = '"&s_PRD_Line&"' and (PRD_byHook_YN is null or PRD_byHook_YN = 'Y') order by PRD_Code desc"
	RS1.open SQL,sys_DBCon
	
	oldPRD_PartNo = ""
	CNT1 = 0
	accCNT1 = 0
	do until RS1.Eof
		PRD_PartNo = RS1("PRD_PartNo")
		if oldPRD_PartNo <> PRD_PartNo then
			
%>
<tr height=38px>
	<td align=center><span style="color:navy;font-size:16px;cursor:hand;" onclick="javascript:selectPartNo('<%=PRD_PartNo%>');"><b><%=PRD_PartNo%></b></span></td>
</tr>
<%
			oldPRD_PartNo = PRD_PartNo
			accCNT1 = accCNT1 + 1
		end if
		
		CNT1 = CNT1 + 1
		if CNT1 = 3000 or accCNT1 = 5 then
			exit do
		end if
		
		RS1.MoveNext
	loop
	RS1.Close
	set RS1 = nothing
%>
<tr height=38px>
	<td align=center>
		<span style="color:navy;font-size:16px;cursor:hand" onclick="javascript:reloadSelector();"><b>Reload</b></span>
		&nbsp;&nbsp;
		&nbsp;&nbsp;<span style="color:navy;font-size:16px;cursor:hand;" onclick="javascript:closeSelector();"><b>Close</b></span>
	</td>
</tr>
</table>
</div>

<script language="javascript">
var strCurrentPartNo = '<%=strCurrentPartNo%>';
<%
if bReload_YN = "Y" then
%>
var strSelectorStyleDisplay = 'block';
<%
else
%>
var strSelectorStyleDisplay = 'none';
<%
end if
%>
function switchSelector()
{
	if(strSelectorStyleDisplay=='block')
	{
		strSelectorStyleDisplay = 'none';
		closeSelector();
	}
	else
	{
		strSelectorStyleDisplay = 'block';
		openSelector();
	}
}

function reloadSelector()
{
	location.href='workguide_viewer.asp?bReload_YN=Y&WG_SlideDelay=<%=WG_SlideDelay%>&s_PRD_Line=<%=s_PRD_Line%>&WG_Auto_YN=<%=WG_Auto_YN%>&WG_Pos=<%=WG_Pos%>&WG_ResX=<%=WG_ResX%>&WG_ResY=<%=WG_ResY%>&strCurrentPartNo='+strCurrentPartNo;
}

function selectPartNo(strPartNo)
{
	strCurrentPartNo = strPartNo;
	spnCurrentPartNo.innerHTML = strPartNo;
	ifrmViewer_ifrm.location.href="workguide_viewer_ifrm.asp?strPartNo="+strPartNo+"&WG_Pos=<%=WG_Pos%>&WG_ResX=<%=WG_ResX%>&WG_ResY=<%=WG_ResY%>&WG_SlideDelay=<%=WG_SlideDelay%>";
	closeSelector();
}
function openSelector()
{
	divManualModelSelector.style.display = 'block';
}

function closeSelector()
{
	divManualModelSelector.style.display = 'none';
}

divManualModelSelector.style.display = strSelectorStyleDisplay;
</script>

<%
end if
%>
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->