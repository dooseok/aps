<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<body topmargin=0 leftmargin=0>
<script language="javascript">

function login_frm_check(form)
{
	var strError = "";
	if(!form.M_ID.value)
	{
		strError += "* 아이디를 입력해주세요.\n";
	}
	if(!form.M_Password.value)
	{
		strError += "* 비밀번호를 입력해주세요.\n";
	}
	if(strError != "" )
	{
		alert(strError);
		return false;
	}
	form.submit();
}

function press_enter(strName)
{ 
	if(event.keyCode == 13) 
	{ 
		if (strName=="M_ID")
		{
			frm_login.M_Password.focus();
			return false;
		}
		else
			login_frm_check(frm_login);
	}
}
</script>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<center>
<table border=0 cellpadding=0 cellspacing=0 background="/img/bg_m_login.gif" width=452 height=232> 
<form name="frm_login" action="m_login_action.asp" method="post">
<tr>
	<td>
		<img src="/img/blank.gif" width=1px height=30px><br>
		<table width=47% align=center>
		<tr>
			<td align=right><font face="돋움" style="font-size:13px" color=#31659C>아<img src="/img/blank.gif" width=7px height=1px>이<img src="/img/blank.gif" width=7px height=1px>디</td>
			<td align=left><INPUT type="text" name="M_ID" style="width:125px" class=input onkeydown="javascript:press_enter('M_ID')"></td>
		</tr>
		<tr>
			<td align=right><font face="돋움" style="font-size:13px" color=#31659C>패스워드</td>
			<td align=left><INPUT type="password" name="M_Password" style="width:125px" class=input onkeydown="javascript:press_enter('M_Password')"></td>
		</tr>
		<tr>
			<td align=center colspan=2>
				<%=Make_BTN("로그인","javascript:login_frm_check(frm_login);","")%>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
</center>
<%
if Request("autologin")="yes" then
%>
<form name="frm_login2" action="m_login_action.asp" method="post">
<input type="hidden" name="M_ID" value="<%=request("autologinID")%>">
<input type="hidden" name="M_Password" value="<%=request("autologinPWD")%>">
</form>
<script language="javascript">
	frm_login2.submit();
</script>
<%
end if
%>
<!-- #include virtual = "/header/html_tail.asp" -->

