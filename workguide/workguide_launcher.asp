<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_tb_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
 
<%
'���������, ���ظ���ʹ�?
dim s_PRD_Line
dim s_Base_WG_Pos
dim s_Checked

s_PRD_Line		= Request("s_PRD_Line")
s_Base_WG_Pos	= Request("s_Base_WG_Pos")
s_Checked		= Request("s_Checked")

'���� ���� ����
dim RS1
dim SQL
dim CNT1, CNT2
dim strWG_Pos
dim strWG_ResX
dim strWG_ResY
dim strWG_MCDelay
dim strWG_SlideDelay
dim strWG_SlideDelay_Main
dim strWG_Auto_YN

dim arrWG_Pos
dim arrWG_ResX
dim arrWG_ResY
dim arrWG_MCDelay
dim arrWG_SlideDelay
dim arrWG_SlideDelay_Main
dim arrWG_Auto_YN

dim PRD_PartNo

'����� ���� (��ġ, �ػ�, MC����, �����̵尣�ݵ�) �������� 
set RS1 = server.CreateObject("ADODB.RecordSet")
SQL = "select WG_Pos, WG_ResX, WG_ResY, WG_MCDelay, WG_SlideDelay, WG_SlideDelay_Main, WG_Auto_YN from tbWorkguide where PRD_Line='"&s_PRD_Line&"' order by WG_Pos asc"
RS1.Open SQL,sys_DBCon

strWG_Pos				= ""
strWG_ResX				= ""
strWG_ResY				= ""
strWG_MCDelay			= ""
strWG_SlideDelay		= ""
strWG_SlideDelay_Main	= ""
strWG_Auto_YN			= ""
do until RS1.Eof
	strWG_Pos				= strWG_Pos				& RS1("WG_Pos")				& ","
	strWG_ResX				= strWG_ResX			& RS1("WG_ResX")			& ","
	strWG_ResY				= strWG_ResY			& RS1("WG_ResY")			& ","
	strWG_MCDelay			= strWG_MCDelay			& RS1("WG_MCDelay")			& ","
	strWG_SlideDelay		= strWG_SlideDelay		& RS1("WG_SlideDelay")		& ","
	strWG_SlideDelay_Main	= strWG_SlideDelay_Main	& RS1("WG_SlideDelay_Main")	& ","
	strWG_Auto_YN			= strWG_Auto_YN			& RS1("WG_Auto_YN")			& ","
	RS1.MoveNext
loop
RS1.Close
arrWG_Pos				= split(strWG_Pos,",")
arrWG_ResX				= split(strWG_ResX,",")
arrWG_ResY				= split(strWG_ResY,",")
arrWG_MCDelay	 		= split(strWG_MCDelay,",")
arrWG_SlideDelay		= split(strWG_SlideDelay,",")
arrWG_SlideDelay_Main	= split(strWG_SlideDelay_Main,",")
arrWG_Auto_YN			= split(strWG_Auto_YN,",")
'[��Ȳ�� ���� ��------------------------------
'SQL = ""
'SQL = SQL & " select top 1 PRD_PartNo from "
'SQL = SQL & " tbPWS_Raw_Data "
'SQL = SQL & " where PRD_Line='"&s_PRD_Line&"' and "
'SQL = SQL & " 	(PRD_byHook_YN is null or PRD_byHook_YN = 'Y') "
'SQL = SQL & " order by PRD_Code desc "
'RS1.Open SQL,sys_DBCon
'if not(RS1.Eof or RS1.Bof) then
'	PRD_PartNo = RS1("PRD_PartNo")
'end if
'RS1.Close
'��Ȳ�� ���� ��------------------------------]
'[��Ȳ�� ���� ��------------------------------
'PRD_PartNo = application(s_PRD_Line&"_Last")
if PRD_PartNo = "" then
	SQL = ""
	SQL = SQL & "select top 1 SML_PartNo from tblStatus_Monitor_Line where "
	SQL = SQL & "SML_Line='"&s_PRD_Line&"' and "
	SQL = SQL & "SML_Type in ('N','F','T') and "
	SQL = SQL & "SML_Process = 'START' " 
	SQL = SQL & "order by SML_Code desc "
	RS1.Open SQL,sys_DBCon
	if not(RS1.Eof or RS1.Bof) then
		PRD_PartNo = RS1("SML_PartNo")
	end if
	RS1.Close
	'if application(s_PRD_Line&"_Last") = "" then
	'	application(s_PRD_Line&"_Last") = PRD_PartNo
	'end if
end if
'��Ȳ�� ���� ��------------------------------]




randomize
dim s_IDMonitorInstance
s_IDMonitorInstance = s_PRD_Line&int(999 * Rnd + 1)
application(s_PRD_Line) = "-"
%>

<script language="javascript">
var arrWorkGuide_VW = new Array();
var current_PartNo = '<%=PRD_PartNo%>';

var strWG_Pos				= '<%=strWG_Pos%>';
var arrWG_Pos				= strWG_Pos.split(',');
var strWG_ResX				= '<%=strWG_ResX%>';
var arrWG_ResX 				= strWG_ResX.split(',');
var strWG_ResY 				= '<%=strWG_ResY%>';
var arrWG_ResY 				= strWG_ResY.split(',');
var strWG_MCDelay 			= '<%=strWG_MCDelay%>';
var arrWG_MCDelay			= strWG_MCDelay.split(',');
var strWG_SlideDelay 		= '<%=strWG_SlideDelay%>';
var arrWG_SlideDelay 		= strWG_SlideDelay.split(',');
var strWG_SlideDelay_Main 	= '<%=strWG_SlideDelay_Main%>';
var arrWG_SlideDelay_Main 	= strWG_SlideDelay_Main.split(',');
var strWG_Auto_YN 			= '<%=strWG_Auto_YN%>';
var arrWG_Auto_YN 			= strWG_Auto_YN.split(',');

function set_current_PartNo(new_PartNo)
{
	current_PartNo = new_PartNo;
	update_ifrmLauncher();
}

function FileList2DB(new_PartNo)
{	
	if(confirm("�� 3���� �ð��� �ҿ�˴ϴ�.\n�Ϸ�޼����� ��ٷ��ֽʽÿ�.\n���κ��� ������ ������ �ʿ�� �����ϴ�."))
	{
		$('html,body').css('cursor','wait');
		ifrmLauncher_filelist2DB.location.href='workguide_launcher_ifrm_filelist2DB.asp';
	}
}

//���ظ���� ����
function changeBase_WG_Pos()
{
	location.href="workguide_launcher.asp?s_PRD_Line=<%=s_PRD_Line%>&s_Base_WG_Pos="+frmBase_WG_Pos.sltBase_WG_Pos.value;
}

//������ ����(����, ���ظ����, �ػ�, ��Ƽ�����)
function popup_workguide(s_PRD_Line,nWG_Pos,WG_ResX,WG_ResY,strMulti_YN)
{
	if (frmBase_WG_Pos.sltBase_WG_Pos.value-1 >= nWG_Pos)
	{
		alert("���� �����("+frmBase_WG_Pos.sltBase_WG_Pos.value+"��) ������ ����͸� �����Ͽ� �ֽʽÿ�");
		return false;
	}
	var sumLeft=0;
	for(var i=frmBase_WG_Pos.sltBase_WG_Pos.value-1; i< nWG_Pos-1; i++)
	{
		sumLeft = sumLeft + parseInt(arrWG_ResX[i]);
		//alert(parseInt(arrWG_ResX[i]));
	}
	sumLeft = String(sumLeft)+"px";
	
	//������ ���������� �ݱ�
	if(typeof(arrWorkGuide_VW[nWG_Pos-1])=='object')
		arrWorkGuide_VW[nWG_Pos-1].close();
		
	//�˾� ����
	arrWorkGuide_VW[nWG_Pos-1] = window.open('workguide_viewer.asp?WG_SlideDelay='+arrWG_SlideDelay[nWG_Pos-1]+'&WG_SlideDelay_Main='+arrWG_SlideDelay_Main[nWG_Pos-1]+'&WG_ResX='+arrWG_ResX[nWG_Pos-1]+'&WG_ResY='+arrWG_ResY[nWG_Pos-1]+'&s_PRD_Line=<%=s_PRD_Line%>&WG_Pos='+nWG_Pos+'&WG_Auto_YN='+arrWG_Auto_YN[nWG_Pos-1],'_blank','width=100px,height=100px,top=0px,left='+sumLeft+',resizable=no,scrollbars=no,status=no,location=no,menubar=no,toolbar=no');
	
	//�ܵ����� ���� �Ŷ��
	if(strMulti_YN != 'Y')
	{
		for(var i=0; i<frmWG_Res.chkPos.length; i++)
		{
			if(i==nWG_Pos-1)
				frmWG_Res.chkPos[i].checked = true;
			else
				frmWG_Res.chkPos[i].checked = false;	
		}
		update_ifrmLauncher();
	}
}

//���İ���
function update_ifrmLauncher()
{
	<%if gM_ID = "shindk" or gM_ID = "shindh" or gM_ID = "woojm" then%>
	ifrmLauncher.location.href='workguide_launcher_ifrm2.asp?s_PRD_Line=<%=s_PRD_Line%>&strIDX=<%=server.urlencode("0,0,0,0,0,0,0,0,0,0,0,0,0,0,0")%>&strSlideCNT=<%=server.urlencode("0,0,0,0,0,0,0,0,0,0,0,0,0,0,0")%>&strPrePartNo=<%=server.urlencode("0,0,0,0,0,0,0,0,0,0,0,0,0,0,0")%>';
	<%else%>
	ifrmLauncher.location.href='workguide_launcher_ifrm.asp?s_PRD_Line=<%=s_PRD_Line%>&strIDX=<%=server.urlencode("0,0,0,0,0,0,0,0,0,0,0,0,0,0,0")%>&strSlideCNT=<%=server.urlencode("0,0,0,0,0,0,0,0,0,0,0,0,0,0,0")%>&strPrePartNo=<%=server.urlencode("0,0,0,0,0,0,0,0,0,0,0,0,0,0,0")%>';
	<%end if%>
	
	ifrmLauncher_mc_checker.location.href='workguide_launcher_ifrm_MC_checker.asp?s_PRD_Line=<%=s_PRD_Line%>&s_IDMonitorInstance=<%=s_IDMonitorInstance%>';
	ifrmLauncher_reloader.location.href='workguide_launcher_ifrm_reloader.asp';
}

function popup_workguide_selected(s_PRD_Line)
{
	var cnt=0;
	for(var i=0; i<frmWG_Res.chkPos.length; i++)
	{
		if(frmWG_Res.chkPos[i].checked == true)
		{
			popup_workguide(s_PRD_Line,i+1,arrWG_ResX[i],arrWG_ResY[i],'Y');
		}
	}
	
	update_ifrmLauncher();
}

function frmWG_Res_Check()
{
	for(var i = 0; i < frmWG_Res.WG_Res.length; i++)
	{
		var WG_Res = frmWG_Res.WG_Res[i].value.toLowerCase();
		if (WG_Res.indexOf('x') < 0)
		{
			alert("["+(i+1)+"]��° �ػ󵵰� ����x������ ���°� �ƴմϴ�.");
			return false;
		}
		
		var WG_ResX = WG_Res.substring(0,WG_Res.indexOf('x'));
		var WG_ResY = WG_Res.substring(WG_Res.indexOf('x')+1,WG_Res.length);

		if (IsNum(WG_ResX) && IsNum(WG_ResY)){}
		else
		{
			alert("["+(i+1)+"]��° �ػ󵵰� ����x������ ���°� �ƴմϴ�!");
			return false;
		}
		
		if (IsNum(frmWG_Res.WG_MCDelay[i].value)) {}
		else
		{
			alert("["+(i+1)+"]��° M/C ������(��)�� ���ڰ� �ƴմϴ�!");
			return false;
		}
		
		if (Number(frmWG_Res.WG_MCDelay[i].value) > 1800)
		{
			alert("["+(i+1)+"]��° M/C ������(��)�� 30��(1800��) �̳��̾�� �մϴ�.!");
			return false;
		}
	}
	
	frmWG_Res.submit();
}

function set_custom_model()
{
	if (!frmBase_WG_Pos.strCustom_Model.value)
	{
		alert("���� �����Ͽ� �ֽʽÿ�.");
		return false;
	}
	else if (frmBase_WG_Pos.strCustom_Model.value.length != 11)
	{
		alert("��Ʈ�ѹ��� 11�ڸ��� ���˴ϴ�.");
		return false;
	}
		
	if(confirm('['+frmBase_WG_Pos.strCustom_Model.value+']�� ���� ������ ������ �Է��մϴ�.'))
	{
		var Custom_Model;
		Custom_Model = frmBase_WG_Pos.strCustom_Model.value;
		frmBase_WG_Pos.strCustom_Model.value = "";
		ifrmCustomModel.location.href='workguide_set_custom_model.asp?strCustom_Model='+Custom_Model+'&s_PRD_Line=<%=s_PRD_Line%>';
	}
}
</script>
<%
call BOMSub_Guide()
%>
<center>
<div class="page-header">
<h2><%=s_PRD_Line%> �۾������� ����</h2>
</div>
<table width=600px border=0>
<form name="frmBase_WG_Pos" method="post" action="#">
<tr width=600px>
	<td align=center>�ָ���� 
		<select name="sltBase_WG_Pos" onchange="javascript:changeBase_WG_Pos()">
<%
for CNT1 = 1 to 15
%>
			<option value="<%=CNT1%>"<%if int(s_Base_WG_Pos)=CNT1 then%> selected<%end if%>><%=CNT1%>��</option>
<%
next
%>
		</select>
	</td>
	<td align=center>��������Է�
		<input type="text" size=13 name="strCustom_Model" onclick="show_BOMSub_Guide(this,'frmBase_WG_Pos',0);">
		<input type="button" value="���" class="button" onclick="javascript:set_custom_model();">
	</td>
	<td align=center>
		&nbsp;
		<!--<input type="button" value="�̹��� DB������Ʈ" class="button" onclick="javascript:FileList2DB();">-->
	</td>
</tr>
</form>
</table>

<br>

<table cellpadding=0 cellspacing=0 border=0 bgcolor=black>
<form name="frmWG_Res" method="post" action="workguide_launcher_action.asp">
<input type="hidden" name="s_PRD_Line" value="<%=s_PRD_Line%>">
<tr bgcolor=white>
	<td bgcolor=white>
		<table width=140px cellpadding=0 cellspacing=0 border=0 style="border-left:solid 1px #cccccc;border-right:solid 1px #cccccc">
		<tr height=25px>
			<td align=center>����</td>
		</tr>
		<tr height=25px>
			<td align=center>�ػ�</td>
		</tr>
		<tr height=25px>
			<td align=center>M/C����(��)</td>
		</tr>
		<tr height=25px>
			<td align=center>�� �����̵尣��(��)</td>
		</tr>
		<tr height=25px>
			<td align=center>��Ÿ �����̵尣��(��)</td>
		</tr>
		<tr height=25px>
			<td align=center>�𵨼���</td>
		</tr>
		<tr height=55px>
			<td align=center>�۾�</td>
		</tr>
		</table>
	</td>
<%for CNT1 = 0 to ubound(arrWG_Pos)-1%>
	<td bgcolor=white>
		<table width=73px cellpadding=0 cellspacing=0 border=0 style="border-right:solid 1px #cccccc">
		<tr height=25px>
			<td align=center><input name="chkPos" type=checkbox><%=arrWG_Pos(CNT1)%>��</td>
		</tr>
		<tr height=25px>
			<td align=center><input style="width:60px" type="text" name="WG_Res" value="<%=arrWG_ResX(CNT1)%>x<%=arrWG_ResY(CNT1)%>"></td>
		</tr>
		<tr height=25px>
			<td align=center><input style="width:60px" type="text" name="WG_MCDelay" value="<%=arrWG_MCDelay(CNT1)%>"></td>
		</tr>
		<tr height=25px>
			<td align=center><select name="WG_SlideDelay_Main" style="width:60px"><%for CNT2 = 10 to 600 step 10%><option value="<%=CNT2%>"<%if int(arrWG_SlideDelay_Main(CNT1))=CNT2 then%> selected<%end if%>><%=CNT2%></option><%next%></select></td>
		</tr>
		<tr height=25px>
			<td align=center><select name="WG_SlideDelay" style="width:60px"><%for CNT2 = 10 to 120 step 10%><option value="<%=CNT2%>"<%if int(arrWG_SlideDelay(CNT1))=CNT2 then%> selected<%end if%>><%=CNT2%></option><%next%></select></td>
		</tr>
		<tr height=25px>
			<td align=center><select name="WG_Auto_YN" style="width:60px">
				<option value="Y"<%if arrWG_Auto_YN(CNT1)="Y" then%> selected<%end if%>>�ڵ�</option>
				<option value="N"<%if arrWG_Auto_YN(CNT1)="N" then%> selected<%end if%>>����</option>
				</select></td>
		</tr>
		<tr height=55px>
			<td align=center><input type="button" value="����" class="btn btn-sm btn-primary" onclick="javascript:popup_workguide('<%=s_PRD_Line%>',<%=arrWG_Pos(CNT1)%>,<%=arrWG_ResX(CNT1)%>,<%=arrWG_ResY(CNT1)%>)"></td>
		</tr>
		</table>
	</td>
<%next%>
</tr>
</form>
</table>
</center>
<Br>
<input type="button" value="��������" class="btn btn-sm btn-info" onclick="javascript:frmWG_Res_Check();">&nbsp;
<input type="button" value="���ÿ���" class="btn btn-sm btn-primary" onclick="javascript:popup_workguide_selected('<%=s_PRD_Line%>');">
<br>
<center>
	<br>
	<h6><span style="background-color:pink">&nbsp;����. [�ػ�] �� [M/C ������]�� �����Ͻ� �Ŀ��� [��������]��ư�� Ŭ���ϼž� �մϴ�.&nbsp;</span></h6>
	<h6>tip. �ָ���Ϳ�, �����̵� ������ �����ϴ� ��� >
	������ ������ �� ��, �������� ������������ [����������]�� �����ϸ� �˴ϴ�.</h6> 
</center>
<%
'response.write Request.Cookies("Member")("M_ID") &"//"
if gM_ID = "shindk" then
%>
ifrmCustomModel<Br>
<iframe frameborder=1 width=1000px height=50px name="ifrmCustomModel" src="about:blank"></iframe><Br>
ifrmLauncher_filelist2DB<Br>
<iframe frameborder=1 width=1000px height=50px name="ifrmLauncher_filelist2DB" src="about:blank"></iframe><Br>
ifrmLauncher_reloader<Br>
<iframe frameborder=1 width=1000px height=50px name="ifrmLauncher_reloader" src="about:blank"></iframe><Br>
ifrmLauncher_mc_checker<Br>
<iframe frameborder=1 width=1000px height=50px name="ifrmLauncher_mc_checker" src="about:blank"></iframe><Br>
ifrmLauncher<Br>
<iframe frameborder=1 width=1000px height=300px name="ifrmLauncher" src="about:blank"></iframe>
<%
else
%>
<iframe frameborder=0 width=0px height=0px name="ifrmCustomModel" src="about:blank"></iframe>
<iframe frameborder=0 width=0px height=0px name="ifrmLauncher_filelist2DB" src="about:blank"></iframe>
<iframe frameborder=0 width=0px height=0px name="ifrmLauncher_reloader" src="about:blank"></iframe>
<iframe frameborder=0 width=0px height=0px name="ifrmLauncher_mc_checker" src="about:blank"></iframe>
<iframe frameborder=0 width=0px height=0px name="ifrmLauncher" src="about:blank"></iframe>
<%
end if
%>


<script language="javascript">
function AutoCheck(strChecked)
{
	var arrChecked = strChecked.split(',');
	
	for(var i=0; i<frmWG_Res.chkPos.length; i++)
	{
		if(i+1 >= parseInt(arrChecked[0]) && i+1 <= parseInt(arrChecked[1]))
			frmWG_Res.chkPos[i].checked = true;
		else
			frmWG_Res.chkPos[i].checked = false;	
	}
}

var s_Checked = '<%=s_Checked%>';
if (!s_Checked) 
{
	if(parseInt(frmBase_WG_Pos.sltBase_WG_Pos.value) == 1)
		AutoCheck('1,7');
	else if(parseInt(frmBase_WG_Pos.sltBase_WG_Pos.value) == 8)
		AutoCheck('8,15');
}
else
{
	AutoCheck(s_Checked);
}	

</script>

<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->