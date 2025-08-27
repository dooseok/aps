<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/html_jq_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->




<% 
dim CNT1
dim CNT2
dim CNT3
dim CNT4

dim SQL
dim RS1
dim RS2
dim B_Code

dim B_Opt_YN

dim nWidth_BOMSub_PNO
dim nWidth_WorkOrder
dim nWidth_Parts_PNO
dim nWidth_Parts_PNO2
dim nWidth_Checkbox
dim nWidth_Type
dim nWidth_CheckSum
dim nWidth_SType
dim nWidth_Desc
dim nWidth_Spec
dim nWidth_Maker
dim nWidth_Maker2
dim nWidth_Remark
dim nWidth_RemarkF
dim nWidth_Sum

dim nWidth_Diff_Assy_PNO
dim nWidth_Diff_Qty

nWidth_BOMSub_PNO = 25
nWidth_WorkOrder = 50
nWidth_Parts_PNO = 100
nWidth_Parts_PNO2 = 100
nWidth_Checkbox = 20
nWidth_Type = 90
nWidth_CheckSum = 50
nWidth_SType = 100
nWidth_Desc = 200
nWidth_Spec = 600
nWidth_Maker = 150
nWidth_Maker2 = 150
nWidth_Remark = 150
nWidth_RemarkF = 100

nWidth_Diff_Assy_PNO = 100
nWidth_Diff_Qty = 50

B_Code = Request("B_Code")
set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")
call Material_Guide()

dim strPD_Desc

dim ComplexKey
dim dicComplexAcc
set dicComplexACC =Server.CreateObject("Scripting.Dictionary")
dim arrtempQty

dim strRemark
dim arrRemark

dim strMaterial_M_PNO
dim strBM_WType
dim strBM_Maker

'Diff===
dim Diff_YN
dim Diff_Disable_YN

dim strTable_B_Code
dim strTable_Diff_B_Code

Diff_YN = Request("Diff_YN")

if B_Code = "" then
	Diff_YN = "N"
	Diff_Disable_YN = "Y"
else
	SQL = "select top 1 B_Code from tbBOM where B_D_No = '"&Request("DNO")&"' and B_Code < "&B_Code&" order by B_Code desc"
	RS1.open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		Diff_YN = "N"
		Diff_Disable_YN = "Y"
	else
		Diff_B_Code = RS1(0)
	end if
	RS1.Close
end if

strPD_Desc = "-"
SQL = "select BMDD_Desc_BOM from tblBOM_Mask_Desc_Detail"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
	strPD_Desc = strPD_Desc & RS1("BMDD_Desc_BOM") & "-"
	RS1.MoveNext
loop
RS1.Close
strPD_Desc = lcase(strPD_Desc)

if Diff_YN = "Y" then
	
	dim strComplexKey
	dim strComplexMaker
	dim strComplexQty
	dim strComplexAcc
	dim arrComplexKey
	dim arrComplexMaker
	dim arrComplexQty
	dim arrComplexAcc

	dim dicParts
	dim oldQty
	dim bChangeQty
	dim bChangeMaker

	dim currentPNO
	dim currentNO
	
	dim Diff_B_Code
	
	SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&B_Code	
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
	else
		if RS1("B_Version_Current_YN") = "Y" then
			strTable_B_Code = "tbBOM_Qty"
		else
			strTable_B_Code = "tbBOM_Qty_Archive"
		end if
	end if
	RS1.Close
	
	SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&Diff_B_Code	
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
	else
		if RS1("B_Version_Current_YN") = "Y" then
			strTable_Diff_B_Code = "tbBOM_Qty"
		else
			strTable_Diff_B_Code = "tbBOM_Qty_Archive"
		end if
	end if
	RS1.Close
	
	SQL = 		"select "&vbcrlf
	SQL = SQL & "	BOM_Sub_BS_D_No, "&vbcrlf
	SQL = SQL & "	Parts_P_P_No, "&vbcrlf
	SQL = SQL & "	BQ_Remark, "&vbcrlf
	SQL = SQL & "	BQ_P_Maker, "&vbcrlf
	SQL = SQL & "	BQ_Order, "&vbcrlf
	SQL = SQL & "	BQ_Qty "&vbcrlf
	SQL = SQL & "from "&strTable_Diff_B_Code&" "&vbcrlf
	SQL = SQL & "where "&vbcrlf
	SQL = SQL & "	BOM_B_Code="&Diff_B_Code&" order by BOM_Sub_BS_D_No, BQ_Code"&vbcrlf
	
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		if ucase(RS1("BQ_Order")) = "R" then
			strComplexKey	= strComplexKey & RS1("BOM_Sub_BS_D_No")&"_"&RS1("Parts_P_P_No")&"_"&RS1("BQ_Remark")& "_R/|"
		else
			strComplexKey	= strComplexKey & RS1("BOM_Sub_BS_D_No")&"_"&RS1("Parts_P_P_No")&"_"&RS1("BQ_Remark")& "_X/|"
		end if
		strComplexMaker	= strComplexMaker & RS1("BQ_P_Maker")& "/|"
		strComplexQty	= strComplexQty & RS1("BQ_Qty")& "/|"
		RS1.MoveNext
	loop
	RS1.Close
	arrComplexKey	= split(strComplexKey,"/|")
	arrComplexMaker	= split(strComplexMaker,"/|")
	arrComplexQty	= split(strComplexQty,"/|")
	
	for CNT1 = 0 to ubound(arrComplexKey) - 1
		CNT3 = 0
		for CNT2 = 0 to CNT1
			if cstr(arrComplexKey(CNT1)) = cstr(arrComplexKey(CNT2)) then
				CNT3 = CNT3 + 1
			end if
		next
		strComplexACC	= strComplexACC & CNT3& "/|"
	next
	arrComplexACC = split(strComplexACC,"/|")
	
	
	set dicParts =Server.CreateObject("Scripting.Dictionary")
	
	SQL = 		"select "&vbcrlf
	SQL = SQL & "	Parts_P_P_No, "&vbcrlf
	SQL = SQL & "	BQ_Order, "&vbcrlf
	SQL = SQL & "	BQ_Remark "&vbcrlf
	SQL = SQL & "from "&strTable_Diff_B_Code&" "&vbcrlf
	SQL = SQL & "where "&vbcrlf
	SQL = SQL & "	BOM_B_Code="&Diff_B_Code
	RS1.Open SQL,sys_DBCon
	do until RS1.Eof
		if ucase(RS1("BQ_Order")) = "R" then
			if not(dicParts.Exists(RS1("Parts_P_P_No")&"_"&RS1("BQ_Remark")&"_R")) then
				dicParts.Add RS1("Parts_P_P_No")&"_"&RS1("BQ_Remark")&"_R",0
			end if
		else
			if not(dicParts.Exists(RS1("Parts_P_P_No")&"_"&RS1("BQ_Remark")&"_X")) then
				dicParts.Add RS1("Parts_P_P_No")&"_"&RS1("BQ_Remark")&"_X",0
			end if
		end if
		RS1.MoveNext
	loop
	RS1.Close
end if
'Diff===


dim oldModel_CNT
dim oldParts_CNT

dim Model_CNT
dim Parts_CNT

Model_CNT		= Request("Model_CNT")
if trim(Model_CNT) = "" then
	Model_CNT	= 10
end if
oldModel_CNT	= Request("oldModel_CNT")
if trim(oldModel_CNT) = "" then
	oldModel_CNT	= 10
end if


Parts_CNT		= Request("Parts_CNT")
if trim(Parts_CNT) = "" then
	Parts_CNT	= 30
end if
oldParts_CNT	= Request("oldParts_CNT")
if trim(oldParts_CNT) = "" then
	oldParts_CNT	= 30
end if


dim ROW
dim COL

dim div_top
dim div_left

div_top		= 60
div_left	= 10

dim arrMODEL
dim MODEL
dim arrDNOSUB
dim arrDNOCONFIRM
dim DNOSUB
dim DNOCONFIRM
dim arrQTY
dim QTY
dim strQty

dim strNewParts

call BOMPNO_Guide(B_Code)
%>

<script language="javascript">

var oldROW_Over = -1;
var oldCOL_Over = -1;

function toggleTopColor(CNT1)
{
	//var objTD;
	//objTD = eval("idCOL_TopOVER_"+CNT1);
	
	if(frmB_Model_Reg.DNOCONFIRM[CNT1-1].checked)
	{
		//objTD.style.backgroundColor="#4ce142";
		document.getElementById("idCOL_TopOVER_"+CNT1).style.backgroundColor="#4ce142";
	}
	else
	{
		//objTD.style.backgroundColor="transparent";
		document.getElementById("idCOL_TopOVER_"+CNT1).style.backgroundColor="transparent";
	}
}
function setMouseOver(ROW, COL)
{
	var objTR;
	var objTD;
	
	//objTR = eval("idROW_OVER_"+ROW);
	//objTR.style.backgroundColor="#86FBFF";
	
	document.getElementById("idROW_OVER_"+ROW).style.backgroundColor="#86FBFF";
	
	//objTD = eval("idCOL_OVER_"+COL);
	//objTD.style.backgroundColor="#86FBFF";
	
	document.getElementById("idCOL_OVER_"+COL).style.backgroundColor="#86FBFF";
	
	if (oldROW_Over > 0 && oldROW_Over != ROW)
	{
		//objTR = eval("idROW_OVER_"+oldROW_Over);
		//objTR.style.backgroundColor="transparent";		
		document.getElementById("idROW_OVER_"+oldROW_Over).style.backgroundColor="transparent";
	}
	
	
	if (oldCOL_Over > 0 && oldCOL_Over != COL)
	{
		//objTD = eval("idCOL_OVER_"+oldCOL_Over);
		//objTD.style.backgroundColor="transparent";
		document.getElementById("idCOL_OVER_"+oldCOL_Over).style.backgroundColor="transparent";
	}
	
	oldROW_Over = ROW;
	oldCOL_Over = COL;
}

function Form_Submit()
{
<%
if gM_ID = "shindk" then
%>
		frmB_Model_Reg.action="b_model_reg_action.asp";
		Show_Progress();
		frmB_Model_Reg.submit();
<%
else
%>
	if(confirm('약 1분 정도의 시간이 소요됩니다.\n창을 닫지 마시고, 잠시 기다려주십시오.'))
	{
		//alert("서버 점검중입니다.");
		frmB_Model_Reg.action="b_model_reg_action.asp";
		Show_Progress();
		frmB_Model_Reg.submit();
	}
<%
end if
%>
}

function XLS_Form_Submit()
{
	if (!frmXLS_Upload.BOM_XLS.value)
	{
		alert("파일을 선택해주세요.")
		return false;
	}
	else
	{
		
		/*
		if(frmXLS_Upload.oldXLS.checked)
			frmXLS_Upload.action="xls_upload_action_old.asp";
		
		confirm('잠시만 기다려주세요.\n예상 소요시간 1분')
		{
			frmXLS_Upload.submit();
		}
		*/
	}
	
	confirm('잠시만 기다려주세요.\n예상 소요시간 1분')
	{
		frmXLS_Upload.submit();
	}
}

function XLS_Download()
{
	confirm('다운로드 메세지가 뜰 때까지, 잠시만 기다려주세요.\n예상 소요시간 1분')
	{
		document.getElementById('ifrmXLSDown').src = "/bom/xls_download_action.asp?b_d_no=<%=Request("DNO")%>&b_code=<%=B_Code%>";
	}
}

var strOldFindObj = "";

function pageFind()
{	
	if (event.keyCode == 13)
	{
		var strNowFindObj = frmXLS_Upload.strFindString.value.toUpperCase();
		$("#"+strOldFindObj).parent().parent().css( "background-color", "transparent" );
		$("#"+strNowFindObj).parent().parent().css( "background-color", "#00FF00" );
		document.getElementById(strNowFindObj).scrollIntoView();
		strOldFindObj = strNowFindObj;
		location.href="#"+frmXLS_Upload.strFindString.value;
	}
}

var toCheck = false;
function Check_Confirm()
{
	if(toCheck)
		toCheck = false;
	else
		toCheck = true;
		
	if(frmB_Model_Reg.DNOCONFIRM.length)
	{
		for(var i = 0; i < frmB_Model_Reg.DNOCONFIRM.length; i++)
			frmB_Model_Reg.DNOCONFIRM[i].checked = toCheck;
	}
	else
		frmB_Model_Reg.DNOCONFIRM.checked = toCheck;
}

function moveVersion(B_Code)
{
	if(confirm('해당 시방번호로 이동합니다.\n약 1분 정도의 시간이 소요됩니다.\n잠시 기다려주십시오.'))
	{
		Show_Progress();
		location.href="db_load_action.asp?B_Code="+B_Code+"&Diff_YN=<%=Diff_YN%>";
	}
}

function chkDiff(obj)
{
	if(obj.checked)
	{
		<%if Diff_YN<>"Y" then%>
		if(confirm('Diff뷰어로 전환합니다.\n잠시 기다려주십시오.'))
		{ 
			Show_Progress();
			location.href="db_load_action.asp?B_Code=<%=B_Code%>&Diff_YN=Y";
		}
		<%end if%>
	} 
	else 
	{
		
		<%if Diff_YN="Y" then%>
		if(confirm('일반 뷰어로 전환합니다.\n잠시 기다려주십시오.'))
		{
			Show_Progress();
			location.href="db_load_action.asp?B_Code=<%=B_Code%>&Diff_YN=N";
		}
		<%end if%>
	}
}

function BOMPrint()
{
	<%if Model_CNT > 30 then%>
	alert("옵션이 30개를 초과하는 경우에는 인쇄를 지원하지 않습니다.\n도면을 엑셀로 다운로드하여 출력해 주십시오.");
	<%else%>
	if(confirm('도면을 인쇄합니다.\n인쇄 창이 뜰 때까지 잠시만 기다려주세요.'))
	{	
		window.open("/bom/db_load_action.asp?b_code=<%=B_Code%>&mode=print","BOMPrint","height=100px,width=100px,top=2000px,lef=2000px,status=yes,toolbar=yes,location=yes,directories=yes,location=yes,menubar=yes,resizable=yes,scrollbars=yes,titlebar=yes")
	}
	<%end if%>
}
</script>
</script>

<%
ROW = 1
COL = 1
%>
<table border=0 cellspacing=1 cellpadding=0 width=1164px style=table-layout:fixed bgcolor="#999999" align=left>
<form name="frmXLS_Upload" action="xls_upload_action.asp" method="post" enctype="MULTIPART/FORM-DATA">
<input name="B_Code" type="hidden" value="<%=B_Code%>">
<input name="Diff_YN" type="hidden" value="<%=Diff_YN%>">
<tr>
	<td bgcolor=white>
		<table border=0 cellspacing=2 cellpadding=0 width=1140px style=table-layout:fixed bgcolor="#ffffff">
		<tr>
			<td width=150px align=left>PartNo:&nbsp;<input name="DNO" id="idDNO" type="text" style="width:100px; border:1px solid #999999;" readonly value="<%=Request("DNO")%>"></td>
<%
'if instr(admin_b_bom_update,"-"&Request.Cookies("ADMIN")("M_Authority")&"-") > 0 then
%>
			<td width=55px align=left><%=Make_S_BTN("저장","javascript:Form_Submit();","")%></td>
			<!--<td width=77px><%=Make_BTN("전체확인","javascript:Check_Confirm();","")%></td>-->

			<td align=cente width=130px>
				Version:&nbsp;<select style="width:70px" onchange="javascript:moveVersion(this.value);">
<%
SQL = "select B_Code,B_Version_Code, B_Opt_YN from tbBOM where B_D_No = '"&Request("DNO")&"' order by B_Code desc"
RS1.Open SQL,sys_DBCon
B_Opt_YN = "Y"
do until RS1.Eof
	if int(RS1("B_Code")) = int(B_Code) then
		if RS1("B_Opt_YN") <> "Y" then
			B_Opt_YN = "N"
		end if
	end if
%>	
<option value="<%=RS1("B_Code")%>"<%if int(RS1("B_Code"))=int(B_Code) then%> selected<%end if%>><%=RS1("B_Version_Code")%></option>		
<%	
	RS1.MoveNext
loop
RS1.Close
%>				
				</select>	
			</td>
			<td>검색<input type="text" size=15 name="strFindString" id="strFindString" onkeydown="pageFind();" onDblclick="javascript:show_BOMPNO_Guide(this,<%=B_Code%>);" ></td>
			<td width=58px><!--구양식<input type="checkbox" name="oldXLS">--></td>
			<td width=200px align=right><input type="file" name="BOM_XLS" style="width:95%" class="input" id="idBOM_XLS"></td>
			<td width=77px><%=Make_BTN("XLS UP","javascript:XLS_Form_Submit();","")%></td>
			<td width=77px><%=Make_BTN("XLS DOWN","javascript:XLS_Download();","")%></td>
			<td width=77px><%=Make_BTN("PRINT","javascript:BOMPrint();","")%></td>
			<td width=30px><%if gM_ID="shindk" then%><%=B_Opt_YN%><%end if%>&nbsp;</td>
			<td width=65px>Diff모드<input type="checkbox" name="chkDiff_YN" onclick="javascript:<%if Diff_Disable_YN = "Y" then%>alert('비교할 이전시방이 없습니다.');<%else%>chkDiff(this);<%end if%>" value="<%=Diff_YN%>" <%if Diff_YN = "Y" then%>checked<%end if%>></td>
			<td width=10px></td>
		</tr>
		</table> 
	</td>
</tr>
</form>
</table>	
<%if gM_ID="shindk" then%>
<iframe id="ifrmXLSDown" src="about:blank" frameborder=1 width=1000px height=500px></iframe>
<%else%>

<iframe id="ifrmXLSDown" src="about:blank" frameborder=0 width=0px height=0px></iframe>
<%end if%>
<%
nWidth_Sum = nWidth_BOMSub_PNO*Model_CNT+nWidth_WorkOrder+nWidth_Parts_PNO+nWidth_Parts_PNO2+nWidth_Checkbox+nWidth_Type+nWidth_CheckSum+nWidth_SType+nWidth_Desc+nWidth_Spec+nWidth_Maker+nWidth_Maker2+nWidth_Remark+nWidth_RemarkF

%>
<div style="position:absolute; top:<%=div_top%>; left:<%=div_left%>; z-index:10">
<table width="<%=nWidth_Sum%>px" height="<%=22*(Parts_CNT+1)+100%>px" cellpadding=0 cellspacing=0 border=1 bordercolordark=white bordercolorlight=#999999 style="font-size:8pt; font-family:tahoma; text-align:center; table-layout:fixed">
<form name="frmB_Model_Reg" action="b_model_reg_action.asp" method="post">
<input name="DNO" type="hidden" value="<%=Request("DNO")%>">
<input name="B_Code" type="hidden" value="<%=B_Code%>">
<input name="Diff_YN" type="hidden" value="<%=Diff_YN%>">
<input type="hidden" name="oldParts_CNT" value="<%=Parts_CNT%>">
<input type="hidden" name="oldModel_CNT" value="<%=Model_CNT%>">
<input type="hidden" name="Parts_CNT" value="<%=Parts_CNT%>">
<input type="hidden" name="Model_CNT" value="<%=Model_CNT%>">
<tr height=100px>
<%
arrDNOSUB		= split(Request("DNOSUB"),", ")
'DNOSUB 정렬
dim slDNOSUB
dim strDNOSUB

set slDNOSUB = server.createObject("System.Collections.Sortedlist")

for CNT1 = 0 to ubound(arrDNOSUB)
	if arrDNOSUB(CNT1) <> "" then
		slDNOSUB.add arrDNOSUB(CNT1), ""
	end if
next

for CNT1 = slDNOSUB.count - 1 to 0 step -1
  strDNOSUB = strDNOSUB & slDNOSUB.getKey(CNT1) & ","
next
arrDNOSUB = split(strDNOSUB,",")
set slDNOSUB = nothing

arrDNOCONFIRM	= split(Request("DNOCONFIRM"),", ")



'BOM_CheckSum 시작 
dim strBOM_CheckSum
dim strSW_PNO
dim dicCheckSum
set dicCheckSum =Server.CreateObject("Scripting.Dictionary")

for CNT1 = 1 to Model_CNT
	DNOSUB		= ""
	DNOCONFIRM	= ""
	if Request("DNOSUB") <> "" then
		if CNT1 <= int(Model_CNT-oldModel_CNT) then
			DNOSUB		= ""
		else
			DNOSUB		= arrDNOSUB(CNT1-1-(Model_CNT-oldModel_CNT))
		end if
	end if
	if Request("DNOCONFIRM") <> "" then
		if CNT1 <= int(Model_CNT-oldModel_CNT) then
			DNOCONFIRM	= ""
		else
			DNOCONFIRM	= arrDNOCONFIRM(CNT1-1-(Model_CNT-oldModel_CNT))
		end if
	end if	
	
	if DNOSUB <> "" then
		SQL = "select top 1 * "
		SQL = SQL & "	from tblBOM_CheckSum "
		SQL = SQL & "	where "
		SQL = SQL & "		BOM_Sub_BS_D_No = '"&DNOSUB&"' "
		SQL = SQL & "	order by BC_Apply_Date desc "
		RS2.Open SQL,sys_DBCon
		if not(RS2.Eof or RS2.Bof) then
			strSW_PNO = cstr(trim(RS2("MC_Merge_SW_PNO")))
			if strSW_PNO <> "" then
				strBOM_CheckSum = cstr(RS2("MC_Merge_CheckSum"))
				if dicCheckSum.Exists(strSW_PNO) then
					dicCheckSum.Item(strSW_PNO) = strBOM_CheckSum
				else
					dicCheckSum.Add strSW_PNO, strBOM_CheckSum
				end if	
			end if
			strSW_PNO = cstr(trim(RS2("MICOM1_SW_PNO")))
			if strSW_PNO <> "" then
				strBOM_CheckSum = cstr(RS2("MICOM1_CheckSum"))
				if dicCheckSum.Exists(strSW_PNO) then
					dicCheckSum.Item(strSW_PNO) = strBOM_CheckSum
				else
					dicCheckSum.Add strSW_PNO, strBOM_CheckSum
				end if	
			end if
			strSW_PNO = cstr(trim(RS2("MICOM2_SW_PNO")))
			if strSW_PNO <> "" then
				strBOM_CheckSum = cstr(RS2("MICOM2_CheckSum"))
				if dicCheckSum.Exists(strSW_PNO) then
					dicCheckSum.Item(strSW_PNO) = strBOM_CheckSum
				else
					dicCheckSum.Add strSW_PNO, strBOM_CheckSum
				end if	
			end if
			strSW_PNO = cstr(trim(RS2("EEPROM1_SW_PNO")))
			if strSW_PNO <> "" then
				strBOM_CheckSum = cstr(RS2("EEPROM1_CheckSum"))
				if dicCheckSum.Exists(strSW_PNO) then
					dicCheckSum.Item(strSW_PNO) = strBOM_CheckSum
				else
					dicCheckSum.Add strSW_PNO, strBOM_CheckSum
				end if	
			end if
			strSW_PNO = cstr(trim(RS2("EEPROM2_SW_PNO")))
			if strSW_PNO <> "" then
				strBOM_CheckSum = cstr(RS2("EEPROM2_CheckSum"))
				if dicCheckSum.Exists(strSW_PNO) then
					dicCheckSum.Item(strSW_PNO) = strBOM_CheckSum
				else
					dicCheckSum.Add strSW_PNO, strBOM_CheckSum
				end if	
			end if
		end if
		RS2.Close
	end if
next
'BOM_CheckSum 끝



for CNT1 = 1 to Model_CNT
	DNOSUB		= ""
	DNOCONFIRM	= ""
	if Request("DNOSUB") <> "" then
		if CNT1 <= int(Model_CNT-oldModel_CNT) then
			DNOSUB		= ""
		else
			DNOSUB		= arrDNOSUB(CNT1-1-(Model_CNT-oldModel_CNT))
		end if
	end if
	if Request("DNOCONFIRM") <> "" then
		if CNT1 <= int(Model_CNT-oldModel_CNT) then
			DNOCONFIRM	= ""
		else
			DNOCONFIRM	= arrDNOCONFIRM(CNT1-1-(Model_CNT-oldModel_CNT))
		end if
	end if
%>
	<td width=<%=nWidth_BOMSub_PNO%> onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" width=<%=nWidth_BOMSub_PNO%> align=center valign=top><input class="trans_obj" type="text" name="DNOSUB"	value="<%=DNOSUB%>" readonly style="writing-mode:tb-rl; font-family:돋움; font-size:8pt; width:12px; height:77%; border:0px solid none; text-align:center; background-color:transparent;"><input onclick="toggleTopColor(<%=CNT1%>);" type="checkbox" name="DNOCONFIRM" value="<%=DNOSUB%>"<%if DNOCONFIRM = "Y" then%> checked<%end if%>></td><%COL = COL + 1%>
<%
next
%>
	<td width=<%=nWidth_WorkOrder%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">WORK</td><%COL = COL + 1%>
	<td width=<%=nWidth_Parts_PNO%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">&nbsp;</td><%COL = COL + 1%>
	<td width=<%=nWidth_Parts_PNO2%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">&nbsp;</td><%COL = COL + 1%>
	<td width=<%=nWidth_Checkbox%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">&nbsp;</td><%COL = COL + 1%>
	<td width=<%=nWidth_Type%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">&nbsp;</td><%COL = COL + 1%>
	<td width=<%=nWidth_CheckSum%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">&nbsp;</td><%COL = COL + 1%>
	<td width=<%=nWidth_SType%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">&nbsp;</td><%COL = COL + 1%>
	<td width=<%=nWidth_Desc%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">&nbsp;</td><%COL = COL + 1%>
	<td width=<%=nWidth_Spec%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">&nbsp;</td><%COL = COL + 1%>
	<td width=<%=nWidth_Maker%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">&nbsp;</td><%COL = COL + 1%>
	<td width=<%=nWidth_Maker2%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">&nbsp;</td><%COL = COL + 1%>
	<td width=<%=nWidth_Remark%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">&nbsp;</td><%COL = COL + 1%>
	<td width=<%=nWidth_RemarkF%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">&nbsp;</td><%COL = COL + 1%>
</tr>
<%
ROW = ROW + 1
COL = 1
%>
<tr height=22px>
<%
for CNT2 = 1 to Model_CNT
%>
	<td width=<%=nWidth_BOMSub_PNO%>	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">Q</td><%COL = COL + 1%>
<%
next
%>
	<td width=<%=nWidth_WorkOrder%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">NO</td><%COL = COL + 1%>
	<td width=<%=nWidth_Parts_PNO%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">P/NO</td><%COL = COL + 1%>
	<td width=<%=nWidth_Parts_PNO2%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">P/No2</td><%COL = COL + 1%>
	<td width=<%=nWidth_Checkbox%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">_</td><%COL = COL + 1%>
	<td width=<%=nWidth_Type%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">TYPE</td><%COL = COL + 1%>
	<td width=<%=nWidth_CheckSum%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">C/S</td><%COL = COL + 1%>
	<td width=<%=nWidth_SType%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">SType</td><%COL = COL + 1%>
	<td width=<%=nWidth_Desc%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">DESCRIPTION</td><%COL = COL + 1%>
	<td width=<%=nWidth_Spec%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">SPEC</td><%COL = COL + 1%>
	<td width=<%=nWidth_Maker%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">MAKER</td><%COL = COL + 1%>
	<td width=<%=nWidth_Maker2%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">MAKER2</td><%COL = COL + 1%>
	<td width=<%=nWidth_Remark%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">REMARK</td><%COL = COL + 1%>
	<td width=<%=nWidth_RemarkF%>px	onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" style="background-color:#FFFAA7;">&nbsp;</td><%COL = COL + 1%>
</tr>
<%
ROW = ROW + 1
COL = 1
%>
<%
strNewParts = "-"
for CNT1 = 1 to Parts_CNT
	strMaterial_M_PNO = ""
	strBM_WType = ""
	strBM_Maker = ""
	SQL = "select top 1 Material_M_PNO, BM_WType, BM_Maker "
	SQL = SQL & "	from tblBOM_Mask "
	SQL = SQL & "	where "
	SQL = SQL & "		BOM_Parts_BP_PNO = '"&Request("PNO_"+CSTR(CNT1))&"' and "
	SQL = SQL & "		(BM_Filter = '_' or BM_Filter like '%"&Request("DNO")&"%') "
	SQL = SQL & "	order by BM_Filter desc "
	RS2.Open SQL,sys_DBCon
	if not(RS2.Eof or RS2.Bof) then
		strMaterial_M_PNO = RS2("Material_M_PNO")
		strBM_WType = RS2("BM_WType")
		strBM_Maker = RS2("BM_Maker")
	end if
	RS2.Close
	
	strBOM_CheckSum = ""
	SQL = "select top 1 * "
	SQL = SQL & "	from tblBOM_CheckSum "
	SQL = SQL & "	where "
	SQL = SQL & "		MC_Merge_SW_PNO = '"&Request("PNO_"+CSTR(CNT1))&"' or "
	SQL = SQL & "		MICOM1_SW_PNO = '"&Request("PNO_"+CSTR(CNT1))&"' or "
	SQL = SQL & "		MICOM2_SW_PNO = '"&Request("PNO_"+CSTR(CNT1))&"' or "
	SQL = SQL & "		EEPROM1_SW_PNO = '"&Request("PNO_"+CSTR(CNT1))&"' or "
	SQL = SQL & "		EEPROM2_SW_PNO = '"&Request("PNO_"+CSTR(CNT1))&"' "	
	SQL = SQL & "	order by BC_Apply_Date desc "
	RS2.Open SQL,sys_DBCon
	if not(RS2.Eof or RS2.Bof) then
		if RS2("MC_Merge_SW_PNO") = Request("PNO_"+CSTR(CNT1)) then
			strBOM_CheckSum = RS2("MC_Merge_CheckSum")
		elseif RS2("MICOM1_SW_PNO") = Request("PNO_"+CSTR(CNT1)) then
			strBOM_CheckSum = RS2("MICOM1_CheckSum")
		elseif RS2("MICOM2_SW_PNO") = Request("PNO_"+CSTR(CNT1)) then
			strBOM_CheckSum = RS2("MICOM2_CheckSum")
		elseif RS2("EEPROM1_SW_PNO") = Request("PNO_"+CSTR(CNT1)) then
			strBOM_CheckSum = RS2("EEPROM1_CheckSum")
		elseif RS2("EEPROM2_SW_PNO") = Request("PNO_"+CSTR(CNT1)) then
			strBOM_CheckSum = RS2("EEPROM2_CheckSum")
		end if
	end if
	RS2.Close

	if Diff_YN = "Y" then
		if ucase(Request("NO_"+CSTR(CNT1))) = "R" then
			if dicParts.Exists(Request("PNO_"+CSTR(CNT1))&"_"&Request("Remark_"+CSTR(CNT1))&"_R") then
				response.write "<tr height=22px>"
			else
				response.write "<tr height=22px style='background-color:yellow;'>"
			end if
		else
			if dicParts.Exists(Request("PNO_"+CSTR(CNT1))&"_"&Request("Remark_"+CSTR(CNT1))&"_X") then
				response.write "<tr height=22px>"
			else
				response.write "<tr height=22px style='background-color:yellow;'>"
			end if
		end if
	else
%>
	<tr height=22px>
<%
	end if
	
	strQty = 1
	arrQTY = split(strQty,", ")
	for CNT2 = 1 to Model_CNT - ubound(arrQTY)
		strQty = strQty & ", 0"
	next
	arrQTY = split(strQty,", ")
	
	for CNT2 = 1 to Model_CNT
		if strDNOSUB = "" then
			ComplexKey = ""
		else
			ComplexKey = arrDNOSUB(CNT2-1-(Model_CNT-oldModel_CNT))
		end if
		if ucase(Request("NO_"+CSTR(CNT1))) = "R" then
			ComplexKey = ComplexKey&"_"&Request("PNO_"+CSTR(CNT1))&"_"&Request("Remark_"+CSTR(CNT1))&"_R"
		else
			ComplexKey = ComplexKey&"_"&Request("PNO_"+CSTR(CNT1))&"_"&Request("Remark_"+CSTR(CNT1))&"_X"
		end if
		
		QTY = ""
		if strQty <> "" Then
			if CNT2 <= int(Model_CNT-oldModel_CNT) then
				QTY = ""
			else
				QTY = arrQTY(CNT2-1-(Model_CNT-oldModel_CNT))
				'if isNumeric(QTY) then
				'	if QTY <= 0 then
				'		QTY = ""
				'	end if
				'end if
			end if
		end If
		
		if strDNOSUB = "" then
			Qty = 0
		else
			if ucase(Request("NO_"+CSTR(CNT1)))="R" then
				Qty = Request("QTY_"&arrDNOSUB(CNT2-1-(Model_CNT-oldModel_CNT))&"_"&Request("PNO_"+CSTR(CNT1))&"_"&Request("Remark_"+CSTR(CNT1))&"_R")
			else
				Qty = Request("QTY_"&arrDNOSUB(CNT2-1-(Model_CNT-oldModel_CNT))&"_"&Request("PNO_"+CSTR(CNT1))&"_"&Request("Remark_"+CSTR(CNT1))&"_X")
			end if
		end if
		
		'수량 중복인 경우,
		if dicComplexAcc.Exists(ComplexKey) then
			dicComplexAcc.Item(ComplexKey) = cint(dicComplexAcc.Item(ComplexKey)) + 1
		else
			dicComplexAcc.Add ComplexKey,1
		end if	
		if instr(Qty,",") > 0 then
			arrtempQty = split(Qty,",")
			'response.write ComplexKey & "<br>"
			if ubound(arrtempQty) >= dicComplexAcc.Item(ComplexKey)-1 then
				Qty = trim(arrtempQty(dicComplexAcc.Item(ComplexKey)-1))
			end if
		end if
		
		bChangeQty = "N"
		bChangeMaker = "N"
		if Diff_YN = "Y" then	
			'비교대상 BOM에 같은 키가 있다면.
			
			CNT4 = 1 '중복아이템 카운터
			for CNT3 = 0 to ubound(arrComplexKey) 'Diff대상 키를 조회한다.
				if cstr(arrComplexKey(CNT3)) = ComplexKey then '일치하는 키가 있다면,
					if dicComplexAcc.Item(ComplexKey) = CNT4 then '해당키의 
						if trim(cstr(arrComplexQty(CNT3))) <> trim(cstr(Qty)) then
							bChangeQty = "Y"
						end if
						
						if trim(cstr(arrComplexMaker(CNT3))) <> trim(Request("MAKER_"+CSTR(CNT1))) then
							bChangeMaker = "Y"
						end if
					end if
					CNT4 = CNT4 + 1
				end if
			next
			
		end if
		
		if isNumeric(QTY) then
			if QTY <= 0 then
				QTY = ""
			end if
		end if
%>	
	<td width=<%=nWidth_BOMSub_PNO%> onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center><input class="trans_obj" type="text" name="QTY_<%=CNT1%>"			value="<%=QTY%>" style="font-family:돋움; font-size:8pt; width:100%; border:0px solid none; text-align:center; <%if bChangeQty = "Y" then%>font-weight: bold;  background-color:yellow;<%else%> background-color:transparent;<%end if%>"></td><%COL = COL + 1%>
<%
	next
%>
	<td width=<%=nWidth_WorkOrder%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center><input class="trans_obj" type="text" name="NO_<%=CNT1%>"			value="<%=Request("NO_"+CSTR(CNT1))%>" id="<%=Request("NO_"+CSTR(CNT1))%>" style="font-family:돋움; font-size:8pt; width:100%; border:0px solid none; text-align:center; background-color:transparent;"></td><%COL = COL + 1%>
	<td width=<%=nWidth_Parts_PNO%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center><input class="trans_obj" type="text" name="PNO_<%=CNT1%>"		value="<%=Request("PNO_"+CSTR(CNT1))%>" id="<%=Request("PNO_"+CSTR(CNT1))%>" style="font-family:돋움; font-size:8pt; width:100%; border:0px solid none; text-align:center; background-color:transparent;"></td><%COL = COL + 1%>
	<td width=<%=nWidth_Parts_PNO2%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center><input class="trans_obj" type="text" name="PNO2_<%=CNT1%>" value="<%=strMaterial_M_PNO%>" id="<%=strMaterial_M_PNO%>" style="font-family:돋움; font-size:8pt; width:100%; border:0px solid none; text-align:center; background-color:transparent;"></td><%COL = COL + 1%>
	<td width=<%=nWidth_Checkbox%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center><input class="trans_obj" type="checkbox" name="PNO2PinYN_<%=CNT1%>" <%if Request("PNO2PinYN_"+CSTR(CNT1)) = "Y" then%>checked<%end if%> value="Y" style="font-family:돋움; font-size:8pt; width:100%; border:0px solid none; text-align:center; background-color:transparent;"></td><%COL = COL + 1%>
	<td width=<%=nWidth_Type%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center><input class="trans_obj" type="text" name="WORKTYPE_<%=CNT1%>"  readonly	value="<%=strBM_WType%>" style="font-family:돋움; font-size:8pt; width:100%; border:0px solid none; text-align:center; background-color:<%if strBM_WType="" then%>yellow<%else%>transparent<%end if%>;"></td><%COL = COL + 1%>
	
	<!--<td width=<%=nWidth_CheckSum%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center><input class="trans_obj" type="text" name="CHECKSUM_<%=CNT1%>"		value="<%=Request("CHECKSUM_"+CSTR(CNT1))%>" style="font-family:돋움; font-size:8pt; letter-spacing:2px; width:100%; border:0px solid none; text-align:center; background-color:transparent;"></td><%COL = COL + 1%>-->
	<td width=<%=nWidth_CheckSum%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center><input class="trans_obj" type="text" name="CHECKSUM_<%=CNT1%>"		value="<%=dicCheckSum.Item(cstr(Request("PNO_"+CSTR(CNT1))))%>" style="font-family:돋움; font-size:8pt; letter-spacing:2px; width:100%; border:0px solid none; text-align:center; background-color:transparent;"></td><%COL = COL + 1%>
	
	<td width=<%=nWidth_SType%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center><input class="trans_obj" type="text" name="STYPE_<%=CNT1%>"	value="<%=Request("STYPE_"+CSTR(CNT1))%>" style="font-family:돋움; font-size:8pt; width:100%; border:0px solid none; text-align:center; background-color:transparent;"></td><%COL = COL + 1%>
	<td width=<%=nWidth_Desc%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center><input class="trans_obj" type="text" name="DESCRIPTION_<%=CNT1%>"	value="<%=Request("DESCRIPTION_"+CSTR(CNT1))%>" style="font-family:돋움; font-size:8pt; width:100%; border:0px solid none; text-align:center; <%if instr(strPD_Desc,"-"&lcase(Request("DESCRIPTION_"+CSTR(CNT1))&"-")) = 0 then%>font-weight: bold;  background-color:yellow;<%else%>background-color:transparent;<%end if%>"></td><%COL = COL + 1%>
	<td width=<%=nWidth_Spec%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center><input class="trans_obj" type="text" name="SPEC_<%=CNT1%>"		value="<%=Request("SPEC_"+CSTR(CNT1))%>" style="font-family:돋움; font-size:8pt; width:100%; border:0px solid none; text-align:center; background-color:transparent;"></td><%COL = COL + 1%>
	<td width=<%=nWidth_Maker%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center><input class="trans_obj" type="text" name="MAKER_<%=CNT1%>"		value="<%=Request("MAKER_"+CSTR(CNT1))%>" style="font-family:돋움; font-size:8pt; width:100%; border:0px solid none; text-align:center;<%if bChangeMaker = "Y" then%>font-weight: bold;  background-color:yellow;<%else%> background-color:transparent;<%end if%>"></td><%COL = COL + 1%>
	<td width=<%=nWidth_Maker2%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center><input class="trans_obj" type="text" name="MAKER2_<%=CNT1%>"		value="<%=strBM_Maker%>" style="font-family:돋움; font-size:8pt; width:100%; border:0px solid none; text-align:center;<%if bChangeMaker = "Y" then%>font-weight: bold;  background-color:yellow;<%else%> background-color:<%if strBM_WType="" then%>yellow<%else%>transparent<%end if%>;<%end if%>"></td><%COL = COL + 1%>
<%
	strRemark = Request("REMARK_"+CSTR(CNT1))
	arrRemark = split(strRemark,",")
%>	
	
	<td width=<%=nWidth_Remark%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');" align=center>
			<input class="trans_obj" type="text" name="REMARK_<%=CNT1%>" title="<%=strRemark%>" value="<%=strRemark%>" style="font-family:돋움; font-size:8pt; width:90%; border:0px solid none; text-align:center; background-color:transparent;"></td><%COL = COL + 1%>	 
	<td width=<%=nWidth_RemarkF%>px onMouseOver="javascript:setMouseOver('<%=ROW%>','<%=COL%>');">
<%
	for CNT3=0 to ubound(arrRemark)
		if trim(arrRemark(CNT3)) <> "" then
%>
		<input type="text" id="<%=trim(arrRemark(CNT3))%>" style="width:1px; border:1px solid transparent;">
<%	
		end if
	next
%>
	</td><%COL = COL + 1%>
</tr>
<%
ROW = ROW + 1
COL = 1
%>
<%
next
%>
</form>
</table>
<br>
</div>

<!--체크하면 하이라이트-->
<div style="position:absolute; top:<%=div_top%>; left:<%=div_left%>; z-index:5">
<table width="<%=nWidth_Sum%>px" height="<%=22*(Parts_CNT+1)+100%>px" cellpadding=0 cellspacing=0 border=1 bordercolordark=white bordercolorlight=blue style="font-size:8pt; table-layout:fixed">
<tr height="<%=22*(Parts_CNT+1)+100%>px">
<%
for CNT1 = 1 to Model_CNT
%>
	<td width="<%=nWidth_BOMSub_PNO%>px" id="idCOL_TopOVER_<%=CNT1%>">&nbsp;</td>
<%
next
%>
	<td width=<%=nWidth_WorkOrder%>px id="idCOL_TopOVER_<%=CNT1%>">&nbsp;</td>
	<td width=<%=nWidth_Parts_PNO%>px id="idCOL_TopOVER_<%=CNT1+1%>">&nbsp;</td>
	<td width=<%=nWidth_Parts_PNO2%>px id="idCOL_TopOVER_<%=CNT1+2%>">&nbsp;</td>
	<td width=<%=nWidth_Checkbox%>px id="idCOL_TopOVER_<%=CNT1+3%>">&nbsp;</td>
	<td width=<%=nWidth_Type%>px id="idCOL_TopOVER_<%=CNT1+4%>">&nbsp;</td>
	<td width=<%=nWidth_CheckSum%>px id="idCOL_TopOVER_<%=CNT1+5%>">&nbsp;</td>
	<td width=<%=nWidth_SType%>px id="idCOL_TopOVER_<%=CNT1+6%>">&nbsp;</td>
	<td width=<%=nWidth_Desc%>px id="idCOL_TopOVER_<%=CNT1+7%>">&nbsp;</td>
	<td width=<%=nWidth_Spec%>px id="idCOL_TopOVER_<%=CNT1+8%>">&nbsp;</td>
	<td width=<%=nWidth_Maker%>px id="idCOL_TopOVER_<%=CNT1+9%>">&nbsp;</td>
	<td width=<%=nWidth_Maker2%>px id="idCOL_TopOVER_<%=CNT1+9%>">&nbsp;</td>
	<td width=<%=nWidth_Remark%>px id="idCOL_TopOVER_<%=CNT1+10%>">&nbsp;</td>
	<td width=<%=nWidth_RemarkF%>px id="idCOL_TopOVER_<%=CNT1+11%>">&nbsp;</td>
</tr>
</table>
</div>

<!--마우스 오버 행-->
<div style="position:absolute; top:<%=div_top%>; left:<%=div_left%>; z-index:4">
<table width="<%=nWidth_Sum%>px" height="<%=22*(Parts_CNT+1)+100%>px" cellpadding=0 cellspacing=0 border=1 bordercolordark=white bordercolorlight=gray style="font-size:8pt; table-layout:fixed">
<%
for CNT1 = 1 to Parts_CNT+2
%>
<tr height="<%if CNT1 = 1 then%>100<%else%>22<%end if%>px" id="idROW_OVER_<%=CNT1%>">
<%
	for CNT2 = 1 to Model_CNT
%>
	<td width="<%=nWidth_BOMSub_PNO%>px">&nbsp;</td>
<%
	next
%>
	<td width=<%=nWidth_WorkOrder%>px>&nbsp;</td>
	<td width=<%=nWidth_Parts_PNO%>px>&nbsp;</td>
	<td width=<%=nWidth_Parts_PNO2%>px>&nbsp;</td>
	<td width=<%=nWidth_Checkbox%>px>&nbsp;</td>
	<td width=<%=nWidth_Type%>px>&nbsp;</td>
	<td width=<%=nWidth_CheckSum%>px>&nbsp;</td>
	<td width=<%=nWidth_SType%>px>&nbsp;</td>
	<td width=<%=nWidth_Desc%>px>&nbsp;</td>
	<td width=<%=nWidth_Spec%>px>&nbsp;</td>
	<td width=<%=nWidth_Maker%>px>&nbsp;</td>
	<td width=<%=nWidth_Maker2%>px>&nbsp;</td>
	<td width=<%=nWidth_Remark%>px>&nbsp;</td>
	<td width=<%=nWidth_RemarkF%>px>&nbsp;</td>
</tr>
<%
next
%>
</table>
</div>

<!--마우스 오버 열-->
<div style="position:absolute; top:<%=div_top%>; left:<%=div_left%>; z-index:3">
<table width="<%=nWidth_Sum%>px" height="<%=22*(Parts_CNT+1)+100%>px" cellpadding=0 cellspacing=0 border=1 bordercolordark=white bordercolorlight=blue style="font-size:8pt; table-layout:fixed">
<tr height="<%=22*(Parts_CNT+1)+100%>px">
<%
for CNT1 = 1 to Model_CNT
%>
	<td width="<%=nWidth_BOMSub_PNO%>px" id="idCOL_OVER_<%=CNT1%>">&nbsp;</td>
<%
next
%>
	<td width=<%=nWidth_WorkOrder%>px id="idCOL_OVER_<%=CNT1%>">&nbsp;</td>
	<td width=<%=nWidth_Parts_PNO%>px id="idCOL_OVER_<%=CNT1+1%>">&nbsp;</td>
	<td width=<%=nWidth_Parts_PNO2%>px id="idCOL_OVER_<%=CNT1+2%>">&nbsp;</td>
	<td width=<%=nWidth_Checkbox%>px id="idCOL_OVER_<%=CNT1+3%>">&nbsp;</td>
	<td width=<%=nWidth_Type%>px id="idCOL_OVER_<%=CNT1+4%>">&nbsp;</td>
	<td width=<%=nWidth_CheckSum%>px id="idCOL_OVER_<%=CNT1+5%>">&nbsp;</td>
	<td width=<%=nWidth_SType%>px id="idCOL_OVER_<%=CNT1+6%>">&nbsp;</td>
	<td width=<%=nWidth_Desc%>px id="idCOL_OVER_<%=CNT1+7%>">&nbsp;</td>
	<td width=<%=nWidth_Spec%>px id="idCOL_OVER_<%=CNT1+8%>">&nbsp;</td>
	<td width=<%=nWidth_Maker%>px id="idCOL_OVER_<%=CNT1+9%>">&nbsp;</td>
	<td width=<%=nWidth_Maker2%>px id="idCOL_OVER_<%=CNT1+9%>">&nbsp;</td>
	<td width=<%=nWidth_Remark%>px id="idCOL_OVER_<%=CNT1+10%>">&nbsp;</td>
	<td width=<%=nWidth_RemarkF%>px id="idCOL_OVER_<%=CNT1+11%>">&nbsp;</td>
</tr>
</table>
</div>

<%
if B_Opt_YN = "N" then
%>
<!--화면을 잠시 가림-->
<div style="position:absolute; top:0; left:0; z-index:20">
<table width="<%=nWidth_Sum%>px" height="<%=22*(Parts_CNT+1)+200%>px" cellpadding=0 cellspacing=0 border=1 bordercolordark=white bordercolorlight=white style="font-size:8pt; table-layout:fixed">
<tr height="<%=22*(Parts_CNT+1)+200%>px" bgcolor="white">
	<td width="<%=nWidth_Sum%>px" bgcolor="white">&nbsp;</td>
</tr>
</table>
</div>
<%
end if
%>


<%
dim bakDiff_YN
if bakDiff_YN = "Y" then
'if Diff_YN = "Y" then
%>
<!--마우스 오버 열-->
<div style="position:absolute; top:<%=div_top+22*(Parts_CNT+1)+100+30%>px; left:<%=div_left%>px; z-index:3">
<table width="<%=nWidth_BOMSub_PNO*Model_CNT+nWidth_WorkOrder+nWidth_Parts_PNO+nWidth_Type+nWidth_CheckSum+nWidth_Desc+nWidth_Spec+nWidth_Maker+nWidth_Remark%>px" cellpadding=0 cellspacing=0 border=1 bordercolordark=white bordercolorlight=gray style="font-size:8pt; table-layout:fixed">
<%
	for CNT1 = 1 to Model_CNT
%>
<col width="<%=nWidth_BOMSub_PNO%>px">
<%
	next
%>
<col width=<%=nWidth_WorkOrder%>px>
<col width=<%=nWidth_Parts_PNO%>px>
<col width=<%=nWidth_Type%>px>
<col width=<%=nWidth_CheckSum%>px>
<col width=<%=nWidth_Desc%>px>
<col width=<%=nWidth_Spec%>px>
<col width=<%=nWidth_Maker%>px>
<col width=<%=nWidth_Remark%>px>
<tr style="background-color:lightyellow" height=30px>

	<td colspan=<%=Model_CNT+8%>><b>이전 시방 대비, 삭제된 부품</b></td>

</tr>

<%
'dim Diff_cntModel

'SQL = "select distinct BOM_Sub_BS_D_No from tbBOM_Qty t1 "
'SQL = SQL & "where  "
'SQL = SQL & "	BOM_B_Code = "&Diff_B_Code& " and  "
'SQL = SQL & "	exists (select top 1 *  "
'SQL = SQL & "		from tbBOM_Sub t2  "
'SQL = SQL & "		where st2.BOM_B_Code = "&B_Code&" and t2.BS_D_No = t1.BOM_Sub_BS_D_No)  "
'RS1.Open SQL,sys_DBCon
'Diff_cntModel = RS1(0)
'RS1.Close

	
	SQL = ""
	SQL = SQL & "select  "
	SQL = SQL & "	BOM_Sub_BS_D_No, Parts_P_P_No, "
	SQL = SQL & "	BQ_P_Desc, BQ_P_Spec, BQ_Remark, BQ_Checksum, BQ_P_Maker,  "
	SQL = SQL & "	P_Work_Type = (select top 1 M_Process from tbMaterial where M_P_No = Parts_P_P_No2),  "
	SQL = SQL & "	BQ_Order, BQ_Qty  "
	SQL = SQL & "from "&strTable_Diff_B_Code&" t1 "
	SQL = SQL & "where  "
	SQL = SQL & "	BOM_B_Code = "&Diff_B_Code& " and  "
	SQL = SQL & "	not exists (select top 1 * "
	SQL = SQL & "		from "&strTable_B_Code&" st1 "
	SQL = SQL & "		where BOM_B_Code = "&B_Code&" and  "
	SQL = SQL & "			st1.BOM_Sub_BS_D_No = t1.BOM_Sub_BS_D_No and   "
	SQL = SQL & "			st1.Parts_P_P_No = t1.Parts_P_P_No and "
	SQL = SQL & "			st1.BQ_Remark = t1.BQ_Remark) and  "
	SQL = SQL & "	exists (select top 1 *  "
	SQL = SQL & "		from "&strTabl_B_Codee&" st2  "
	SQL = SQL & "		where st2.BOM_B_Code = "&B_Code&" and st2.BOM_Sub_BS_D_No = t1.BOM_Sub_BS_D_No)  "
	'response.write SQL
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
%>
<tr height="22px">
	<td colspan=<%=Model_CNT+8%>>해당하는 내용이 없습니다.</td>
</tr>
<%
	else
		do until RS1.Eof
%>
<tr height="22px">

<%
			for CNT1 = 1 to Model_CNT
				if RS1("BOM_Sub_BS_D_No") = arrDNOSUB(CNT1-1) then
					RS1.MoveNext
%>	
			<td width="<%=nWidth_BOMSub_PNO%>px"><%=RS1("BQ_Qty")%></td>
<%
				else
%>
			<td width="<%=nWidth_BOMSub_PNO%>px">0</td>
<%
				end if
			next
%>
	<td width=<%=nWidth_WorkOrder%>px><%=RS1("BQ_Order")%>&nbsp;</td>
	<td width=<%=nWidth_Parts_PNO%>px><%=RS1("Parts_P_P_No")%>&nbsp;</td>
	<td width=<%=nWidth_Type%>px><%=RS1("P_Work_Type")%>&nbsp;</td>
	<td width=<%=nWidth_CheckSum%>px><%=RS1("BQ_Checksum")%>&nbsp;</td>
	<td width=<%=nWidth_Desc%>px><%=RS1("BQ_P_Desc")%>&nbsp;</td>
	<td width=<%=nWidth_Spec%>px><%=RS1("BQ_P_Spec")%>&nbsp;</td>
	<td width=<%=nWidth_Maker%>px><%=RS1("BQ_P_Maker")%>&nbsp;</td>
	<td width=<%=nWidth_Remark%>px><%=RS1("BQ_Remark")%>&nbsp;</td>
</tr>
<%
	
			'RS1.MoveNext
		loop
	end if
	RS1.Close	
%>

</table>
</div>
<%

end if
%>
<%
if Diff_YN = "Y" then
%>


<!--마우스 오버 열-->
<div style="position:absolute; top:<%=div_top+22*(Parts_CNT+1)+100+30%>px; left:<%=div_left%>px; z-index:3">
<table width="<%=nWidth_Diff_Assy_PNO+nWidth_Parts_PNO+nWidth_Diff_Qty+nWidth_WorkOrder+nWidth_Type+nWidth_CheckSum+nWidth_Desc+nWidth_Spec+nWidth_Maker+nWidth_Remark%>px" cellpadding=0 cellspacing=0 border=1 bordercolordark=white bordercolorlight=gray style="font-size:8pt; table-layout:fixed">
<col width=<%=nWidth_Diff_Assy_PNO%>px>
<col width=<%=nWidth_Parts_PNO%>px>
<col width=<%=nWidth_Diff_Qty%>px>
<col width=<%=nWidth_WorkOrder%>px>
<col width=<%=nWidth_Type%>px>
<col width=<%=nWidth_CheckSum%>px>
<col width=<%=nWidth_Desc%>px>
<col width=<%=nWidth_Spec%>px>
<col width=<%=nWidth_Maker%>px>
<col width=<%=nWidth_Remark%>px>
<tr style="background-color:lightyellow" height=30px>
	<td colspan=10><b>이전 시방 대비, 삭제된 부품</b></td>
</tr>
<tr style="background-color:lightyellow" height=30px>
	<td width=<%=nWidth_Diff_Assy_PNO%>px>Assy P/N</td>
	<td width=<%=nWidth_Parts_PNO%>px>Part P/N</td>
	<td width=<%=nWidth_Diff_Qty%>px>Qty</td>
	<td width=<%=nWidth_WorkOrder%>px>W/O</td>
	<td width=<%=nWidth_Type%>px>TYPE</td>
	<td width=<%=nWidth_CheckSum%>px>C/S</td>
	<td width=<%=nWidth_Desc%>px>DESC</td>
	<td width=<%=nWidth_Spec%>px>SPEC</td>
	<td width=<%=nWidth_Maker%>px>MAKER</td>
	<td width=<%=nWidth_Remark%>px>REMARK</td>
</tr>
<%
	SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&Diff_B_Code	
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
	else
		if RS1("B_Version_Current_YN") = "Y" then
			strTable_Diff_B_Code = "tbBOM_Qty"
		else
			strTable_Diff_B_Code = "tbBOM_Qty_Archive"
		end if
	end if
	RS1.Close
	
	SQL = "select B_Version_Current_YN from tbBOM where B_Code = "&B_Code	
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
	else
		if RS1("B_Version_Current_YN") = "Y" then
			strTable_B_Code = "tbBOM_Qty"
		else
			strTable_B_Code = "tbBOM_Qty_Archive"
		end if
	end if
	RS1.Close
	
	SQL = ""
	SQL = SQL & "select  "
	SQL = SQL & "	BOM_Sub_BS_D_No, Parts_P_P_No, "
	SQL = SQL & "	BQ_P_Desc, BQ_P_Spec, BQ_Remark, BQ_Checksum, BQ_P_Maker,  "
	SQL = SQL & "	P_Work_Type = (select top 1 M_Process from tbMaterial where M_P_No = Parts_P_P_No2),  "
	SQL = SQL & "	BQ_Order, BQ_Qty  "
	SQL = SQL & "from "&strTable_Diff_B_Code&" t1 "
	SQL = SQL & "where  "
	SQL = SQL & "	BOM_B_Code = "&Diff_B_Code& " and  "
	
	SQL = SQL 
	SQL = SQL & "	not exists (select top 1 * "
	SQL = SQL & "		from "&strTable_B_Code&" st1 "
	SQL = SQL & "		where BOM_B_Code = "&B_Code&" and  "
	SQL = SQL & "			st1.BOM_Sub_BS_D_No = t1.BOM_Sub_BS_D_No and   "
	SQL = SQL & "			st1.Parts_P_P_No = t1.Parts_P_P_No and "
	SQL = SQL & "			st1.BQ_Remark = t1.BQ_Remark) and  "
	SQL = SQL & "	exists (select top 1 *  "
	SQL = SQL & "		from "&strTable_B_Code&" st2  "
	SQL = SQL & "		where st2.BOM_B_Code = "&B_Code&" and st2.BOM_Sub_BS_D_No = t1.BOM_Sub_BS_D_No)  "

	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
%>
<tr height="22px">
	<td colspan=10>해당하는 내용이 없습니다.</td>
</tr>
<%
	else
		do until RS1.Eof
%>
<tr height="22px">
	<td width=<%=nWidth_Diff_Assy_PNO%>px><%=RS1("BOM_Sub_BS_D_No")%>&nbsp;</td>
	<td width=<%=nWidth_Parts_PNO%>px><%=RS1("Parts_P_P_No")%>&nbsp;</td>
	<td width=<%=nWidth_Diff_Qty%>px><%=RS1("BQ_Qty")%>&nbsp;</td>
	<td width=<%=nWidth_WorkOrder%>px><%=RS1("BQ_Order")%>&nbsp;</td>
	<td width=<%=nWidth_Type%>px><%=RS1("P_Work_Type")%>&nbsp;</td>
	<td width=<%=nWidth_CheckSum%>px><%=RS1("BQ_Checksum")%>&nbsp;</td>
	<td width=<%=nWidth_Desc%>px><%=RS1("BQ_P_Desc")%>&nbsp;</td>
	<td width=<%=nWidth_Spec%>px><%=RS1("BQ_P_Spec")%>&nbsp;</td>
	<td width=<%=nWidth_Maker%>px><%=RS1("BQ_P_Maker")%>&nbsp;</td>
	<td width=<%=nWidth_Remark%>px><%=RS1("BQ_Remark")%>&nbsp;</td>
</tr>
<%
			RS1.MoveNext
		loop
	end if
	RS1.Close	
%>

</table>
</div>
<%
end if
%>

<script language="javascript">
function fRun()
{
	if(document.readyState == "complete")
	{
		Form_Submit();
	}
	else
	{
		setTimeout("fRun()",2000);
	}
}
<%
if B_Opt_YN = "N" then
%>
fRun();
<%
end if
%>



</script>

<%
if Diff_YN = "Y" then
	set dicComplexAcc = nothing
	set dicParts = nothing
end if
set RS1 = nothing
set RS2 = nothing
%>

<%
if B_Opt_YN = "Y" then
%>
<script language='javascript'>
function customShortCut()
{
	$(document).keydown(function(event) {
	    if (event.altKey && event.ctrlKey &&event.which === 70) //F
	    {
	        //alert('Ctrl + Alt + F pressed!');
			$('#idBOM_XLS').trigger('click');
	        e.preventDefault();
	    }
	    else if (event.altKey && event.ctrlKey &&event.which === 85) //U
	    {
	        //alert('Ctrl + Alt + U pressed!');
	        XLS_Form_Submit();
	        e.preventDefault();
	    }
	    else if (event.altKey && event.ctrlKey &&event.which === 83) //S
	    {
	    	//alert('Ctrl + Alt + S pressed!');
	    	Form_Submit();
	        e.preventDefault();
	    }
	});
}

function RPA_Alert()
{
	if(confirm('Ready'))
		$("#strFindString").focus();
}

$('document').ready(function(){
	<%if instr("-shindk-shindh-rnd-","-"&lcase(request.cookies("ADMIN")("M_ID"))&"-") > 0 then%>setTimeout(RPA_Alert, 1000);<%end if%>
	customShortCut();
});
</script>
<%
end if
%>
<!-- #include virtual="/header/layout_tail.asp" -->
<!-- #include virtual="/header/html_tail.asp" -->
<!-- #include virtual="/header/db_tail.asp" -->
<!-- include virtual = "/header/session_check_tail.asp" -->
