<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
dim RS1
dim RS2
dim SQL

dim s_Work_Date
dim s_Line

s_Work_Date = date()
s_Line = request("s_Line")

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")

calcNow = left(FormatDateTime(now(),4),2)*60 + right(FormatDateTime(now(),4),2)
calcNow = calcNow * 60
calcNow = getRestedCalcNow(calcNow)
function getRestedCalcNow(calcNow)
	'���� �ð� ���̶��, ���� �ð� ���� ���·� ����
	if calcNow > 620*60 and calcNow <= 630*60 then '10�� 20�� ~ 30�� 
		calcNow = 620*60
	end if
	if calcNow > 750*60 and calcNow <= 790*60 then '12�� 30��~13�� 10�� 
		calcNow = 750*60
	end if
	if calcNow > 910*60 and calcNow <= 920*60 then '3�� 10�� ~ 20�� 
		calcNow = 910*60
	end if
	if calcNow > 1040*60 and calcNow <= 1060*60 then '5�� 20��~40��
		calcNow = 1040*60
	end if
	
	'���� �ð��� ��ģ �� ��ŭ ���� �ð� ����
	if calcNow > 1060*60 then '17�� 40��
		calcNow = calcNow - (20+10+40+10)*60
	elseif calcNow > 920*60 then '15�� 20�� �����Ѱ�, �����Ѱ�, ���� �Ѱ� ���� 
		calcNow = calcNow - (10+40+10)*60
	elseif calcNow > 790*60 then '13�� 10�� �������½ð� + ��������
		calcNow = calcNow - (40+10)*60
	elseif calcNow > 630*60 then '10�� 30�� �������½ð� �ϳ� ����
		calcNow = calcNow - 10*60
	end if
	getRestedCalcNow = calcNow
end function

'
function MakePlanTable()
	dim CNT1
		
	dim SQL
	dim RS1
	dim tQty
	
	dim BS_D_No
	dim B_D_No
	dim oldB_D_No
	dim lenDiff
	
	dim PSP_Count
	dim BP_PPH
	dim PSP_ST
	dim ChangeOverHead
	
	dim accSec
	dim accQty
	
	dim calcPRD_Start
	
	dim strSimilar
	dim arrSimilar
	dim arrSimilarDetail
	strSimilar = strSimilar & "EBR715081$-EBR741529-//"
	strSimilar = strSimilar & "EBR644383$-EBR662511-EBR737344-//"
	strSimilar = strSimilar & "EBR624063$-EBR654006-//"
	strSimilar = strSimilar & "EBR391877$-EBR622045-EBR784017-EBR806189-EBR813330-//"
	strSimilar = strSimilar & "6871A10233$-EBR535783-EBR610631-//"
	strSimilar = strSimilar & "6871A20181$-EBR515892-EBR515959-//"
	strSimilar = strSimilar & "EBR337500$-EBR420488-EBR489280-EBR568373-EBR622537-//"
	strSimilar = strSimilar & "EBR351584$-EBR412888-EBR420085-EBR442026-EBR564637-EBR577053-EBR577370-EBR618909-EBR740459-EBR743984-EBR775510-EBR775954-EBR775955-EBR779234-EBR779637-EBR784020-EBR788952-EBR792664-EBR792665-EBR798438-EBR801090-EBR815512-//"
	strSimilar = strSimilar & "EBR355200$-EBR399048-EBR631040-EBR715171-EBR774722-EBR794405-//"
	strSimilar = strSimilar & "6871A20156$-EBR356394-EBR441696-EBR604434-//"
	strSimilar = strSimilar & "6871A20889$-6871A20891-//"
	strSimilar = strSimilar & "6871A10161$-6871A20684-//"
	strSimilar = strSimilar & "6871A10117$-6871A20679-EBR431272-//"
	strSimilar = strSimilar & "6871A20272$-6871A20547-//"
	strSimilar = strSimilar & "6871A20294$-6871A20309-6871A20310-6871A20311-6871A20312-6871A20373-6871A20493-6871A20494-6871A20495-6871A20562-6871A20565-//"
	strSimilar = strSimilar & "6871A20225$-6871A20229-6871A20235-//"
	strSimilar = strSimilar & "6871A20107$-6871A20222-//"
	strSimilar = strSimilar & "6871A20164$-6871A20216-6871A20218-6871A20220-6871A20240-//"
	strSimilar = strSimilar & "6871A20146$-6871A20160-6871A20212-6871A20232-6871A20352-//"
	strSimilar = strSimilar & "6871A10042$-6871A20040-6871A20067-6871A20082-6871A20152-6871A20158-6871A20415-//"
	strSimilar = strSimilar & "6871A20007$-6871A20008-//"
	strSimilar = strSimilar & "6871A10231$-6871A10362-6871A10363-6871A10366-EBR341635-//"
	strSimilar = strSimilar & "6871A10158$-6871A10209-6871A10338-6871A10370-//"
	strSimilar = strSimilar & "6871A10108$-6871A10167-//"
	strSimilar = strSimilar & "6871A10105$-6871A10165-//"
	strSimilar = strSimilar & "6871A10056$-6871A10143-//"
	strSimilar = strSimilar & "6871A00089$-6871A10140-6871A10342-//"
	strSimilar = strSimilar & "6871A01002$-6871A10070-6871A10188-6871A20188-EBR615952-//"
	strSimilar = strSimilar & "6871A10008$-6871A10038-6871A10040-6871A10109-6871A10116-//"
	strSimilar = strSimilar & "6871A10009$-6871A10020-6871A10023-6871A10026-6871A10030-6871A10048-//"
	strSimilar = strSimilar & "6871A00012$-6871A00090-6871A10131-//"
	strSimilar = strSimilar & "6871A00007$-6871A00009-6871A10089-6871A10106-6871A10107-6871A10124-6871A10125-6871A10148-6871A10166-6871A10187-6871A10217-//"
	arrSimilar = split(strSimilar,"//")
	set RS1 = server.CreateObject("ADODB.RecordSet")
	
	'���� ����ð� ���ϱ�
	SQL = "select minPRD_Input_Time = min(PRD_Input_Time) from tbPWS_Raw_Data "
	SQL = SQL & "where "&vbcrlf
	SQL = SQL & "	PRD_Input_Date = '"&s_Work_Date&"' and PRD_Line = '"&s_Line&"' and "&vbcrlf
	SQL = SQL & "	PRD_Input_Date is not null"&vbcrlf
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		calcPRD_Start = 30000
	else
		calcPRD_Start = (int(left(RS1("minPRD_Input_Time"),2)*60) + int(right(RS1("minPRD_Input_Time"),2)))*60
		if calcPRD_Start < 30000 then '8�� 20�� ������ ������ ���۵Ǿ��ٸ�
			calcPRD_Start = 30000 '�׳� 8�� 20������ ����
		end if
	end if
	RS1.Close
	
	'��ȹ ��������
	tQty = 0
	accSec = 0
	accQty = 0
	SQL = ""
	SQL = SQL & "select "
	SQL = SQL & "	t1.BOM_Sub_BS_D_No, "
	SQL = SQL & "	t1.PSP_Count, "
	SQL = SQL & "	BP_PPH = isnull((select top 1 t2.BP_PPH from tbBOM_PPH t2 where t2.BOM_Sub_BS_D_No = t1.BOM_Sub_BS_D_No),0) "
	SQL = SQL & "from tbProcess_State_Plan t1 "
	SQL = SQL & "where t1.PSP_Line = '"&s_Line&"' and t1.PSP_Work_Date = '"&s_Work_Date&"' "
	SQL = SQL & "order by PSP_Code asc "
	RS1.Open SQL,sys_DBCon
	
	ChangeOverHead	= 0
	oldB_D_No		= ""
	do until RS1.Eof 
		
		'��/�ɼ� ü���� üũ
		B_D_No	= RS1("BOM_Sub_BS_D_No")
		'���� ó���� �н�
		if oldB_D_No <> "" then
			'�ɼǹ�ȣ �����
			if isnumeric(left(B_D_No,4)) then '6871�迭�̶��
				B_D_No = left(B_D_No,10)
			else
				B_D_No = left(B_D_No,9)
			end if
			
			'Ȥ�� �����ü�������� ��ü���� ���� Ȯ������
			for CNT1 = 0 to ubound(arrSimilar) - 1
				arrSimilarDetail = split(arrSimilar(CNT1),"$")
				
				'���� ����� ����Ʈ�� �ִٸ�, ��ǥ��Ʈ�ѹ��� �ٲ۴�
				if instr(arrSimilarDetail(1), "-"&B_D_No&"-") > 0 then
					B_D_No = arrSimilarDetail(0)
				end if
			next
			
			'���� ���� ���̶� �⺻���� �ɼ��� �ٲ����.
			ChangeOverHead = 1
			if B_D_No <> oldB_D_No then
				ChangeOverHead = 4
			end if
		end if
		oldB_D_No = B_D_No
		
		BS_D_No		= RS1("BOM_Sub_BS_D_No")
		PSP_Count 	= RS1("PSP_Count") '��ȹ����
		BP_PPH		= RS1("BP_PPH")
		if BP_PPH = 0 then
			BP_PPH = 300
		end if
		
		PSP_ST	= cint(3600 / BP_PPH) '���� ����ð�
		
		'�̹� ���ڵ��� �� �����ʿ�ð��� accSec�� ���� / �� ��ȹ������ accQty�� ���� / ������� �ݿ�
		accSec = accSec + (PSP_Count * PSP_ST) + (ChangeOverHead*60)
		accQty = accQty + PSP_Count
		
		'������ �ʿ�ð��� 2�ð��� ���ٸ�
		if accSec = 2*60*60 then
			accSec = 0
			accQty = 0
			strPlanTable = strPlanTable & BS_D_No & "$" & PSP_Count & "//"
			strPlanTable = strPlanTable & "�޽�" & "$" & "0" & "//"
			splitYN = "N"
			RS1.MoveNext
		'2�ð��� �������ٸ�
		elseif 2*60*60 < accSec then
			'��Ȯ�� ������ ����ϱ� ����...
			accSec = accSec - (PSP_Count * PSP_ST) '���������� ������ �����ʿ�ð��� ����.
			accQty = accQty - PSP_Count '���������� ������ ��ȹ������ ����.
			
			do until 2*60*60 < accSec '�ִ���갡�ɼ�������
				accSec = accSec + PSP_ST '����ð��� ���Ѵ�
				accQty = accQty + 1 '������ �ϳ��� �ø��� 
				accQtyPre = accQtyPre + 1
			loop
			
			accSec = 0
			accQty = 0
			strPlanTable = strPlanTable & BS_D_No & "$" & accQtyPre & "//"
			strPlanTable = strPlanTable & "�޽�" & "$" & "0" & "//"
			PSP_Count = PSP_Count - accQtyPre
			splitYN = "Y"
		else
			strPlanTable = strPlanTable & BS_D_No & "$" & PSP_Count & "//"
			splitYN = "N"
			RS1.MoveNext
		end if
		
		
	loop
	RS1.Close

	set RS1 = nothing
end function

if isnull(getTargetQty) or getTargetQty = "" then
	getTargetQty = accQty
end if
RS1.Close




set RS1 = nothing

response.end
%>

<script language="javascript">
var strHTML = "";
strHTML += "<table width=100% cellpadding=0 cellspacing=1 bgcolor='white' style='color:white;font-size:37px;text-align:center;font-weight:bold'>";
strHTML += "<col></col>";
strHTML += "<col width=150px></col>";
strHTML += "<col width=150px></col>";
strHTML += "<col width=150px></col>";
strHTML += "<col width=300px></col>";
strHTML += "<col width=150px></col>";
<%
for CNT1=0 to ubound(arrPWS_Opt_Data)

	if arrPWS_Opt_Data(CNT1,0) = "" then
	

	elseif arrPWS_Opt_Data(CNT1,0) = "���۾�" and cint(replace(arrPWS_Opt_Data(CNT1,3),":","")) > cint(replace(FormatDateTime(now(),4),":","")) then
		
	elseif arrPWS_Opt_Data(CNT1,0) = "���۾�" then
%>
strHTML += "<tr bgcolor=black>";
strHTML += "	<td colspan=4><%if arrPWS_Opt_Data(CNT1,5)="�޽�" then%>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��<%else%><%=arrPWS_Opt_Data(CNT1,5)%><%end if%></td>";
strHTML += "	<td><%=left(arrPWS_Opt_Data(CNT1,3),2)%>:<%=right(arrPWS_Opt_Data(CNT1,3),2)%> - <%=left(arrPWS_Opt_Data(CNT1,4),2)%>:<%=right(arrPWS_Opt_Data(CNT1,4),2)%></td>";
strHTML += "	<td>&nbsp;</td>";
strHTML += "</tr>";
<%
	else
		'�������� ������ �迭�̶��

		if arrPWS_Opt_Data(CNT1,5) = "raw" then
			
			nRawCount = nRawCount + 1
		
			if CurrentRecord = int(CNT1) then
%>
strHTML += "<tr bgcolor=green style='color:white'>";
<%
			else
%>
strHTML += "<tr bgcolor=black>";
<%
			end if
%>
strHTML += "	<td><span style='cursor:hand' onclick=\"javascript:Pop_Print('<%=arrPWS_Opt_Data(CNT1,0)%>');\"><%=arrPWS_Opt_Data(CNT1,0)%></span></td>";	//�۾�����
strHTML += "	<td><%=arrPWS_Opt_Data(CNT1,6)%></td>";	//��ȹ����
strHTML += "	<td><%=arrPWS_Opt_Data(CNT1,1)%></td>"; //��������
strHTML += "	<td><%=arrRemain(CNT1)%></td>";	//��ȹ-��������
strHTML += "	<td><%=left(arrPWS_Opt_Data(CNT1,3),2)%>:<%=right(arrPWS_Opt_Data(CNT1,3),2)%> - <%=left(arrPWS_Opt_Data(CNT1,4),2)%>:<%=right(arrPWS_Opt_Data(CNT1,4),2)%></td>";	//�۾��ð�
strHTML += "	<td align=right><%=arrRndRate(CNT1)%>&nbsp;</td>";	
strHTML += "</tr>";
<%
		'���� ������ ���� ��ȹ�� �ִ� ���ڵ� ���
		else
%>
strHTML += "<tr bgcolor=black>";
strHTML += "	<td><span style='cursor:hand' onclick=\"javascript:Pop_Print('<%=arrPWS_Opt_Data(CNT1,0)%>');\"><%=arrPWS_Opt_Data(CNT1,0)%></span></td>";	//�۾�����
strHTML += "	<td><%=arrPWS_Opt_Data(CNT1,6)%></td>";	//��ȹ����
strHTML += "	<td>0</td>";	//��������
strHTML += "	<td><%=arrPWS_Opt_Data(CNT1,6)%></td>";	//�ܷ�
strHTML += "	<td><%=left(arrPWS_Opt_Data(CNT1,7),2)%>:<%=right(arrPWS_Opt_Data(CNT1,7),2)%> - <%=left(arrPWS_Opt_Data(CNT1,8),2)%>:<%=right(arrPWS_Opt_Data(CNT1,8),2)%></td>"; //��ȹ�ð�
strHTML += "	<td align=center>-</td>";	//�޼���
strHTML += "</tr>";
<%
		end if
	end if
	
	oldPWS_Opt_Data_5 = arrPWS_Opt_Data(CNT1,5)
next

set RS1 = Nothing
%>

parent.idContent.innerHTML = strHTML;
var nScroll = <%=nRawCount%>*43-600;
if(nScroll > 0)
	parent.scrollTo(0,nScroll);

function reload_handle()
{
	location.reload();
}

</script>


<!-- include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->


