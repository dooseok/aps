<!-- #include virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->
<%
Response.Buffer = False

rem SQL ����ȭ
rem update tbBOM_Qty set BQ_Use_YN = '' where BQ_Use_YN is null

dim BOM_B_Code
dim BOM_Sub_BS_D_No
dim BQ_Code
dim BOM_Sub_BS_Code
dim Parts_P_P_No
dim Parts_P_P_No2
dim Parts_P_P_No2_PinYN
dim BQ_Qty
dim BQ_Use_YN
dim BQ_Order
dim BQ_Remark
dim BQ_CheckSum
dim BQ_P_Desc
dim BQ_P_Spec
dim BQ_P_Maker
dim BQ_ready2del_YN
		
dim strTable
strTable=getTable()
dim oldcode
dim oldb_code
dim oldremark

'���� ����Ʈ��, �� ����ũ�� ǰ���� 2���ְ�, ���ٿ� R�� �ִ� ��ü����̴�.
'strTable=strTable&"3961/_/0CQ22418679/_//_/Capacitor,Film,Box/_/C01J,C02J/_/1/|/"
dim b_code		'0
dim p_p_no		'1
dim isR			'2
dim desc		'3
dim remark		'4
dim reorder		'5

dim bs_d_no

dim arrTable
arrTable = split(strTable,"/|/")
dim arrRecord

dim strSubTable

dim CNT1
dim SQL
dim RS1
set RS1 = server.CreateObject("ADODB.RecordSet")
SQL = "select code,strTable from tbBOM_Table where code <= 7800 and bRun='0' and code > "&request("preCode")&" order by code asc"
RS1.Open SQL,sys_DBCon
do until RS1.Eof
		arrRecord = split(RS1("strTable"),"/_/")
		b_code = arrRecord(0)
		remark = arrRecord(4)
		
		'������ �޶����ų� ����ũ�� �޶�����,
		if oldb_code <> b_code or oldremark <> remark then '���ο� �׸�
			if strSubTable <> "" then
				call updateDB_DH_Check(strSubTable,oldb_code)	'���� ����� ReOrder ó��
				response.redirect "bom_re_action.asp?preCode="&oldcode
			end if
			
			
			strSubTable = RS1("strTable")&"/|/" '������ �� ���� ����
		else '������ �����̸�,
			strSubTable = strSubTable & RS1("strTable")&"/|/" '������ �ٿ� �ֱ�
		end if
		
		oldcode = RS1("code")
		oldb_code = b_code
		oldremark = remark
		SQL = "update tbBOM_Table set bRun = '1' where code = "&RS1("Code")
		sys_DBCon.execute(SQL)
	RS1.MoveNext
loop
RS1.Close
call updateDB_DH_Check(strSubTable,oldb_code)	'���� ����� ReOrder ó�� (������ ��� ó��)
response.write "DONE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
set RS1 = nothing
%>
	
<%
sub updateDB_DH_Check(strSubTable,oldb_code)
	
	dim RS1
	dim SQL
	dim CNT3
	
	dim arrSubRecord
	dim arrSubTable
	dim strRetable
	
	arrSubTable = split(strSubTable,"/|/")
	
	updateDB(strSubTable)
	
	if instr(strSubTable,"/_/DH/|/") > 0 then
		set RS1 = server.CreateObject("ADODB.RecordSet")
		SQL = "select B_Code from tbBOM where b_d_no = (select top 1 b_d_no from tbBOM where b_code = "&oldb_code&") and (B_Version_Current_YN <> 'Y' or B_Version_Current_YN is null)"
		RS1.Open SQL, sys_DBCon
		do until RS1.Eof
			for CNT3 = 0 to ubound(arrSubTable)-1
				arrSubRecord = split(arrSubTable(CNT3),"/_/")
				strRetable = strRetable & RS1("B_Code")&"/_/"
				strRetable = strRetable & arrSubRecord(1)&"/_/"
				strRetable = strRetable & arrSubRecord(2)&"/_/"
				strRetable = strRetable & arrSubRecord(3)&"/_/"
				strRetable = strRetable & arrSubRecord(4)&"/_/"
				strRetable = strRetable & arrSubRecord(5)&"/|/"
			next
			updateDB(strRetable)
			strRetable = ""
			RS1.MoveNext
		loop
		RS1.Close
		set RS1 = nothing
	end if
end sub
%>
	
<%
sub updateDB(strSubTable)
	dim SQL
	dim RS1
	dim RS2

	dim b_code		'0
	dim p_p_no		'1
	dim isR			'2
	dim desc		'3
	dim remark		'4
	dim reorder		'5
	
	dim strIDX
	strIDX = now()
	
	dim CNT2
	dim arrSubTable
	dim arrSubRecord
	arrSubTable = split(strSubTable,"/|/")
	
	set RS1 = server.CreateObject("ADODB.RecordSet")
	set RS2 = server.CreateObject("ADODB.RecordSet")
	for CNT2 = 0 to ubound(arrSubTable)-1
		'response.write arrSubTable(CNT2) &"<br>"
		arrSubRecord = split(arrSubTable(CNT2),"/_/")
		b_code = arrSubRecord(0)
		p_p_no = arrSubRecord(1)
		isR = arrSubRecord(2)
		desc = arrSubRecord(3)
		remark = arrSubRecord(4)
		reorder = arrSubRecord(5)
		
		'�ϴ� �ش�Ǵ� 'BOM_QTY������ BOM_SORT�� ��� �����Ѵ�. ������ 1�྿��!
		SQL = "select bs_d_no from tbBOM_Sub where BOM_B_Code = "&b_code
		RS1.Open SQL, sys_DBCon
		do until RS1.Eof

			SQL = "select top 1 * from tbBOM_QTY where "
			SQL = SQL & "BOM_B_Code = "&b_code&" and "
			SQL = SQL & "BOM_Sub_BS_D_No = '"&RS1("BS_D_No")&"' and "
			SQL = SQL & "Parts_P_P_No = '"&p_p_no&"' and "
			if isR = "R" then
				SQL = SQL & "BQ_Order = 'R' and "
			else
				SQL = SQL & "BQ_Order <> 'R' and "
			end if
			SQL = SQL & "BQ_Remark = '"&remark&"'"
			RS2.Open SQL,sys_DBCon
			if not(RS2.Eof or RS2.Bof) then
				
				
				'BOM_QTY������ BOM_SORT�� ���
				SQL = "insert into tbBOM_QTY_Sort (BOM_B_Code, BOM_Sub_BS_D_No,BQ_Code,BOM_Sub_BS_Code,Parts_P_P_No,Parts_P_P_No2,Parts_P_P_No2_PinYN,BQ_Qty,BQ_Use_YN,BQ_Order,BQ_Remark,BQ_CheckSum,BQ_P_Desc,BQ_P_Spec,BQ_P_Maker,BQ_ready2del_YN,idx,reorder) values ("
				SQL = SQL & freplace(RS2("BOM_B_Code"),"'","''")&","
				SQL = SQL & "'"&freplace(RS2("BOM_Sub_BS_D_No"),"'","''")&"',"
				SQL = SQL & freplace(RS2("BQ_Code"),"'","''")&","
				SQL = SQL & freplace(RS2("BOM_Sub_BS_Code"),"'","''")&","
				SQL = SQL & "'"&freplace(RS2("Parts_P_P_No"),"'","''")&"',"
				SQL = SQL & "'"&freplace(RS2("Parts_P_P_No2"),"'","''")&"',"
				SQL = SQL & "'"&freplace(RS2("Parts_P_P_No2_PinYN"),"'","''")&"',"
				SQL = SQL & freplace(RS2("BQ_Qty"),"'","''")&","
				SQL = SQL & "'"&freplace(RS2("BQ_Use_YN"),"'","''")&"',"
				SQL = SQL & "'"&freplace(RS2("BQ_Order"),"'","''")&"',"
				SQL = SQL & "'"&freplace(RS2("BQ_Remark"),"'","''")&"',"
				SQL = SQL & "'"&freplace(RS2("BQ_CheckSum"),"'","''")&"',"
				SQL = SQL & "'"&freplace(RS2("BQ_P_Desc"),"'","''")&"',"
				SQL = SQL & "'"&freplace(RS2("BQ_P_Spec"),"'","''")&"',"
				SQL = SQL & "'"&freplace(RS2("BQ_P_Maker"),"'","''")&"',"
				SQL = SQL & "'"&freplace(RS2("BQ_ready2del_YN"),"'","''")&"',"
				SQL = SQL & "'"&strIDX&"',"
				SQL = SQL & reorder&")"
				sys_DBCon.execute(SQL)
				
				'BOM_QTY�� �����
				SQL = "delete from tbBOM_Qty where BQ_Code = "&RS2("BQ_Code")
				sys_DBCon.execute(SQL)
			end if		
			RS2.Close
			
			RS1.MoveNext
		loop
		RS1.Close
	next
	'response.write b_code &" bq->bs | bq del | "
	
	'BOM Sort�� �켱���� ������� �����ͼ� BOM_QTY�� �ִ´�. 0���� ū�͸�
	'SQL = "select * from tbBOM_QTY_Sort where IDX='"&strIDX&"' and reorder > 0 order by reorder asc"
	'RS1.Open SQL, sys_DBCon
	'do until RS1.Eof
		'BOM_SORT������ BOM_QTY�� ���
		'SQL = "insert into tbBOM_QTY (BOM_B_Code, BOM_Sub_BS_D_No,BOM_Sub_BS_Code,Parts_P_P_No,Parts_P_P_No2,Parts_P_P_No2_PinYN,BQ_Qty,BQ_Use_YN,BQ_Order,BQ_Remark,BQ_CheckSum,BQ_P_Desc,BQ_P_Spec,BQ_P_Maker,BQ_ready2del_YN) values ("
		'SQL = SQL & freplace(RS1("BOM_B_Code"),"'","''")&","
		'SQL = SQL & "'"&freplace(RS1("BOM_Sub_BS_D_No"),"'","''")&"',"
		'SQL = SQL & freplace(RS1("BOM_Sub_BS_Code"),"'","''")&","
		'SQL = SQL & "'"&freplace(RS1("Parts_P_P_No"),"'","''")&"',"
		'SQL = SQL & "'"&freplace(RS1("Parts_P_P_No2"),"'","''")&"',"
		'SQL = SQL & "'"&freplace(RS1("Parts_P_P_No2_PinYN"),"'","''")&"',"
		'SQL = SQL & freplace(RS1("BQ_Qty"),"'","''")&","
		'SQL = SQL & "'"&freplace(RS1("BQ_Use_YN"),"'","''")&"',"
		'SQL = SQL & "'"&freplace(RS1("BQ_Order"),"'","''")&"',"
		'SQL = SQL & "'"&freplace(RS1("BQ_Remark"),"'","''")&"',"
		'SQL = SQL & "'"&freplace(RS1("BQ_CheckSum"),"'","''")&"',"
		'SQL = SQL & "'"&freplace(RS1("BQ_P_Desc"),"'","''")&"',"
		'SQL = SQL & "'"&freplace(RS1("BQ_P_Spec"),"'","''")&"',"
		'SQL = SQL & "'"&freplace(RS1("BQ_P_Maker"),"'","''")&"',"
		'SQL = SQL & "'"&freplace(RS1("BQ_ready2del_YN"),"'","''")&"')"
		'sys_DBCon.execute(SQL)
		'RS1.MoveNext
	'loop
	'RS1.Close
	'response.write " bs->bq | "
	
	'BOM_Sort�� �����
	'SQL = "delete from tbBOM_QTY_Sort where IDX='"&strIDX&"'"
	'sys_DBCon.execute(SQL)
	'response.write " bs del"
	
	'response.write "<Br>"
	
	set RS2 = nothing
	set RS1 = nothing
end sub
%>

<%
function fReplace(strSource,strChr1,strChr2)
	dim strResult
	
	if isnull(strSource) then
		strResult = ""
	else
		strResult = replace(strSource,strChr1,strChr2)
	end if
	fReplace = strResult

end function
%>

<%
function getRecordBOM_QTY(b_code,remark,p_p_no)
	dim RS1
	dim SQL
	dim strBQ
	
	'�ش� ������ bom_qty���� �ҷ��´�.
	set RS1 = server.CreateObject("ADODB.RecordSet")
	SQL = "select * from tbBOM_Qty where "
	SQL = SQL & "bom_b_code = " & b_code & " and "
	SQL = SQL & "parts_p_p_no = '" & p_p_no & "' and "
	SQL = SQL & "bq_remark = '" & remark & "'"
	RS1.Open SQL,sys_DBCon
	strBQ = ""
	do until RS1.Eof
		strBQ = strBQ & reorder&"/_/"
		strBQ = strBQ & RS1("BOM_B_Code")&"/_/"
		strBQ = strBQ & RS1("BOM_Sub_BS_D_No")&"/_/"
		strBQ = strBQ & RS1("BQ_Code")&"/_/"
		strBQ = strBQ & RS1("BOM_Sub_BS_Code")&"/_/"
		strBQ = strBQ & RS1("Parts_P_P_No")&"/_/"
		strBQ = strBQ & RS1("Parts_P_P_No2")&"/_/"
		strBQ = strBQ & RS1("Parts_P_P_No2_PinYN")&"/_/"
		strBQ = strBQ & RS1("BQ_Qty")&"/_/"
		strBQ = strBQ & RS1("BQ_Use_YN")&"/_/"
		strBQ = strBQ & RS1("BQ_Order")&"/_/"
		strBQ = strBQ & RS1("BQ_Remark")&"/_/"
		strBQ = strBQ & RS1("BQ_CheckSum")&"/_/"
		strBQ = strBQ & RS1("BQ_P_Desc")&"/_/"
		strBQ = strBQ & RS1("BQ_P_Spec")&"/_/"
		strBQ = strBQ & RS1("BQ_P_Maker")&"/_/"
		strBQ = strBQ & RS1("BQ_ready2del_YN")&"/|/"
		RS1.MoveNext
	loop
	RS1.Close
	set RS1 = nothing
	
	getRecordBOM_QTY = strBQ
end function
%>

<%
function getTable()
dim strTable
strTable = ""
strTable=strTable&"3961/_/0CQ22418679/_//_/Capacitor,Film,Box/_/C01J,C02J/_/1/_/DH/|/"
strTable=strTable&"3961/_/0CF2242867A/_/R/_/Capacitor,Film,Box/_/C01J,C02J/_/2/_/DH/|/"
strTable=strTable&"3961/_/0CQ22418678/_//_/Capacitor,Film,Box/_/C01J,C02J/_/3/_/DH/|/"
strTable=strTable&"3961/_/EAH60706601/_//_/Diode,Bridge/_/BD01J/_/1/_/DH/|/"
strTable=strTable&"3961/_/0DD360000AA/_//_/Diode,Bridge/_/BD01J/_/2/_/DH/|/"
strTable=strTable&"3961/_/EAH60664601/_/R/_/Diode,Bridge/_/BD01J/_/3/_/DH/|/"
strTable=strTable&"3961/_/0DD414809AA/_/R/_/Diode,Switching/_/D01B,D02B,D03B,D04B,D01W/_/2/_/DH/|/"
strTable=strTable&"3961/_/0DSSB00029A/_//_/Diode,Switching/_/D01B,D02B,D03B,D04B,D01W/_/1/_/DH/|/"
strTable=strTable&"3961/_/0DSVH00024A/_/R/_/Diode,Switching/_/D01B,D02B,D03B,D04B,D01W/_/3/_/DH/|/"
strTable=strTable&"3961/_/EAF60673101/_//_/Fuse,Time Delay/_/FUSE/_/1/_/DH/|/"
strTable=strTable&"3961/_/0FZZA90005C/_/R/_/Fuse,Time Delay/_/FUSE/_/2/_/DH/|/"
strTable=strTable&"3961/_/6900AQ9028D/_//_/Fuse,Time Delay/_/FUSE/_/3/_/DH/|/"


getTable=strTable
end function
%>
<!-- #include virtual = "/header/db_tail.asp" -->
