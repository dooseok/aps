<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<%
dim FileName

dim strFileName
dim arrFileName

strFileName			= Request("strFileName")

if instr(strFileName,"/") > 0 then
	arrFileName = split(strFileName,"/")
	FileName = arrFileName(ubound(arrFileName))
else
	FileName = strFileName
end if
FileName = replace(FileName,".asp","")
FileName = right(replace(date(),"-",""),6) & "_" & FileName

dim RS1
dim CNT1

dim SQL
dim strSelectName
dim arrSelectName
dim strSelect
dim arrSelect
dim strTable
dim strWhere
dim strOrderBy

dim ExistTaxCodeYN

dim arrTaxCode(17,1)

arrTaxCode(0,0)="BUZZER"
arrTaxCode(0,1)="8531800000"
arrTaxCode(1,0)="BUZZER,PIEZO"
arrTaxCode(1,1)="8531809000"
arrTaxCode(2,0)="IPMModule"
arrTaxCode(2,1)="8504909000"
arrTaxCode(3,0)="LCD,MODULE-STN"
arrTaxCode(3,1)="9013801190"
arrTaxCode(4,0)="PHOTO,COUPLER"
arrTaxCode(4,1)="8504909000"
arrTaxCode(5,0)="RECEIVERMODULE"
arrTaxCode(5,1)="8548909000"
arrTaxCode(6,0)="RELAY"
arrTaxCode(6,1)="8536410000"
arrTaxCode(7,0)="RELAY,CONTACT"
arrTaxCode(7,1)="8536410000"
arrTaxCode(8,0)="RELAY,CONTACT"
arrTaxCode(8,1)="8536490000"
arrTaxCode(9,0)="RELAY,NONCONTACT"
arrTaxCode(9,1)="8536410000"
arrTaxCode(10,0)="RESIN"
arrTaxCode(10,1)="3903300000"
arrTaxCode(11,0)="RESIN,ABS"
arrTaxCode(11,1)="3903300000"
arrTaxCode(12,0)="RESIN,HIPS"
arrTaxCode(12,1)="3903190000"
arrTaxCode(13,0)="RESONATOR"
arrTaxCode(13,1)="8543909090"
arrTaxCode(14,0)="RESONATOR,CERAMIC"
arrTaxCode(14,1)="8543909090"
arrTaxCode(15,0)="RESONATOR,CERMIC"
arrTaxCode(15,1)="8543909090"
arrTaxCode(16,0)="TRANSFORMER,POWER"
arrTaxCode(16,1)="8504311000"
arrTaxCode(17,0)="TRANSFORMER"
arrTaxCode(17,1)="8504311000"

SQL					= Request("SQL")
strSelectName		= Request("strSelectName")
strSelect			= Request("strSelect")
strTable			= Request("strTable")
strWhere			= Request("strWhere")
strOrderBy			= Request("strOrderBy")

arrSelectName		= split(strSelectName,",")
arrSelect			= split(strSelect,",")

SQL = replace(SQL,"from",",strP_Desc=(select P_Desc from tbParts where P_P_No=Parts_P_P_No) from")

set RS1 = Server.CreateObject("ADODB.RecordSet") 
RS1.Open SQL,sys_DBCon
if RS1.Eof or RS1.Bof then
%>
<script language="javascript">
alert("조회결과가 없습니다.")
window.close();
</script>
<%
else

	Response.Buffer = false
	Response.Expires = 0
	Response.ContentType = "application/vnd.ms-excel"
	Response.CacheControl = "public"
	Response.AddHeader "Content-Disposition","attachment;filename="&FileName&".xls"

	response.write "제품코드"
	response.write vbtab
	response.write "자재코드"
	response.write vbtab
	response.write "자재-세번부호"
	response.write vbtab
	response.write "자재-품명"
	response.write vbtab
	response.write "자재-규격"
	response.write vbtab
	response.write "자재-소요식"
	response.write vbtab
	response.write "자재-단위소요량"
	response.write vbtab
	response.write "자재-물량단위"	
	response.write vbtab
	response.write "조사란구분"
	response.write vbtab
	response.write "자재-규격2"	
	response.write vbcrlf	
	do until RS1.Eof
		response.write RS1("BOM_Sub_BS_D_No")
		response.write vbtab
		response.write RS1("Parts_P_P_No")
		response.write vbtab
		ExistTaxCodeYN = "N"
		for CNT1 = 0 to ubound(arrTaxCode)
			if ucase(RS1("strP_Desc")) = arrTaxCode(CNT1,0) then
				response.write arrTaxCode(CNT1,1)
				ExistTaxCodeYN = "Y"
			end if
		next
		if ExistTaxCodeYN = "N" then
			response.write ""
		end if
		response.write vbtab
		response.write RS1("strP_Desc")
		response.write vbtab
		response.write RS1("Parts_P_P_No")
		response.write vbtab
		response.write RS1("BQ_Qty")
		response.write vbtab
		response.write RS1("BQ_Qty")
		response.write vbtab
		response.write "EA"
		response.write vbtab
		response.write ""
		response.write vbtab
		response.write ""
		response.write vbcrlf	
		RS1.MoveNext
	loop
end if
RS1.close
set RS1 = nothing
%>
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->