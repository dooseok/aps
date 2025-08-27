<!-- #include Virtual = "/header/asp_header.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include Virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_full_header.asp" -->

<%
dim SQL
dim RS1
dim RS2

dim strFormName
dim strObjIdx
dim s_P_P_No
dim s_Partner_P_Name

dim Partner_P_Name
dim PP_Price
dim P_Payment_Type

dim strPartner_P_Name
dim strPP_Price
dim strP_Payment_Type

dim arrPartner_P_Name
dim arrPP_Price
dim arrP_Payment_Type

strFormName			= Request("strFormName")
strObjIdx			= Request("strObjIdx")
s_P_P_No			= Request("s_P_P_No")
s_Partner_P_Name	= Request("s_Partner_P_Name")

set RS1 = Server.CreateObject("ADODB.RecordSet")
set RS2 = Server.CreateObject("ADODB.RecordSet")

if s_P_P_No <> "" and s_Partner_P_Name = "" then
	SQL = "select PP_Price, Partner_P_Name, P_Payment_Type from tbParts_Price t1 left outer join tbPartner t2 on t1.Partner_P_Name=t2.P_Name where Parts_P_P_No='"&s_P_P_No&"' order by PP_Last_YN desc"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
		SQL ="select P_Name from tbPartner order by P_Sort asc, P_Name asc"
		RS2.Open SQL,sys_DBCon
		strPartner_P_Name = Server.URLEncode("-선택-|/|")
		do until RS2.Eof
			Partner_P_Name	= RS2("P_Name")
			if isnull(Partner_P_Name) then
				Partner_P_Name = ""
			end if
			RS2.MoveNext
			Partner_P_Name	= Server.URLEncode(Partner_P_Name)
			strPartner_P_Name	= strPartner_P_Name	& Partner_P_Name 	& "|/|"
		loop
		RS2.Close
%>
<script language="javascript">
//parent.MakeSelectBox("<%=strFormName%>","Partner_P_Name[<%=strObjIdx%>]","<%=strPartner_P_Name%>");
//parent.Reset_Form("<%=strFormName%>",<%=strObjIdx%>);
</script>	
<%
	else
		do until RS1.Eof
			Partner_P_Name	= RS1("Partner_P_Name")
			PP_Price		= RS1("PP_Price")
			P_Payment_Type	= RS1("P_Payment_Type")
			
			if isnull(Partner_P_Name) then
				Partner_P_Name = ""
			end if
			if isnull(PP_Price) then
				PP_Price = ""
			end if
			if isnull(P_Payment_Type) then
				P_Payment_Type = ""
			end if
			
			Partner_P_Name	= Server.URLEncode(Partner_P_Name)
			PP_Price		= Server.URLEncode(PP_Price)
			P_Payment_Type	= Server.URLEncode(P_Payment_Type)
			
			strPartner_P_Name	= strPartner_P_Name	& Partner_P_Name 	& "|/|"
			strPP_Price			= strPP_Price		& PP_Price 			& "|/|"
			strP_Payment_Type	= strP_Payment_Type	& P_Payment_Type 	& "|/|"
			
			RS1.MoveNext
		loop
		
		arrPartner_P_Name	= split(strPartner_P_Name,"|/|")
		arrPP_Price			= split(strPP_Price,"|/|")
		arrP_Payment_Type	= split(strP_Payment_Type,"|/|")
%>
<script language="javascript">
//parent.MakeSelectBox("<%=strFormName%>","Partner_P_Name[<%=strObjIdx%>]","<%=strPartner_P_Name%>");
parent.Fill_Form("<%=strFormName%>","PO_Price[<%=strObjIdx%>]","<%=arrPP_Price(0)%>");
//parent.Fill_Form("<%=strFormName%>","PO_Payment_Type[<%=strObjIdx%>]","<%=arrP_Payment_Type(0)%>");
parent.cal_Sum_Price_Qty("<%=strFormName%>","<%=strObjIdx%>");
</script>
<%
	end if
	RS1.Close
elseif s_P_P_No <> "" and s_Partner_P_Name <> "" then
	SQL = "select PP_Price, Partner_P_Name, P_Payment_Type from tbParts_Price t1 left outer join tbPartner t2 on t1.Partner_P_Name=t2.P_Name where Parts_P_P_No='"&s_P_P_No&"' and Partner_P_Name='"&s_Partner_P_Name&"' order by PP_Last_YN desc"
	RS1.Open SQL,sys_DBCon
	if RS1.Eof or RS1.Bof then
	else
		do until RS1.Eof
			Partner_P_Name	= RS1("Partner_P_Name")
			PP_Price		= RS1("PP_Price")
			P_Payment_Type	= RS1("P_Payment_Type")
			
			if isnull(Partner_P_Name) then
				Partner_P_Name = ""
			end if
			if isnull(PP_Price) then
				PP_Price = ""
			end if
			if isnull(P_Payment_Type) then
				P_Payment_Type = ""
			end if
			
			Partner_P_Name	= Server.URLEncode(Partner_P_Name)
			PP_Price		= Server.URLEncode(PP_Price)
			P_Payment_Type	= Server.URLEncode(P_Payment_Type)
			
			strPartner_P_Name	= strPartner_P_Name	& Partner_P_Name 	& "|/|"
			strPP_Price			= strPP_Price		& PP_Price 			& "|/|"
			strP_Payment_Type	= strP_Payment_Type	& P_Payment_Type 	& "|/|"
			
			RS1.MoveNext
		loop
		
		arrPartner_P_Name	= split(strPartner_P_Name,"|/|")
		arrPP_Price			= split(strPP_Price,"|/|")
		arrP_Payment_Type	= split(strP_Payment_Type,"|/|")
%>
<script language="javascript">
//parent.Fill_Form("<%=strFormName%>","PO_Price[<%=strObjIdx%>]","<%=arrPP_Price(0)%>");
parent.Fill_Form("<%=strFormName%>","PO_Payment_Type[<%=strObjIdx%>]","<%=arrP_Payment_Type(0)%>");
//parent.cal_Sum_Price_Qty("<%=strFormName%>","<%=strObjIdx%>");
</script>
<%
	end if
	RS1.Close
end if

set RS2 = nothing
set RS1 = nothing
%>
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->