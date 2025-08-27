<!-- #include Virtual = "/header/asp_header_longwait.asp" -->
<!-- #include virtual = "/header/session_check_header.asp" -->
<!-- #include virtual = "/header/db_header.asp" -->
<!-- #include virtual = "/header/html_header.asp" -->
<!-- #include virtual = "/header/layout_header.asp" -->
<!-- #include virtual = "/function/inc_share_function.asp" -->

<% 
Dim objRs

Dim strSQL
Dim Strsearch,blnSearch,strsearchSQL

Dim intNowPage, intTotalCount, intTotalPage, intBlockPage, intPageSize
Dim intTemp, intLoop

intNowPage = Request("page")   
Strsearch = Request("searchParts")
blnSearch = false  
intPageSize = 100
intBlockPage = 100



If Len(intNowPage) = 0 Then
    intNowPage = 1
End If

Strsearch = replace(Strsearch,chr(13)&chr(10),"','")

If Len(Strsearch) <> 0 Then
    blnSearch = true
	strsearchSQL = "and " 
	strsearchSQL = strsearchSQL & " Parts_P_P_No in ('"& Strsearch &"') "
End if

strSQL = "Select Count(*)"
strSQL = strSQL & ",CEILING(CAST(Count(*) AS FLOAT) /" & intPageSize & ") "
strSQL = strSQL & " from tbBOM_Qty"
strSQL = strSQL & " where BQ_Qty > 0 "

If blnSearch Then
    strSQL = strSQL & strsearchSQL
End If

Set objRs = Server.CreateObject("ADODB.RecordSet")

if blnSearch = true then
	objRs.Open strSQL, sys_DBcon
	intTotalCount = objRs(0)
	intTotalPage = objRs(1)
	objRs.close
else
	intTotalCount = 0
	intTotalPage = 0
end if

strSQL = "select Top " & intNowPage * intPageSize & ""
strSQL = strSQL & "	BOM_Sub_BS_D_No, "
strSQL = strSQL & "	Parts_P_P_No, "
strSQL = strSQL & "	BQ_Qty = sum(BQ_Qty)"
strSQL = strSQL & " from "
strSQL = strSQL & "	tbBOM_Qty "
strSQL = strSQL & " where BQ_Qty > 0 "

If blnSearch Then
    strSQL = strSQL & strsearchSQL
End If 

strSQL = strSQL & " group by "
strSQL = strSQL & "		BOM_Sub_BS_D_No,"
strSQL = strSQL & "		Parts_P_P_No "
strSQL = strSQL & " order by "
strSQL = strSQL & "		BOM_Sub_BS_D_No,"
strSQL = strSQL & "		Parts_P_P_No "

if blnSearch = true then
	objRs.Open strSQL
end if
%>


<div align="center">
<h2>부품소요량조회</h2>	

<table border=1>
<iframe name="ifrmXLSDown" src="about:blank" frameborder=0 width=0px height=0px></iframe>
<form name="frmList2Excel" action="/function/inc_List2Excel.asp" method="post" target="ifrmXLSDown">
<input type="hidden"	name="SQL"				value="<%=replace(strSQL,"Top " & intNowPage * intPageSize,"")%>">
<input type="hidden"	name="strSelectName"	value="도면파트넘버,부품파트넘버,소요수량">
<input type="hidden"	name="strSelect"		value="BOM_Sub_BS_D_No,Parts_P_P_No,BQ_Qty">
<input type="hidden"	name="strFileName"		value="b_parts_multi_list.asp">
</form>
<form name="frmList2Excel4FTA" action="b_parts_multi_list2excel4fta.asp" method="post" target="_blank">
<input type="hidden"	name="SQL"				value="<%=replace(strSQL,"Top " & intNowPage * intPageSize,"")%>">
<input type="hidden"	name="strSelectName"	value="제품코드,자재코드,소요식">
<input type="hidden"	name="strSelect"		value="BOM_Sub_BS_D_No,Parts_P_P_No,BQ_Qty">
<input type="hidden"	name="strFileName"		value="b_parts_multi_list.asp">
</form>
<form name="searchForm" method="get" action="b_parts_multi_list.asp">
<tr>
	<td>
		<textarea name="searchParts" style="width:300px;height=150px"><%=Request("searchParts")%></textarea>
	</td>
	<td>
		<input type="submit" value="검 색" style="width:70px"><br><br>
		<input type="button" value="값지우기" style="width:70px" onclick="searchParts.value=''"><br><br>
		<input type="button" value="엑셀로받기" style="width:70px" onclick="frmList2Excel.submit();"><Br><Br>
		<input type="button" value="FTA용엑셀" style="width:70px" onclick="frmList2Excel4FTA.submit();">
	</td>
</tr>
</table>
</form>

<% If intTotalCount > 0 Then %>
전체게시 <%=intTotalCount%> 개 &nbsp;&nbsp;&nbsp;&nbsp;
현재페이지 : <%=intNowPage%> / <%=intTotalPage%>
<%  End If  %>
<table border width="600">
<tr align="center">
	<td bgcolor=skyblue width=200px>도면파트넘버</td>
	<td bgcolor=skyblue width=200px>부품파트넘버</td>
	<td bgcolor=skyblue>소요수량</td>
</tr>
<%
If blnSearch = false Then
%>
<tr align="center">

	<td colspan="5">검색어를 입력하세요</td>
</tr>
<%
Else
%>
<%
	If objRs.BOF or objRs.EOF Then 
%>

<tr align="center">

	<td colspan="5">등록된 정보가 없습니다</td>
</tr>
<%
	Else
		objRs.Move (intNowPage - 1) * intPageSize
	
		Do Until objRs.EOF
%>
<tr align="center">
	<td><%=objRs("BOM_Sub_BS_D_No")%></td>
	<td><%=objRs("Parts_P_P_No")%></td>
	<td><%=objRs("BQ_Qty")%></td>
</tr>
<%
			objRs.MoveNext
		Loop
	End If

	objRs.Close
	Set objRs = nothing
%>
</table>


<table>
  <tr>
    <td align="center">
    <%
            intTemp = Int((intNowPage - 1) / intBlockPage) * intBlockPage + 1

            If intTemp = 1 Then
                Response.Write "[이전 " & intBlockPage & "개]"
            Else
                Response.Write"<a href=b_parts_multi_list.asp?page=" & intTemp - intBlockPage &_
                "&searchParts=" & Strsearch &_
                ">[이전 " & intBlockPage & "개]</a>"
            End If

            intLoop = 1

            Do Until intLoop > intBlockPage Or intTemp > intTotalPage
                If intTemp = CInt(intNowPage) Then
                    Response.Write "<font size= 3><b>" & intTemp &"</b></font>&nbsp;"
                Else
                    Response.Write"<a href=b_parts_multi_list.asp?page=" & intTemp &_
                    "&searchParts=" & Strsearch &_
                    ">" & intTemp & "</a>&nbsp;"

                End If
                intTemp = intTemp + 1
                intLoop = intLoop + 1
            Loop

            If intTemp > intTotalPage Then
                Response.Write "[다음 " &intBlockPage&"개]"
            Else
                Response.Write"<a href=b_parts_multi_list.asp?page=" & intTemp &_
                "&searchParts=" & Strsearch &_
                ">[다음 " & intBlockPage & "개]</a>"
            End If
    %>
    </td>
  </tr>
  <%
End if
%>
</table>


</div>
</body>
</html> 






<!-- #include virtual = "/header/layout_tail.asp" -->
<!-- #include virtual = "/header/html_tail.asp" -->
<!-- #include virtual = "/header/db_tail.asp" -->
<!-- #include virtual = "/header/session_check_tail.asp" -->