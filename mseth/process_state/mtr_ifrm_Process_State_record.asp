<!-- #include Virtual = "/mseth/header/asp_header.asp" -->
<!-- #include Virtual = "/mseth/header/db_header.asp" -->
<!-- include Virtual = "/mseth/header/html_header.asp" -->
<!-- #include Virtual = "/mseth/header/layout_full_header.asp" -->
<!-- #include Virtual = "/mseth/header/inc_share_function.asp" -->

<%
dim s_Work_Date
dim s_Line

s_Work_Date = request("s_Work_Date")
s_Line = request("s_Line")
%>
<html>
<head>
</head>
<body topmargin=0 leftmargin=0 bgcolor=black>
<div id="idContent" style="width:100%;height:100%;background-color:black;"></div>
</body>
</html>

<iframe src="mtr_Content_process_state_record.asp?s_Work_Date=<%=s_Work_Date%>&s_Line=<%=s_Line%>" width=0px height=0px frameborder=0></iframe>

<!-- include Virtual = "/mseth/header/html_tail.asp" -->
<!-- #include Virtual = "/mseth/header/db_tail.asp" -->


	