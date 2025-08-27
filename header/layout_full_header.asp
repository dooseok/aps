</head>

<%
'자재 현황판 때문에
if instr(lcase(Request.ServerVariables("HTTP_URL")),"mtr_") > 0 then
%>
<body topmargin=0 leftmargin=0 bgcolor="#ffffff">
<%
elseif instr(lcase(Request.ServerVariables("HTTP_URL")),"workguide") > 0 then
%>
<body topmargin=0 leftmargin=0 bgcolor="#ffffff">
<%
else
%>
<body topmargin=0 leftmargin=0 bgcolor="#ffffff" onload="self.focus();">
<%
end if
%>