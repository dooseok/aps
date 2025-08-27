<%
dim sys_DBcon
dim sys_DBconString

sys_DBConString = "Driver={SQL Server}; Server=localhost,1011; Database=MSEKOREA; Uid=sa; Pwd=mse7750?!;"

set sys_DBcon = server.CreateObject("adodb.connection")
sys_DBcon.ConnectionTimeout	= 300
sys_DBcon.CommandTimeout 	= 300

'on error resume next

'do
   sys_DBcon.Open sys_DBConString
'loop until err.number = 0

'On Error Goto 0
%>