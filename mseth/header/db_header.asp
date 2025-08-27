<%
dim sys_DBcon
dim sys_DBconString

sys_DBConString = "Driver={SQL Server}; Server=localhost,1011; Database=Product_Monitor_MSETH; Uid=sa; Pwd=mse7750?!;"

set sys_DBcon = server.CreateObject("adodb.connection")
sys_DBcon.ConnectionTimeout	= 300
sys_DBcon.CommandTimeout 	= 300

sys_DBcon.Open sys_DBConString
%>