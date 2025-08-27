<%
dim v1
v1 = request("v1")

if isnumeric(left(v1,3)) then
	v1 = left(v1,10)
else
	v1 = left(v1,9)
end if

response.redirect "/index.asp?strURL="&server.URLEncode("/bom/new_bu_list.asp?s_BOM_B_D_No="&v1)
%>
