<%@ Language=VBScript %>
<%
option explicit
Server.ScriptTimeOut = 300 '5��
Response.Buffer = true

dim gM_ID
gM_ID = lcase(request.Cookies("ADMIN")("M_ID"))

dim gHOST
if instr(Request.ServerVariables("HTTP_HOST"),"th.msekorea.com") > 0 then
	gHOST = "th.msekorea.com"
elseif instr(Request.ServerVariables("HTTP_HOST"),"dev.msekorea.com") > 0 then
	gHOST = "dev.msekorea.com"
elseif instr(Request.ServerVariables("HTTP_HOST"),"dt2070.iptime.org") > 0 then
	gHOST = "dev.msekorea.com"
else
	gHOST = "kr.msekorea.com"
end if

'if left(Request.ServerVariables("REMOTE_ADDR"),12) <> "192.168.123." then
	'Response.Redirect("http://www.naver.com")
'end if
const DefaultPath_workguide_img		= "d:\my_website\msekorea\admin\workguide\workguide_img\"
const DefaultPath_BOM_XLS_Reader	= "d:\my_website\msekorea\admin\upload\bom_temp\"
const DefaultPath_BOM				= "d:\my_website\msekorea\admin\upload\bom_upload\"
const DefaultPath_BOM_UPH			= "d:\my_website\msekorea\admin\upload\bom_uph\"
const DefaultPath_BOM_Update		= "d:\my_website\msekorea\admin\upload\bom_upload\"
const DefaultPath_SCS_XLS_Reader	= "d:\my_website\msekorea\admin\lge_plan\scs_upload\"
const DefaultPath_SAGUP_XLS_Reader	= "d:\my_website\msekorea\admin\parts_incoming_lge\scs_upload\"
const DefaultPath_IPGO_XLS_Reader	= "d:\my_website\msekorea\admin\product_delivery\scs_upload\"
const DefaultPath_Notice			= "d:\my_website\msekorea\admin\notice\upload\"
const DefaultPath_Error_Reporting	= "d:\my_website\msekorea\admin\upload\error_reporting\"
const DefaultPath_SCS_Upload_Update	= "d:\my_website\msekorea\admin\scs_upload\upload\"
const DefaultURL					= "http://www.msekorea.com"
const DefaultTitle					= "MSE Admin"
const DefaultURLAdmin				= "http://admin.msekorea.com"
const DefaultBusinessNo				= "609-81-53434"

const BasicDataPart					= "slt>��ȹ:��ȹ;����:����;������:������;ǰ��:ǰ��;����:����;����:����;����:����;����1:����1;����2:����2;�ѹ�:�ѹ�;�濵��:�濵��"
const BasicDataPosition				= "slt>���:���;����:����;����:����;����:����;����:����;�븮:�븮;����:����;����:����;����:����;�̻�:�̻�;��:��;����:����"

const BasicDataFAWorkType			= "slt>�۾�:�۾�;���۾�:���۾�;�𵨺���:�𵨺���;������:������;�۾����:�۾����;�������:�������;��������:��������;��ȹ����:��ȹ����;�系���:�系���;�����޽�:�����޽�;��Ÿ:��Ÿ;���ְ���:���ְ���;����:����"
const BasicDataMANWorkType			= "slt>�۾�:�۾�;���۾�:���۾�;�𵨺���:�𵨺���;������:������;�۾����:�۾����;�������:�������;��������:��������;��ȹ����:��ȹ����;�系���:�系���;�����޽�:�����޽�;��Ÿ:��Ÿ;���ְ���:���ְ���;����:����"
const BasicDataDLVWorkType			= "slt>�۾�:�۾�"
const BasicDataPartsTransactionType	= "slt>�������:�������;��ǰ����:��ǰ����;����԰�:����԰�"

const BasicDataYN					= "slt>Y:Y;N:N"
const BasicDataPartsType			= "slt>:Ȯ����;IMD:IMD;SMD:SMD;MAN:MAN;ASM:ASM;CBX:CBX;BOX:BOX"
const BasicDataProcess				= "slt>:����;IMD:IMD;SMD:SMD;MAN:MAN;ASM:ASM;CBX:CBX;DLV:DLV"
const BasicDataMaterialProcess		= "slt>:-----;IMD:IMD;SMD:SMD;MAN:MAN;ASM:ASM;DLV:DLV" 
const BasicDataIMDLine				= "slt>UIA1:UIA1;UIA2:UIA2;AVK2:AVK2;AVK2B:AVK2B;UIR1:UIR1;UIR2:UIR2;RH-SGU:RH-SGU;RH:RH;RHU:RHU;RHU2:RHU2;RG131_1:RG131_1;RG131_2:RG131_2"
const BasicDataSMDLine				= "slt>S1:S1;S2:S2;S3:S3;NPM:NPM;NPM2:NPM2"
const BasicDataMANLine				= "slt>P1:P1;P2:P2;P3:P3;P4:P4;P5:P5;P6:P6;N1:N1;N2:N2;NS1:NS1;NS2:NS2;A1:A1"
const BasicDataASMLine				= "slt>C1:C1;C2:C2;C3:C3;C4:C4;C5:C5;C6:C6"
const BasicDataDLVLine				= "slt>LGE:LGE;������:������;���α���:���α���;���ȱ���:���ȱ���;�ݼ�:�ݼ�;�뿵����:�뿵����;�ż���Ÿ:�ż���Ÿ;�츮��:�츮��;SVC:SVC;��ǰû��:��ǰû��;����:����"

const BasicDataLine					= "slt>P1:P1;P2:P2;P3:P3"
'const BasicDataHalfTime			= "slt>T1:T1;T2:T2;T3:T3;T4:T4;T5:T5"
'const BasicDataHalfTimeStr			= "slt>T1:08|20<br>-<br>10|20;T2:10|30<br>-<br>12|30;T3:13|10<br>-<br>15|10;T4:15|20<br>-<br>17|20;T5:17|40<br>-<br>20|40"
const BasicDataHalfTime				= "slt>T1:T1;T2:T2;T3:T3;T4:T4;T5:T5;N1:N1;N2:N2;N3:N3;N4:N4;N5:N5"
const BasicDataHalfTimeStr			= "slt>T1:08|20<br>-<br>10|20;T2:10|30<br>-<br>12|30;T3:13|10<br>-<br>15|10;T4:15|20<br>-<br>17|20;T5:17|40<br>-<br>20|40;N1:20|20<br>-<br>22|20;N2:22|30<br>-<br>24|30;N3:25|10<br>-<br>27|10;N4:27|20<br>-<br>29|20;N5:29|20<br>-<br>32|20"

const BasicDataFullTime				= "slt>T1:T1;T2:T2;T3:T3;T4:T4;T5:T5;N1:N1;N2:N2;N3:N3;N4:N4;N5:N5"
const BasicDataFullTimeStr			= "slt>T1:08|20<br>-<br>10|20;T2:10|30<br>-<br>12|30;T3:13|10<br>-<br>15|10;T4:15|20<br>-<br>17|20;T5:17|40<br>-<br>20|40;N1:20|20<br>-<br>22|20;N2:22|30<br>-<br>24|30;N3:25|10<br>-<br>27|10;N4:27|20<br>-<br>29|20;N5:29|20<br>-<br>32|20"
const BasicModelCompany				= "slt>MSE:MSE;Ÿ��:Ÿ��;�̺з�:�̺з�"

const BasicDataTool					= "slt>�и���:�и���;��ġ��:��ġ��;�ǿܱ�:�ǿܱ�;��ǳ��:��ǳ��;�ߴ���:�ߴ���;â����:â����;ī��Ʈ:ī��Ʈ;����Ƽ��:����Ƽ��;��Ʈ��:��Ʈ��;��Ʈ:��Ʈ;�����͸�:�����͸�"

const BasicDataPartnerType			= "slt>����:����;����:����;������:������;���»�:���»�;��Ÿ:��Ÿ"
const BasicDataPartsIncomingState	= "slt>�����غ�:�����غ�;���ֿϷ�:���ֿϷ�;�԰�Ϸ�:�԰�Ϸ�"
const BasicDataPartsOutgoingState	= "slt>����غ�:����غ�;�Ͱ�Ϸ�:���Ϸ�"
const BasicDataPartsOutgoingComp	= "slt>��������:��������;���������±�:���������±�"
const BasicDataPartnerPaymentType	= "slt>����(1):����(1);����(2):����(2);����(2):����(2);����(3):����(3);����:����"

const BasicDataLGEPlanETCType		= "slt>SVC:SVC;��û:��û;����:����;��ǰ:��ǰ;�ѵ�:�ѵ�;���۾�:���۾�;��Ÿ:��Ÿ;�ֹ����:�ֹ����"
const BasicDataChannel				= "slt>MSE:MSE;��������:��������;�����±�:�����±�"

const BasicDataAuthoriy				= "���߰�����:;�濵������:;���������:;���������:;����������:;���������:;����������:;�����ο�:;ǰ��������:;ǰ���ο�:"
const BasicDataMaterialOrderState	= "slt>�����غ�:�����غ�;���ֿϷ�:���ֿϷ�;�԰�Ϸ�:�԰�Ϸ�;�۾����:�۾����"
const BasicDataMaterialTransactionState	= "slt>�԰�:�԰�;���:���"

const BasicDataMaterialDivision		= "slt>PCBA:PCBA;C/BOX:C/BOX;Remocon:Remocon;LED:LED"
const BasicDataMaterialOSP			= "slt>���:���;������:������"

const BasicDataMaterialTransactionCompany = "slt>��������:��������;��������:��������;��������ũ:��������ũ;��������:��������"
const BasicDataMaterialStockHistoryType	= "slt>test:test"

const BasicDataCostReportUse1		= "slt>��Ÿ��������:��Ÿ��������;ġ������:ġ������;�������:�������;��Ÿ:��Ÿ;����-FCT:����-FCT;����-ICT:����-ICT;����-JIG����:����-JIG����;����-OTP:����-OTP;����-����:����-����;����-��������:����-��������"

const admin_material_handler		= "-no7008-shindk-leehg-ohkh-leejw-shindh-"				'����/�ŷ�ó/�ܰ� �űԵ��

const admin_n_list					= "-shindk-"				'��������

const admin_b_model_reg_form		= "-shindk-moonhj-leejw-kimdh-parksj-no7008-rnd-"	'BOM����ȭ��
const admin_lp_view					= "-shindk-kimjb-no7008-"			'BOM����ȭ��
const admin_b_price_list			= "-shindk-kimdh-no7008-"	'��ǰ�ǰ�
const admin_b_list					= "-shindk-moonhj-leejw-kimdh-no7008-shindh-rnd-"	'BOM����Ʈ
const admin_bu_list					= "-shindk-moonhj-leejw-kimdh-no7008-shindh-rnd-woojm-leehg-"	'�ù�

const admin_ps_list					= "-shindk-no7008-"			'�������
const admin_pd_list					= "-shindk-no7008-"			'LG��ǰ�԰���ȸ
const admin_pil_list				= "-shindk-no7008-"				'LG��������ȸ

const admin_p_list					= "-shindk-no7008-"				'�������(��)
const admin_p_data_list				= "-shindk-no7008-"				'�������(�ܰ�)
const admin_P_Qty_list				= "-shindk-no7008-"				'�������(��ȹ)1
const admin_p_plan_qty_list			= "-shindk-no7008-"				'�������(��ȹ)2 use

const admin_b_qty_list				= "-shindk-no7008-"				'�𵨺��ҿ䷮
const admin_pi_list					= "-shindk-no7008-"			'����ó��
const admin_po_list					= "-shindk-no7008-"			'���ó��

const admin_partner_p_list			= "-shindk-no7008-"			'�ŷ�ó����
const admin_lpe_list				= "-shindk-no7008-"			'��Ÿ��ȹ����
const admin_lm_list					= "-shindk-leehh-no7008-"		'������
const admin_ti_list					= "-shindk-no7008-"			'�������ڷ�
const admin_bom_price_viewer		= "-shindk-kimdh-leejw-no7008-rnd-shindh-"

const ifrm_cr_chart_1				= "-kimys-no7008-"			'����������� ����'
const ifrm_cr_chart_2				= "-sungmd-no7008-"			'����������� �̻�'
const ifrm_cr_chart_3				= "-leejw-leehh-no7008-rnd-"	'����������� ����'

const BasicDataRestStart			= "620-750-910-1040-1360-1470-1630-1760-1900"
const BasicDataRestDiff				= "10-40-10-20-10-40-10-20-40"

const Time_To_Point_Y1				= 0.005	'���� ��, vwPR_List_For_Report�� ������Ʈ �ʿ�
const Time_To_Point_Y2				= 0.005
const Time_To_Point_Y3				= 0.005
const Time_To_Point_F1				= 0.005
const Time_To_Point_F2				= 0.005
const Time_To_Point_RH_U			= 0.005
const Time_To_Point_RHSG			= 0.005
const Time_To_Point_RH_5			= 0.005
const Time_To_Point_RHAV			= 0.005

const Worker_CNT_MAN_P1				= "12"
const Worker_CNT_MAN_P2				= "12"
const Worker_CNT_MAN_P3				= "12"
const Worker_CNT_MAN_P4				= "12"
const Worker_CNT_MAN_P5				= "12"
const Worker_CNT_MAN_P6				= "12"
const Worker_CNT_MAN_N1				= "12"
const Worker_CNT_MAN_N2				= "12"
const Worker_CNT_ASM_C1				= "17"
const Worker_CNT_ASM_C2				= "17"
const Worker_CNT_ASM_C3				= "17"
const Worker_CNT_ASM_C4				= "17"
const Worker_CNT_ASM_C5				= "17"
%>