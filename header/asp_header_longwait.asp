<%@ Language=VBScript %>
<%
option explicit
Server.ScriptTimeOut = 300 '5분
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

const BasicDataPart					= "slt>기획:기획;개발:개발;생산기술:생산기술;품질:품질;구매:구매;자재:자재;영업:영업;제조1:제조1;제조2:제조2;총무:총무;경영진:경영진"
const BasicDataPosition				= "slt>사원:사원;조장:조장;반장:반장;주임:주임;계장:계장;대리:대리;과장:과장;차장:차장;부장:부장;이사:이사;상무:상무;사장:사장"

const BasicDataFAWorkType			= "slt>작업:작업;재작업:재작업;모델변경:모델변경;자재대기:자재대기;작업대기:작업대기;설비고장:설비고장;설비점검:설비점검;계획없음:계획없음;사내행사:사내행사;정규휴식:정규휴식;기타:기타;외주가공:외주가공;재해:재해"
const BasicDataMANWorkType			= "slt>작업:작업;재작업:재작업;모델변경:모델변경;자재대기:자재대기;작업대기:작업대기;설비고장:설비고장;설비점검:설비점검;계획없음:계획없음;사내행사:사내행사;정규휴식:정규휴식;기타:기타;외주가공:외주가공;재해:재해"
const BasicDataDLVWorkType			= "slt>작업:작업"
const BasicDataPartsTransactionType	= "slt>개별등록:개별등록;부품정보:부품정보;사급입고:사급입고"

const BasicDataYN					= "slt>Y:Y;N:N"
const BasicDataPartsType			= "slt>:확인중;IMD:IMD;SMD:SMD;MAN:MAN;ASM:ASM;CBX:CBX;BOX:BOX"
const BasicDataProcess				= "slt>:선택;IMD:IMD;SMD:SMD;MAN:MAN;ASM:ASM;CBX:CBX;DLV:DLV"
const BasicDataMaterialProcess		= "slt>:-----;IMD:IMD;SMD:SMD;MAN:MAN;ASM:ASM;DLV:DLV" 
const BasicDataIMDLine				= "slt>UIA1:UIA1;UIA2:UIA2;AVK2:AVK2;AVK2B:AVK2B;UIR1:UIR1;UIR2:UIR2;RH-SGU:RH-SGU;RH:RH;RHU:RHU;RHU2:RHU2;RG131_1:RG131_1;RG131_2:RG131_2"
const BasicDataSMDLine				= "slt>S1:S1;S2:S2;S3:S3;NPM:NPM;NPM2:NPM2"
const BasicDataMANLine				= "slt>P1:P1;P2:P2;P3:P3;P4:P4;P5:P5;P6:P6;N1:N1;N2:N2;NS1:NS1;NS2:NS2;A1:A1"
const BasicDataASMLine				= "slt>C1:C1;C2:C2;C3:C3;C4:C4;C5:C5;C6:C6"
const BasicDataDLVLine				= "slt>LGE:LGE;디케이:디케이;정민기전:정민기전;성안기전:성안기전;금석:금석;대영전자:대영전자;신성델타:신성델타;우리텍:우리텍;SVC:SVC;물품청구:물품청구;샘플:샘플"

const BasicDataLine					= "slt>P1:P1;P2:P2;P3:P3"
'const BasicDataHalfTime			= "slt>T1:T1;T2:T2;T3:T3;T4:T4;T5:T5"
'const BasicDataHalfTimeStr			= "slt>T1:08|20<br>-<br>10|20;T2:10|30<br>-<br>12|30;T3:13|10<br>-<br>15|10;T4:15|20<br>-<br>17|20;T5:17|40<br>-<br>20|40"
const BasicDataHalfTime				= "slt>T1:T1;T2:T2;T3:T3;T4:T4;T5:T5;N1:N1;N2:N2;N3:N3;N4:N4;N5:N5"
const BasicDataHalfTimeStr			= "slt>T1:08|20<br>-<br>10|20;T2:10|30<br>-<br>12|30;T3:13|10<br>-<br>15|10;T4:15|20<br>-<br>17|20;T5:17|40<br>-<br>20|40;N1:20|20<br>-<br>22|20;N2:22|30<br>-<br>24|30;N3:25|10<br>-<br>27|10;N4:27|20<br>-<br>29|20;N5:29|20<br>-<br>32|20"

const BasicDataFullTime				= "slt>T1:T1;T2:T2;T3:T3;T4:T4;T5:T5;N1:N1;N2:N2;N3:N3;N4:N4;N5:N5"
const BasicDataFullTimeStr			= "slt>T1:08|20<br>-<br>10|20;T2:10|30<br>-<br>12|30;T3:13|10<br>-<br>15|10;T4:15|20<br>-<br>17|20;T5:17|40<br>-<br>20|40;N1:20|20<br>-<br>22|20;N2:22|30<br>-<br>24|30;N3:25|10<br>-<br>27|10;N4:27|20<br>-<br>29|20;N5:29|20<br>-<br>32|20"
const BasicModelCompany				= "slt>MSE:MSE;타사:타사;미분류:미분류"

const BasicDataTool					= "slt>분리형:분리형;상치형:상치형;실외기:실외기;온풍기:온풍기;중대형:중대형;창문형:창문형;카셋트:카셋트;컨브티블:컨브티블;아트쿨:아트쿨;덕트:덕트;유니터리:유니터리"

const BasicDataPartnerType			= "slt>매입:매입;매출:매출;매입출:매입출;협력사:협력사;기타:기타"
const BasicDataPartsIncomingState	= "slt>발주준비:발주준비;발주완료:발주완료;입고완료:입고완료"
const BasicDataPartsOutgoingState	= "slt>출고준비:출고준비;촐고완료:출고완료"
const BasicDataPartsOutgoingComp	= "slt>문성전자:문성전자;문성전자태국:문성전자태국"
const BasicDataPartnerPaymentType	= "slt>현금(1):현금(1);현금(2):현금(2);어음(2):어음(2);어음(3):어음(3);미정:미정"

const BasicDataLGEPlanETCType		= "slt>SVC:SVC;물청:물청;샘플:샘플;초품:초품;한도:한도;재작업:재작업;기타:기타;주문취소:주문취소"
const BasicDataChannel				= "slt>MSE:MSE;문성전자:문성전자;문성태국:문성태국"

const BasicDataAuthoriy				= "개발관리자:;경영관리자:;전산관리자:;서브관리자:;영업관리자:;자재관리자:;제조관리자:;제조인원:;품질관리자:;품질인원:"
const BasicDataMaterialOrderState	= "slt>발주준비:발주준비;발주완료:발주완료;입고완료:입고완료;작업취소:작업취소"
const BasicDataMaterialTransactionState	= "slt>입고:입고;출고:출고"

const BasicDataMaterialDivision		= "slt>PCBA:PCBA;C/BOX:C/BOX;Remocon:Remocon;LED:LED"
const BasicDataMaterialOSP			= "slt>사급:사급;직구입:직구입"

const BasicDataMaterialTransactionCompany = "slt>엠에스이:엠에스이;문성전자:문성전자;에이피테크:에이피테크;에스엠텍:에스엠텍"
const BasicDataMaterialStockHistoryType	= "slt>test:test"

const BasicDataCostReportUse1		= "slt>기타제조관련:기타제조관련;치공구류:치공구류;설비수리:설비수리;기타:기타;개발-FCT:개발-FCT;개발-ICT:개발-ICT;개발-JIG관련:개발-JIG관련;개발-OTP:개발-OTP;개발-관리:개발-관리;자재-차량관리:자재-차량관리"

const admin_material_handler		= "-no7008-shindk-leehg-ohkh-leejw-shindh-"				'자재/거래처/단가 신규등록

const admin_n_list					= "-shindk-"				'공지사항

const admin_b_model_reg_form		= "-shindk-moonhj-leejw-kimdh-parksj-no7008-rnd-"	'BOM수정화면
const admin_lp_view					= "-shindk-kimjb-no7008-"			'BOM수정화면
const admin_b_price_list			= "-shindk-kimdh-no7008-"	'제품판가
const admin_b_list					= "-shindk-moonhj-leejw-kimdh-no7008-shindh-rnd-"	'BOM리스트
const admin_bu_list					= "-shindk-moonhj-leejw-kimdh-no7008-shindh-rnd-woojm-leehg-"	'시방

const admin_ps_list					= "-shindk-no7008-"			'영업출고
const admin_pd_list					= "-shindk-no7008-"			'LG제품입고조회
const admin_pil_list				= "-shindk-no7008-"				'LG사급출고조회

const admin_p_list					= "-shindk-no7008-"				'파츠목록(상세)
const admin_p_data_list				= "-shindk-no7008-"				'파츠목록(단가)
const admin_P_Qty_list				= "-shindk-no7008-"				'파츠목록(계획)1
const admin_p_plan_qty_list			= "-shindk-no7008-"				'파츠목록(계획)2 use

const admin_b_qty_list				= "-shindk-no7008-"				'모델별소요량
const admin_pi_list					= "-shindk-no7008-"			'발주처리
const admin_po_list					= "-shindk-no7008-"			'출고처리

const admin_partner_p_list			= "-shindk-no7008-"			'거래처관리
const admin_lpe_list				= "-shindk-no7008-"			'기타계획관리
const admin_lm_list					= "-shindk-leehh-no7008-"		'모델정보
const admin_ti_list					= "-shindk-no7008-"			'툴기초자료
const admin_bom_price_viewer		= "-shindk-kimdh-leejw-no7008-rnd-shindh-"

const ifrm_cr_chart_1				= "-kimys-no7008-"			'지출관리결제 사장'
const ifrm_cr_chart_2				= "-sungmd-no7008-"			'지출관리결제 이사'
const ifrm_cr_chart_3				= "-leejw-leehh-no7008-rnd-"	'지출관리결제 팀장'

const BasicDataRestStart			= "620-750-910-1040-1360-1470-1630-1760-1900"
const BasicDataRestDiff				= "10-40-10-20-10-40-10-20-40"

const Time_To_Point_Y1				= 0.005	'수정 후, vwPR_List_For_Report도 업데이트 필요
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