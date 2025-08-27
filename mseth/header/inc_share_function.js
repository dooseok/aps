function Show_Progress()
{
	document.all.Progress.style.display="";
	document.all.Contents.style.display="none";
}

function InStr(strSearch, charSearchFor)
{
    for (i=0; i < strSearch.length; i++)
    {
		if (charSearchFor == strSearch.substr(i,1))
		{
			return i;
		}
    }
    return -1;
}

function Left(str, n)
{
        if (n <= 0)
                return "";
        else if (n > String(str).length)
                return str;
        else
                return String(str).substring(0,n);
}

function Right(str, n)
{
    if (n <= 0)
         return "";
    else if (n > String(str).length)
         return str;
    else {
         var iLen = String(str).length;
         return String(str).substring(iLen, iLen - n);
    }
}


function List_Validater(strColumn,strName,strType)
{
	var arrColumn	= strColumn.split(",");
	var arrName		= strName.split(",");
	var arrType		= strType.split(",");
	var strTemp		= "";
	var strError	= "";
	
	
	if(frmCommonList.strID_All.length)
	{
		for (cnt1 = 0; cnt1 < frmCommonList.strID_All.length; cnt1++)
		{
			if(arrColumn.length)
			{
				for(var cnt2 = 0; cnt2 < arrColumn.length; cnt2++)
				{
					strTemp = eval("frmCommonList." + arrColumn[cnt2] + "[" + cnt1 + "]");
					if(!strTemp.value)
					{
						strError += "* " + parseInt(cnt1+1) + "  °  ׸  [" + arrName[cnt2] + "]    ʼ  ׸   Դϴ .\n";
					}
					else if(arrType[cnt2] == "num" && !IsNum(strTemp.value))
					{
						strError += "* " + parseInt(cnt1+1) + "  °  ׸  [" + arrName[cnt2] + "]      ڸ   Է°    մϴ .\n";
					}
					else if(arrType[cnt2].substring(0,3) == "fit")
					{
						var nLenLimit = 0;
						arrType[cnt2] = arrType[cnt2].replace('fit','')
						nLenLimit = parseInt(arrType[cnt2]);
						if (strTemp.value.length != nLenLimit)
						{
							strError += "*  ׸  [" + arrName[cnt2] + "]   "+ nLenLimit +" ڸ     Է°    մϴ .\n";
						}
					}
				}
			}
			else
			{
				strTemp = eval("frmCommonList." + strColumn + "[" + cnt1 + "]");
				if(!strTemp.value)
				{
					strError += "* " + parseInt(cnt1+1) + "  °  ׸  [" + strName + "]    ʼ  ׸   Դϴ .\n";
				}
				else if(strType == "num" && !IsNum(strTemp.value))
				{
					strError += "* " + parseInt(cnt1+1) + "  °  ׸  [" + strName + "]      ڸ   Է°    մϴ .\n";
				}
				else if(strType.substring(0,3) == "fit")
				{
					var nLenLimit = 0;
					strType[cnt2] = strType[cnt2].replace('fit','')
					nLenLimit = parseInt(strType[cnt2]);
					if (strTemp.value.length != nLenLimit)
					{
						strError += "*  ׸  [" + strName + "]   "+ nLenLimit +" ڸ     Է°    մϴ .\n";
					}
				}
			}
		}
	}
	else
	{
		
		if(arrColumn.length)
		{
			for(var cnt2 = 0; cnt2 < arrColumn.length; cnt2++)
			{
				
				strTemp = eval("frmCommonList." + arrColumn[cnt2]);
				if(!strTemp.value)
				{
					strError += "*  ׸  [" + arrName[cnt2] + "]    ʼ  ׸   Դϴ .\n";
				}
				else if(arrType[cnt2] == "num" && !IsNum(strTemp.value))
				{
					strError += "*  ׸  [" + arrName[cnt2] + "]      ڸ   Է°    մϴ .\n";
				}
				else if(arrType[cnt2].substring(0,3) == "fit")
				{
					var nLenLimit = 0;
					arrType[cnt2] = arrType[cnt2].replace('fit','')
					nLenLimit = parseInt(arrType[cnt2]);
					if (strTemp.value.length != nLenLimit)
					{
						strError += "*  ׸  [" + arrName[cnt2] + "]   "+ nLenLimit +" ڸ     Է°    մϴ .\n";
					}
				}
			}
		}
		else
		{
			strTemp = eval("frmCommonList." + strColumn + "[" + cnt1 + "]");
			if(!strTemp.value)
			{
				strError += "*  ׸  [" + strName + "]    ʼ  ׸   Դϴ .\n";
			}
			else if(strType == "num" && !IsNum(strTemp.value))
			{
				strError += "*  ׸  [" + strName + "]      ڸ   Է°    մϴ .\n";
			}
			else if(strType.substring(0,3) == "fit")
			{
				var nLenLimit = 0;
				strType[cnt2] = strType[cnt2].replace('fit','')
				nLenLimit = parseInt(strType[cnt2]);
				if (strTemp.value.length != nLenLimit)
				{
					strError += "*  ׸  [" + strName + "]   "+ nLenLimit +" ڸ     Է°    մϴ .\n";
				}
			}
		}
	}
	return strError;
}

function List_Reg_Validater(strColumn,strName,strType)
{
	var arrColumn	= strColumn.split(",");
	var arrName		= strName.split(",");
	var arrType		= strType.split(",");
	var strTemp		= "";
	var strError	= "";
	
	if(arrColumn.length)
	{
		for(var cnt2 = 0; cnt2 < arrColumn.length; cnt2++)
		{
			strTemp = eval("frmCommonListReg." + arrColumn[cnt2]);
			if(!strTemp.value)
			{
				strError += "*  ׸  [" + arrName[cnt2] + "]    ʼ  ׸   Դϴ .\n";
			}
			else if(arrType[cnt2] == "num" && !IsNum(strTemp.value))
			{
				strError += "*  ׸  [" + arrName[cnt2] + "]      ڸ   Է°    մϴ .\n";
			}
			else if(arrType[cnt2].substring(0,3) == "fit")
			{
				var nLenLimit = 0;
				arrType[cnt2] = arrType[cnt2].replace('fit','')
				nLenLimit = parseInt(arrType[cnt2]);
				if (strTemp.value.length != nLenLimit)
				{
					strError += "*  ׸  [" + arrName[cnt2] + "]   "+ nLenLimit +" ڸ     Է°    մϴ .\n";
				}
			}
		}
	}
	else
	{
		strTemp = eval("frmCommonListReg." + strColumn + "[" + cnt1 + "]");
		if(!strTemp.value)
		{
			strError += "*  ׸  [" + strName + "]    ʼ  ׸   Դϴ .\n";
		}
		else if(strType == "num" && !IsNum(strTemp.value))
		{
			strError += "*  ׸  [" + strName + "]      ڸ   Է°    մϴ .\n";
		}
		else if(strType.substring(0,3) == "fit")
		{
			var nLenLimit = 0;
			strType = strType.replace('fit','')
			nLenLimit = parseInt(strType);
			if (strTemp.value.length != nLenLimit)
			{
				strError += "*  ׸  [" + strName + "]   "+ nLenLimit +" ڸ     Է°    մϴ .\n";
			}
		}
	}
	return strError;
}

function List_Reg_Multi_Validater(strColumn,strName,strType,strJudgeColumn)
{
	var arrColumn	= strColumn.split(",");
	var arrName		= strName.split(",");
	var arrType		= strType.split(",");
	var strTemp		= "";
	var strError	= "";
	
	var objColumn = eval("frmCommonListReg." + strColumn);
	
	var flag_yn = 'N';
	for (cnt1 = 0; cnt1 < objColumn.length; cnt1++)
	{
		var objJudgeColumn = eval("frmCommonListReg." + strJudgeColumn + "[" + cnt1 + "]");
		if (objJudgeColumn.value)
			flag_yn = 'Y';
	}
	if (flag_yn=='N')
		return 'CANCEL';

	if(objColumn.length)
	{
		for (cnt1 = 0; cnt1 < objColumn.length; cnt1++)
		{
			var objJudgeColumn = eval("frmCommonListReg." + strJudgeColumn + "[" + cnt1 + "]");
			if (objJudgeColumn.value)
			{
				if(arrColumn.length)
				{
					for(var cnt2 = 0; cnt2 < arrColumn.length; cnt2++)
					{		
						strTemp = eval("frmCommonListReg." + arrColumn[cnt2] + "[" + cnt1 + "]");
						if(!strTemp.value)
						{
							strError += "* " + parseInt(cnt1+1) + "  °  ׸  [" + arrName[cnt2] + "]    ʼ  ׸   Դϴ .\n";
						}
						else if(arrType[cnt2] == "num" && !IsNum(strTemp.value))
						{
							strError += "* " + parseInt(cnt1+1) + "  °  ׸  [" + arrName[cnt2] + "]      ڸ   Է°    մϴ .\n";
						}
						else if(arrType[cnt2].substring(0,3) == "fit")
						{
							var nLenLimit = 0;
							arrType[cnt2] = arrType[cnt2].replace('fit','')
							nLenLimit = parseInt(arrType[cnt2]);
							if (strTemp.value.length != nLenLimit)
							{
								strError += "*  ׸  [" + arrName[cnt2] + "]   "+ nLenLimit +" ڸ     Է°    մϴ .\n";
							}
						}
					}
				}
			}
			else
			{
				var objJudgeColumn = eval("frmCommonListReg." + strJudgeColumn + "[" + cnt1 + "]");
				if (objJudgeColumn.value)
				{
					strTemp = eval("frmCommonListReg." + strColumn + "[" + cnt1 + "]");
					if(!strTemp.value)
					{
						strError += "* " + parseInt(cnt1+1) + "  °  ׸  [" + strName + "]    ʼ  ׸   Դϴ .\n";
					}
					else if(strType == "num" && !IsNum(strTemp.value))
					{
						strError += "* " + parseInt(cnt1+1) + "  °  ׸  [" + strName + "]      ڸ   Է°    մϴ .\n";
					}
					else if(strType.substring(0,3) == "fit")
					{
						var nLenLimit = 0;
						strType[cnt2] = strType[cnt2].replace('fit','')
						nLenLimit = parseInt(strType[cnt2]);
						if (strTemp.value.length != nLenLimit)
						{
							strError += "*  ׸  [" + strName + "]   "+ nLenLimit +" ڸ     Է°    մϴ .\n";
						}
					}
				}
			}
		}
	}
	else
	{
		
		if(arrColumn.length)
		{
			for(var cnt2 = 0; cnt2 < arrColumn.length; cnt2++)
			{
				
				strTemp = eval("frmCommonListReg." + arrColumn[cnt2]);
				if(!strTemp.value)
				{
					strError += "*  ׸  [" + arrName[cnt2] + "]    ʼ  ׸   Դϴ .\n";
				}
				else if(arrType[cnt2] == "num" && !IsNum(strTemp.value))
				{
					strError += "*  ׸  [" + arrName[cnt2] + "]      ڸ   Է°    մϴ .\n";
				}
				else if(arrType[cnt2].substring(0,3) == "fit")
				{
					var nLenLimit = 0;
					arrType[cnt2] = arrType[cnt2].replace('fit','')
					nLenLimit = parseInt(arrType[cnt2]);
					if (strTemp.value.length != nLenLimit)
					{
						strError += "*  ׸  [" + arrName[cnt2] + "]   "+ nLenLimit +" ڸ     Է°    մϴ .\n";
					}
				}
			}
		}
		else
		{
			strTemp = eval("frmCommonListReg." + strColumn + "[" + cnt1 + "]");
			if(!strTemp.value)
			{
				strError += "*  ׸  [" + strName + "]    ʼ  ׸   Դϴ .\n";
			}
			else if(strType == "num" && !IsNum(strTemp.value))
			{
				strError += "*  ׸  [" + strName + "]      ڸ   Է°    մϴ .\n";
			}
			else if(strType.substring(0,3) == "fit")
			{
				var nLenLimit = 0;
				strType[cnt2] = strType[cnt2].replace('fit','')
				nLenLimit = parseInt(strType[cnt2]);
				if (strTemp.value.length != nLenLimit)
				{
					strError += "*  ׸  [" + strName + "]   "+ nLenLimit +" ڸ     Է°    մϴ .\n";
				}
			}
		}
	}
	return strError;
}

function IsNum(strValue)
{
	var chr;
	
	for(var i=0 ; i<strValue.length ; i++)
	{
		chr = strValue.substr(i,1);
		if ((chr < '0' || chr > '9'))
		{
			return false;
		}
	}
	return true;
}

function IsFloat(strValue)
{
	var chr;
	
	if (strValue=="")
		return false;
		
	if (strValue.substr(0,1)==".")
		return false;
	
	if (strValue.substr((strValue.length)-1,1)==".")
		return false;
	
	for(var i=0 ; i<strValue.length ; i++)
	{
		chr = strValue.substr(i,1);
		if ((chr < '0' || chr > '9') && chr != '.' )
		{
			return false;
		}
	}
	return true;
}


function SR_Image_PopUp(SR_Image,SR_CO_ID)
{
	window.open("/service/popup_SR_Image.asp?SR_Image="+SR_Image+"&SR_CO_ID="+SR_CO_ID,"Popup_SR_Image","scrollbars=no,resizable=no,top=100,left=100,width=10,height=10");
}

function SR_Image_Viewer(SR_Code,SR_CO_ID,SetFrame)
{
	window.open("/service/viewer/v_view.asp?SR_Code="+SR_Code+"&SR_CO_ID="+SR_CO_ID+"&SetFrame="+SetFrame,"Popup_SR_Image","scrollbars=no,resizable=no,top=20,left=20,width=970,height=660");
}

function printWindow()
{ 
	factory.printing.header			= "www.HanulJasu.com" 
	factory.printing.footer			= "" 
	factory.printing.portrait		= true
	factory.printing.leftMargin		= 1.0
	factory.printing.topMargin		= 1.0
	factory.printing.rightMargin	= 1.0
	factory.printing.bottomMargin	= 1.0
	factory.printing.Print(false, window)
} 

function trim(Str){
 var tempStr = "";
 
 for (i = 0 ; i < Str.length; i++){
  if(Str.charAt(i) == " "){
   tempStr = tempStr;
  }else{
   tempStr = tempStr + Str.charAt(i);
  }
 }
 
 return tempStr;
}


var target;
var pop_top;
var pop_left;
var cal_Day;
var oPopup = window.createPopup();

function Calendar_Click(e) {
	cal_Day = e.title;
	if (cal_Day.length > 6) {
		if(location.href=="http://admin.spstest.com/mse_plan/inc_mp_view_list.asp")
		{
			if (cal_Day.length == 10)
				target.value = cal_Day.substring(2,10);
			else if (cal_Day.length == 7)
				target.value = '0' + cal_Day;
			else
				target.value = cal_Day;
		}
		else
		{
			target.value = cal_Day;
		}
	}
	oPopup.hide();
}

function Calendar_D(obj) {
	var now = obj.value.split("-");
	target = obj;
	pop_top = document.body.clientTop + GetObjectTop(obj) - document.body.scrollTop;
	pop_left = document.body.clientLeft + GetObjectLeft(obj) -  document.body.scrollLeft;

	if (now.length == 3) {
		Show_cal(now[0],now[1],now[2]);					
	} else {
		now = new Date();
		Show_cal(now.getFullYear(), now.getMonth()+1, now.getDate());
	}
}

function Calendar_M(obj) {
	var now = obj.value.split("-");
	target = obj;
	pop_top = document.body.clientTop + GetObjectTop(obj) - document.body.scrollTop;
	pop_left = document.body.clientLeft + GetObjectLeft(obj) -  document.body.scrollLeft;

	if (now.length == 2) {
		Show_cal_M(now[0],now[1]);					
	} else {
		now = new Date();
		Show_cal_M(now.getFullYear(), now.getMonth()+1);
	}
}

function doOver(el) {
	cal_Day = el.title;

	if (cal_Day.length > 7) {
		el.style.borderColor = "#FF0000";
	}
}

function doOut(el) {
	cal_Day = el.title;

	if (cal_Day.length > 7) {
		el.style.borderColor = "#FFFFFF";
	}
}

function day2(d) {	// 2 ڸ     ڷ      
	var str = new String();
	
	if (parseInt(d) < 10) {
		str = "0" + parseInt(d);
	} else {
		str = "" + parseInt(d);
	}
	return str;
}

function Show_cal(sYear, sMonth, sDay) {
	var Months_day = new Array(0,31,28,31,30,31,30,31,31,30,31,30,31)
	var Month_Val = new Array("01","02","03","04","05","06","07","08","09","10","11","12");
	var intThisYear = new Number(), intThisMonth = new Number(), intThisDay = new Number();

	datToday = new Date();														//               
	
	intThisYear = parseInt(sYear,10);
	intThisMonth = parseInt(sMonth,10);
	intThisDay = parseInt(sDay,10);
	
	if (intThisYear == 0) intThisYear = datToday.getFullYear();					//              
	if (intThisMonth == 0) intThisMonth = parseInt(datToday.getMonth(),10)+1;	//                     -1          ŵ        .
	if (intThisDay == 0) intThisDay = datToday.getDate();
	
	switch(intThisMonth) {
		case 1:
				intPrevYear = intThisYear -1;
				intPrevMonth = 12;
				intNextYear = intThisYear;
				intNextMonth = 2;
				break;
		case 12:
				intPrevYear = intThisYear;
				intPrevMonth = 11;
				intNextYear = intThisYear + 1;
				intNextMonth = 1;
				break;
		default:
				intPrevYear = intThisYear;
				intPrevMonth = parseInt(intThisMonth,10) - 1;
				intNextYear = intThisYear;
				intNextMonth = parseInt(intThisMonth,10) + 1;
				break;
	}
	intPPyear = intThisYear-1
	intNNyear = intThisYear+1

	NowThisYear = datToday.getFullYear();									//        
	NowThisMonth = datToday.getMonth()+1;									//        
	NowThisDay = datToday.getDate();										//        
	
	datFirstDay = new Date(intThisYear, intThisMonth-1, 1);					//           1 Ϸ         ü     (     0     11           (1       12  ))
	intFirstWeekday = datFirstDay.getDay();									//         1                 (0: Ͽ   , 1:      )
	//intSecondWeekday = intFirstWeekday;
	intThirdWeekday = intFirstWeekday;
	
	datThisDay = new Date(intThisYear, intThisMonth, intThisDay);			//  Ѿ                 
	//intThisWeekday = datThisDay.getDay();									//  Ѿ                 
	
	intPrintDay = 1;														//               
	secondPrintDay = 1;
	thirdPrintDay = 1;

	Stop_Flag = 0
	
	if ((intThisYear % 4)==0) {												// 4 ⸶   1   ̸  (  γ              )
		if ((intThisYear % 100) == 0) {
			if ((intThisYear % 400) == 0) {
				Months_day[2] = 29;
			}
		} else {
			Months_day[2] = 29;
		}
	}
	intLastDay = Months_day[intThisMonth];									//                 

	Cal_HTML = "<html><body>";
	Cal_HTML += "<form name='calendar'>";
	Cal_HTML += "<table id=Cal_Table border=0 bgcolor='#f4f4f4' cellpadding=1 cellspacing=1 width=100% onmouseover='parent.doOver(window.event.srcElement)' onmouseout='parent.doOut(window.event.srcElement)' style='font-size : 12;font-family:    ;'>";
	Cal_HTML += "<tr height='35' align=center bgcolor='#f4f4f4'>";
	Cal_HTML += "<td colspan=7 align=center>";
	Cal_HTML += "	<select name='selYear' STYLE='font-size:11;' OnChange='parent.fnChangeYearD(calendar.selYear.value, calendar.selMonth.value, "+intThisDay+")';>";
	for (var optYear=(intThisYear-2); optYear<(intThisYear+2); optYear++) {
		Cal_HTML += "		<option value='"+optYear+"' ";
		if (optYear == intThisYear) Cal_HTML += " selected>\n";
		else Cal_HTML += ">\n";
		Cal_HTML += optYear+"</option>\n";
	}
	Cal_HTML += "	</select>";
	Cal_HTML += "&nbsp;&nbsp;&nbsp;<a style='cursor:hand;' OnClick='parent.Show_cal("+intPrevYear+","+intPrevMonth+","+intThisDay+");'>  </a> ";
	Cal_HTML += "<select name='selMonth' STYLE='font-size:11;' OnChange='parent.fnChangeYearD(calendar.selYear.value, calendar.selMonth.value, "+intThisDay+")';>";
	for (var i=1; i<13; i++) {	
		Cal_HTML += "		<option value='"+Month_Val[i-1]+"' ";
		if (intThisMonth == parseInt(Month_Val[i-1],10)) Cal_HTML += " selected>\n";
		else Cal_HTML += ">\n";
		Cal_HTML += Month_Val[i-1]+"</option>\n";
	}
	Cal_HTML += "	</select>&nbsp;";
	Cal_HTML += "<a style='cursor:hand;' OnClick='parent.Show_cal("+intNextYear+","+intNextMonth+","+intThisDay+");'>  </a>";
	Cal_HTML += "</td></tr>";
	Cal_HTML += "<tr align=center bgcolor='#87B3D6' style='color:#2065DA;' height='25'>";
	Cal_HTML += "	<td style='padding-top:3px;' width='24'><font color=black>  </font></td>";
	Cal_HTML += "	<td style='padding-top:3px;' width='24'><font color=black>  </font></td>";
	Cal_HTML += "	<td style='padding-top:3px;' width='24'><font color=black>ȭ</font></td>";
	Cal_HTML += "	<td style='padding-top:3px;' width='24'><font color=black>  </font></td>";
	Cal_HTML += "	<td style='padding-top:3px;' width='24'><font color=black>  </font></td>";
	Cal_HTML += "	<td style='padding-top:3px;' width='24'><font color=black>  </font></td>";
	Cal_HTML += "	<td style='padding-top:3px;' width='24'><font color=black>  </font></td>";
	Cal_HTML += "</tr>";
		
	for (intLoopWeek=1; intLoopWeek < 7; intLoopWeek++) {	//  ִ             ,  ִ  6  
		Cal_HTML += "<tr height='24' align=right bgcolor='white'>"
		for (intLoopDay=1; intLoopDay <= 7; intLoopDay++) {	//    ϴ             ,  Ͽ        
			if (intThirdWeekday > 0) {											// ù            1     ũ  
				Cal_HTML += "<td>";
				intThirdWeekday--;
			} else {
				if (thirdPrintDay > intLastDay) {								//  Է    ¦          ũ ٸ 
					Cal_HTML += "<td>";
				} else {																//  Է³ ¥            ش   Ǹ 
					Cal_HTML += "<td onClick=parent.Calendar_Click(this); title="+intThisYear+"-"+day2(intThisMonth).toString()+"-"+day2(thirdPrintDay).toString()+" style=\"cursor:Hand;border:1px solid white;";
					if (intThisYear == NowThisYear && intThisMonth==NowThisMonth && thirdPrintDay==intThisDay) {
						Cal_HTML += "background-color:#C6F2ED;";
					}
					
					switch(intLoopDay) {
						case 1:															//  Ͽ    ̸             
							Cal_HTML += "color:red;"
							break;
						//case 7:
						//	Cal_HTML += "color:blue;"
						//	break;
						default:
							Cal_HTML += "color:black;"
							break;
					}
					Cal_HTML += "\">"+thirdPrintDay;
				}
				thirdPrintDay++;
				
				if (thirdPrintDay > intLastDay) {								//        ¥                  ũ          Ż  
					Stop_Flag = 1;
				}
			}
			Cal_HTML += "</td>";
		}
		Cal_HTML += "</tr>";
		if (Stop_Flag==1) break;
	}
	Cal_HTML += "</table></form></body></html>";

	var oPopBody = oPopup.document.body;
	oPopBody.style.backgroundColor = "lightyellow";
	oPopBody.style.border = "solid black 1px";
	oPopBody.innerHTML = Cal_HTML;

	var calHeight = oPopBody.document.all.Cal_Table.offsetHeight;
	//     6         , 5           
	if (intLoopWeek == 6)	calHeight = 214;
	else	calHeight = 189;
	
	oPopup.show(pop_left, (pop_top + target.offsetHeight), 170, calHeight, document.body);
}


function Show_cal_M(sYear, sMonth) {
	var intThisYear = new Number(), intThisMonth = new Number()
	datToday = new Date();													//               
	
	intThisYear = parseInt(sYear,10);
	intThisMonth = parseInt(sMonth,10);
	
	if (intThisYear == 0) intThisYear = datToday.getFullYear();				//              
	if (intThisMonth == 0) intThisMonth = parseInt(datToday.getMonth(),10)+1;	//                     -1          ŵ        .
			
	switch(intThisMonth) {
		case 1:
				intPrevYear = intThisYear -1;
				intNextYear = intThisYear;
				break;
		case 12:
				intPrevYear = intThisYear;
				intNextYear = intThisYear + 1;
				break;
		default:
				intPrevYear = intThisYear;
				intNextYear = intThisYear;
				break;
	}
	intPPyear = intThisYear-1
	intNNyear = intThisYear+1

	Cal_HTML = "<html><head>\n";
	Cal_HTML += "</head><body>\n";
	Cal_HTML += "<table id=Cal_Table border=0 bgcolor='#f4f4f4' cellpadding=1 cellspacing=1 width=100% onmouseover='parent.doOver(window.event.srcElement)' onmouseout='parent.doOut(window.event.srcElement)' style='font-size : 12;font-family:    ;'>\n";
	Cal_HTML += "<tr height='30' align=center bgcolor='#f4f4f4'>\n";
	Cal_HTML += "<td colspan='4' align='center'>\n";
	Cal_HTML += "<a style='cursor:hand;' OnClick='parent.Show_cal_M("+intPPyear+","+intThisMonth+");'>  </a>&nbsp;";
	Cal_HTML += "<select name='selYear' STYLE='font-size:11;' OnChange='parent.fnChangeYearM(this.value, "+intThisMonth+")';>";
	for (var optYear=(intThisYear-2); optYear<(intThisYear+2); optYear++) {
			Cal_HTML += "		<option value='"+optYear+"' ";
			if (optYear == intThisYear) Cal_HTML += " selected>\n";
			else Cal_HTML += ">\n";
			Cal_HTML += optYear+"</option>\n";
	}
	Cal_HTML += "	</select>\n";
	Cal_HTML += "<a style='cursor:hand;' OnClick='parent.Show_cal_M("+intNNyear+","+intThisMonth+");'>  </a>";
	Cal_HTML += "</td></tr>\n";
	Cal_HTML += "<tr><td colspan=4 height='1' bgcolor='#000000'></td></tr>";
	Cal_HTML += "<tr height='20' align=center bgcolor=white>";
	Cal_HTML += "<td onClick=parent.Calendar_Click(this); title="+intThisYear+"-01"+" style=\"cursor:Hand;\">1  </td>";
	Cal_HTML += "<td onClick=parent.Calendar_Click(this); title="+intThisYear+"-02"+" style=\"cursor:Hand;\">2  </td>";
	Cal_HTML += "<td onClick=parent.Calendar_Click(this); title="+intThisYear+"-03"+" style=\"cursor:Hand;\">3  </td>";
	Cal_HTML += "<td onClick=parent.Calendar_Click(this); title="+intThisYear+"-04"+" style=\"cursor:Hand;\">4  </td>";
	Cal_HTML += "</tr>\n";
	Cal_HTML += "<tr height='20' align=center bgcolor=white>";
	Cal_HTML += "<td onClick=parent.Calendar_Click(this); title="+intThisYear+"-05"+" style=\"cursor:Hand;\">5  </td>";
	Cal_HTML += "<td onClick=parent.Calendar_Click(this); title="+intThisYear+"-06"+" style=\"cursor:Hand;\">6  </td>";
	Cal_HTML += "<td onClick=parent.Calendar_Click(this); title="+intThisYear+"-07"+" style=\"cursor:Hand;\">7  </td>";
	Cal_HTML += "<td onClick=parent.Calendar_Click(this); title="+intThisYear+"-08"+" style=\"cursor:Hand;\">8  </td>";
	Cal_HTML += "</tr>\n";
	Cal_HTML += "<tr height='20' align=center bgcolor=white>";
	Cal_HTML += "<td onClick=parent.Calendar_Click(this); title="+intThisYear+"-09"+" style=\"cursor:Hand;\">9  </td>";
	Cal_HTML += "<td onClick=parent.Calendar_Click(this); title="+intThisYear+"-10"+" style=\"cursor:Hand;\">10  </td>";
	Cal_HTML += "<td onClick=parent.Calendar_Click(this); title="+intThisYear+"-11"+" style=\"cursor:Hand;\">11  </td>";
	Cal_HTML += "<td onClick=parent.Calendar_Click(this); title="+intThisYear+"-12"+" style=\"cursor:Hand;\">12  </td>";
	Cal_HTML += "</tr>\n";
	Cal_HTML += "</table>\n</body></html>";

	var oPopBody = oPopup.document.body;
	oPopBody.style.backgroundColor = "lightyellow";
	oPopBody.style.border = "solid black 1px";
	oPopBody.innerHTML = Cal_HTML;

	oPopup.show(pop_left, (pop_top + target.offsetHeight), 160, 99, document.body);
}


//----------------------------------
//	 ϴ޷   ⵵    Ʈ      ⵵     
//----------------------------------
function fnChangeYearD(sYear,sMonth,sDay){
	Show_cal(sYear, sMonth, sDay);
}


//----------------------------------
//	   ޷   ⵵    Ʈ      ⵵     
//----------------------------------
function fnChangeYearM(sYear,sMonth){
	Show_cal_M(sYear, sMonth);
}


/**
	HTML   ü     ƿ  Ƽ  Լ 
**/
function GetObjectTop(obj)
{
	if (obj.offsetParent == document.body)
		return obj.offsetTop;
	else
		return obj.offsetTop + GetObjectTop(obj.offsetParent);
}

function GetObjectLeft(obj)
{
	if (obj.offsetParent == document.body)
		return obj.offsetLeft;
	else
		return obj.offsetLeft + GetObjectLeft(obj.offsetParent);
}


/*  Function Equivalent to java.net.URLEncoder.encode(String, "UTF-8")
     Copyright (C) 2002, Cresc Corp.
     Version: 1.0
 */
 function encodeURL(str){
     var s0, i, s, u;
     s0 = "";                // encoded str
     for (i = 0; i < str.length; i++){   // scan the source
         s = str.charAt(i);
         u = str.charCodeAt(i);          // get unicode of the char
         if (s == " "){s0 += "+";}       // SP should be converted to "+"
         else {
             if ( u == 0x2a || u == 0x2d || u == 0x2e || u == 0x5f || ((u >= 0x30) && (u <= 0x39)) || ((u >= 0x41) && (u <= 0x5a)) || ((u >= 0x61) && (u <= 0x7a))){       // check for escape
                 s0 = s0 + s;            // don't escape
             }
             else {                  // escape
                 if ((u >= 0x0) && (u <= 0x7f)){     // single byte format
                     s = "0"+u.toString(16);
                     s0 += "%"+ s.substr(s.length-2);
                 }
                 else if (u > 0x1fffff){     // quaternary byte format (extended)
                     s0 += "%" + (oxf0 + ((u & 0x1c0000) >> 18)).toString(16);
                     s0 += "%" + (0x80 + ((u & 0x3f000) >> 12)).toString(16);
                     s0 += "%" + (0x80 + ((u & 0xfc0) >> 6)).toString(16);
                     s0 += "%" + (0x80 + (u & 0x3f)).toString(16);
                 }
                 else if (u > 0x7ff){        // triple byte format
                     s0 += "%" + (0xe0 + ((u & 0xf000) >> 12)).toString(16);
                     s0 += "%" + (0x80 + ((u & 0xfc0) >> 6)).toString(16);
                     s0 += "%" + (0x80 + (u & 0x3f)).toString(16);
                 }
                 else {                      // double byte format
                     s0 += "%" + (0xc0 + ((u & 0x7c0) >> 6)).toString(16);
                     s0 += "%" + (0x80 + (u & 0x3f)).toString(16);
                 }
             }
         }
     }
     return s0;
 }
 
 /*  Function Equivalent to java.net.URLDecoder.decode(String, "UTF-8")
     Copyright (C) 2002, Cresc Corp.
     Version: 1.0
 */
 function decodeURL(str){
     var s0, i, j, s, ss, u, n, f;
     s0 = "";                // decoded str
     for (i = 0; i < str.length; i++){   // scan the source str
         s = str.charAt(i);
         if (s == "+"){s0 += " ";}       // "+" should be changed to SP
         else {
             if (s != "%"){s0 += s;}     // add an unescaped char
             else{               // escape sequence decoding
                 u = 0;          // unicode of the character
                 f = 1;          // escape flag, zero means end of this sequence
                 while (true) {
                     ss = "";        // local str to parse as int
                         for (j = 0; j < 2; j++ ) {  // get two maximum hex characters for parse
                             sss = str.charAt(++i);
                             if (((sss >= "0") && (sss <= "9")) || ((sss >= "a") && (sss <= "f"))  || ((sss >= "A") && (sss <= "F"))) {
                                 ss += sss;      // if hex, add the hex character
                             } else {--i; break;}    // not a hex char., exit the loop
                         }
                     n = parseInt(ss, 16);           // parse the hex str as byte
                     if (n <= 0x7f){u = n; f = 1;}   // single byte format
                     if ((n >= 0xc0) && (n <= 0xdf)){u = n & 0x1f; f = 2;}   // double byte format
                     if ((n >= 0xe0) && (n <= 0xef)){u = n & 0x0f; f = 3;}   // triple byte format
                     if ((n >= 0xf0) && (n <= 0xf7)){u = n & 0x07; f = 4;}   // quaternary byte format (extended)
                     if ((n >= 0x80) && (n <= 0xbf)){u = (u << 6) + (n & 0x3f); --f;}         // not a first, shift and add 6 lower bits
                     if (f <= 1){break;}         // end of the utf byte sequence
                     if (str.charAt(i + 1) == "%"){ i++ ;}                   // test for the next shift byte
                     else {break;}                   // abnormal, format error
                 }
             s0 += String.fromCharCode(u);           // add the escaped character
             }
         }
     }
     return s0;
 }