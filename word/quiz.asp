<html>
<head> 
	<title>WORD QUIZ</title>
<style>
h1 {font-size: 40pt}
h2 {font-size: 35pt}
h3 {font-size: 30pt}
h4 {font-size: 25pt}

input[type="text"]
{ 
	width: 250px; /* 원하는 너비 설정 */ 
	height: 35px; /* 높이값 초기화 */ 
	line-height : normal; /* line-height 초기화 */ 
	padding: .8em .5em; 
	/* 원하는 여백 설정, 상하단 여백으로 높이를 조절 */ 
	font-family: inherit; /* 폰트 상속 */ 
	font-size: 30px;
	color: blue;
	border: 1px solid #999; border-radius: 10; /* iSO 둥근모서리 제거 */ 
	appearance: none; 
}
</style>
</head>
<body>
<center>
	<table width=100% height=100%>
		<form name="frmQuiz" action="quiz.asp" method="post" onsubmit="return false;">
		<tr>
			<td align=center valign=middle>
	
	<h1>WORD QUIZ</h1>
	<table width=100% onclick="javascript:viewScore();">
	<tr>
		<td align=right><div id="dScore"></div></td>
	</tr>
	<tr>
		<td align=right><div id="dScoreItem"></div></td>
	</tr>
	</table>
	
	<hr>
	<div id="dQuiz"></div>
	
	<hr>
	<h2>A. <input type="text" name="strAnswer" onkeydown="javascript:press_enter()"></h1>
	<div id="dMessage"></div>
</td>
</tr>
	</form>
</table>
</center>
</body>


<script language="javascript">
//=================힌트모드 설정 및 단어장 입력===============
var bHintMode = false;
var strQuiz = "";




//strQuiz += "only/not more than/Only three people will win an award|";
//strQuiz += "plan/an idea about how to do or accomplish something/We made a plan, and it worked|";
//strQuiz += "group/a number of people or things that are together/We sorted the objects into groups of big things and small|";
//strQuiz += "idea/a thought, suggestion, opinion, or belief about something/I have a great idea for my science project|";
//strQuiz += "unsure/not certain; having doubts/I am unsure where I left my keys|";
strQuiz += "announce/to say something loudly or publicly/The teacher announced that there will be a pop quiz today|";
strQuiz += "candidate/Someone who is being Considered for a job or is competing in an election/I watched the two presidential candidates have a debate on TV|";
strQuiz += "convince/to get someone to believe or do something/I convinced my classmate that my idea for the project was the best.|";
strQuiz += "decision/a choice that you make after thinking about it/After thinking about it for several days, Jane made the decision to buy a new car.|";
strQuiz += "elect/to choose someone for a position, job, etc., by voting/The nation voted and elected a new president last Sunday.|";
strQuiz += "estimate/to calculate something approximately/The teacher estimated that it would take her students about 15 minutes to finish the assignment.|";
//strQuiz += "government/an organization that makes laws and keeps order/|";
strQuiz += "independent/able to take care of oneself; acting or thinking freely instead of being guided or controlled by other people/|";
strQuiz += "vote/to make a choice for a leader or a law/|";
//strQuiz += "power/an ability to do something/|";
//strQuiz += "right/Something that a person is or should be allowed to have, get, or do/|";
//strQuiz += "discuss/to talk about an issue or topic/|";
strQuiz += "bar graph/a chart that uses columns of different heights to show different/|";
strQuiz += "constitution/a Written document stating the basic laws and principles of a Country, state, etc./|";
//strQuiz += "ballot/a piece of paper that you use to vote/|";





var arrQuiz = strQuiz.split("|"); //퀴즈문자열 분할
//=================힌트모드 설정 및 단어장 입력===============

var bQuizIng = true;

//총 퀴즈 수
var cntQuiz = arrQuiz.length-1

//점수계산용 변수
var bAllPassed = false;
var nTotalPass = 0;
var nTotalCount = 0;
var arrLast = new Array(cntQuiz)
var arrPass = new Array(cntQuiz) //정답카운트
var arrCount = new Array(cntQuiz) //테스트횟수
for (var i=0; i<cntQuiz; i++)
{
	arrLast[i] = false;
	arrPass[i] = 0;
	arrCount[i] = 0;
}

//현재 문제 번호(전역변수)
var currentQuiz;


function viewScore()
{
	var strScore="";
	for(var i=0; i<cntQuiz; i++)
	{
		strScore += i+1 + ". "+getAnswer(i+1).toLowerCase() + " :   " + arrPass[i] + " / " + arrCount[i] + "   "
		if (arrCount[i] > 0)
			strScore += Math.round((arrPass[i]/arrCount[i]*100),0)+"%\n"
		else
			strScore += "0%\n"
	}
	alert(strScore);
}

function checkAnswer()
{
	var strAnswer = frmQuiz.strAnswer.value;
	if(!strAnswer)
	{
		dMessage.innerHTML = "<h2><font color='red'>type [Answer]</font></h2>";
		frmQuiz.strAnswer.focus();
		return false;
	}
	
	bQuizIng=false;
	dMessage.innerHTML = "";
	if (getAnswer(currentQuiz).toLowerCase() == strAnswer.toLowerCase())
	{
		dMessage.innerHTML = "<h2><font color='green'>Correct Answer<br>(<a target='_blank' href='http://m.endic.naver.com/search.nhn?query="+getAnswer(currentQuiz)+"'>"+getAnswer(currentQuiz)+"</a>)</font></h2>";
		updateStat(true);		
	}
	else
	{
		dMessage.innerHTML = "<h2><font color='red'>Wrong Answer<br>(<a target='_blank' href='http://m.endic.naver.com/search.nhn?query="+getAnswer(currentQuiz)+"'>"+getAnswer(currentQuiz)+"</a>)</font></h2>";
		updateStat(false);
	}
	//setTimeout(nextQuiz(), 3000); 
}

function updateStat(bResult)
{
	if (bResult)
	{
		nTotalPass += 1;
		arrPass[currentQuiz-1] += 1;
		arrLast[currentQuiz-1] = true;
		
		checkAllPass();
	}
	
	nTotalCount += 1;
	arrCount[currentQuiz-1] += 1;
	
	//dScore.innerHTML = "Total: "+nTotalPass+" / "+nTotalCount+" ("+Math.round((nTotalPass/nTotalCount*100),0)+"%)"
	dScoreItem.innerHTML = "Score: "+arrPass[currentQuiz-1]+" / "+arrCount[currentQuiz-1]+" ("+Math.round((arrPass[currentQuiz-1]/arrCount[currentQuiz-1]*100),0)+"%)"	
}

function nextQuiz()
{
	
	frmQuiz.strAnswer.value = "";
	frmQuiz.strAnswer.focus();
	
	var newQuiz = currentQuiz;
	while(currentQuiz == newQuiz){
		
		newQuiz = getRndInt(1,cntQuiz)
		
		//새로구한 퀴즈의 정답율을 가져온다.
		var nPassRate = Math.round((arrPass[newQuiz-1]/arrCount[newQuiz-1])*100,0)
		
		//if(getRndInt(0,100) <= nPassRate)
			//newQuiz = getRndInt(1,cntQuiz)	
	}
	currentQuiz = newQuiz;
	
	 //랜덤 값 가져오기
	
	var strQuestion = getQuestion(currentQuiz); //문제 가져오기

	var strHTML = "";
	strHTML += "<h2>Q. "+strQuestion+"</h1>";
	dQuiz.innerHTML = strHTML
	frmQuiz.strAnswer.focus();
}

function getQuestion(currentQuiz)
{
	var arrQuizItem = arrQuiz[currentQuiz-1].split("/")
	if (bHintMode)
		return arrQuizItem[1].toLowerCase() + " <img src='/img/blank.gif' width=1px height=1px><br><h3>[" + arrQuizItem[2] + "]</h3>";
	else
		return arrQuizItem[1].toLowerCase();
}

function getAnswer(currentQuiz)
{
	var arrQuizItem = arrQuiz[currentQuiz-1].split("/")
	return arrQuizItem[0];
}

function getRndInt(min, max) 
{
	return Math.floor(Math.random() * (max - min + 1)) + min; 
}

function checkAllPass() 
{
	for (var i=0; i<cntQuiz; i++)
	{
		if(!arrLast[i])
			return false;
	}
	
	if(!bAllPassed)
	{
		alert("ALL PASSED!");
		bAllPassed = true;
	}
}

function press_enter()
{ 
	if(event.keyCode == 13) 
	{
		if(bQuizIng)
			checkAnswer();
		else
		{
			bQuizIng = true;
			dMessage.innerHTML = "";
			nextQuiz();
		}
	}
}

nextQuiz();

</script>
</html>