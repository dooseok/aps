<%
Function GoogleSendMail(strTo, strFrom, strSubject, strBody) 

    On Error Resume Next 

    Set iMsg = CreateObject("CDO.Message") 

    Set iConf = CreateObject("CDO.Configuration") 

    Set Flds = iConf.Fields

    schema = "http://schemas.microsoft.com/cdo/configuration/"

    Flds.Item(schema & "sendusing")=2

    Flds.Item(schema & "smtpaccountname") = "cambi78dev <cambi78dev@gmail.com>"

    Flds.Item(schema & "sendemailaddress") = "cambi78dev <cambi78dev@gmail.com>"

    Flds.Item(schema & "smtpuserreplyemailaddress") = "cambi78dev <cambi78dev@gmail.com>"

    Flds.Item(schema & "smtpserver") = "smtp.gmail.com"

    Flds.Item(schema & "smtpserverport") = 465

    Flds.Item(schema & "smtpauthenticate") = 1

    Flds.Item(schema & "sendusername") = "cambi78dev"

    Flds.Item(schema & "sendpassword") = "wiisl4olc."

    Flds.Item(schema & "smtpusessl") = 1

    Flds.Update 

    Set Flds = Nothing

    Set iMsg = Server.CreateObject("CDO.Message") 

    With iMsg

        .Configuration = iConf

        .To = strTo ' 받는넘

        .From = strFrom ' 보내는넘

        .Subject = strSubject ' 제목

        .HTMLBody = strBody ' 내용

        SendEmailGmail = .Send

    End With

    set iMsg = nothing 

    set iConf = nothing 

    set Flds = nothing 

    If Err.number <> 0 Then

        GoogleSendMail = Err.Description

    Else

        GoogleSendMail = 0

    End If

End Function



    Ret = GoogleSendMail("cambi78@naver.com","cambi78dev <cambi78dev@gmail.com>",request("msg"),request("msg"))

    response.write Ret
%>
<script language="javascript">

setTimeout(window.open('about:blank','_self').close(), 2000);

</script>