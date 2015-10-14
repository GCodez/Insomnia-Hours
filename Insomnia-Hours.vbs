'Insomnia Hours,created by Gulam Mustafa, me_salman148@yahoo.com

Dim VarHrs 

VarHrs = InputBox("Enter the approximate Execution time in Hours","Insomnia Hours") 

If TypeName(VarHrs) = "Empty"  Then 

MsgBox("You pressed cancel")

Else If Len(Trim(VarHrs)) = 0 Then 

MsgBox("You entered nothing")

Else

exptime = DateAdd("h",VarHrs,Now) 

et=FormatDateTime(exptime,4)

tv=TimeValue(et)

MsgBox("Execution time is " & VarHrs & " hour " & "i.e until " & tv)

exptime = DateAdd("h",VarHrs,Now) 

Set Wshell = WScript.CreateObject("WScript.Shell") 

While Now < exptime 

Wscript.Sleep 120000 

Wshell.SendKeys "{NUMLOCK}" 

Wshell.SendKeys "{NUMLOCK}" 

Wend 

MsgBox("Expired at " & Time)

End If

End If
