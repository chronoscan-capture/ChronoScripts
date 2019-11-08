'https://www.chronoscan.org/doc/running_a_vbscript_in_daemon_mode.htm?q=c2NyaXB0&ms=AAAAAAAAAA==&mw=NDcx&st=Mg==&sct=MA==

'This example shows how to setup an vbs script to run some code in an infinite loop:
 
'Defining this helper function on your VBS will ensure a minimal CPU usage while sleeping:
 
Sub subSleep(strSeconds) ' subSleep(2)
    Dim objShell
    Dim strCmd
    If strSeconds <= 1 Then
        strSeconds = 2
    End If
    objShell = CreateObject("wscript.Shell")
    strCmd = "%COMSPEC% /c ping -n " & strSeconds & " 127.0.0.1>nul"
    objShell.Run(strCmd, 0, 1)
End Sub
 
 
' Enclose your main betwen a Do Loop
 
x = 1
 
Do
' Execute your tasks here.....
    ChronoApp.AddToOutputWindow "Executing Steep " & x
 
    subSleep 5 ' in seconds
 
    x = x + 1
Loop