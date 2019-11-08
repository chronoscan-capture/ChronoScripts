'https://www.chronoscan.org/doc/execute_a_program_and_wait_until_it_finishes.htm?q=c2NyaXB0&ms=AAAAAAAAAA==&mw=NDcx&st=Mg==&sct=MA==

' This example shows how to call a program and how to wait until it finishes. It can be used, for example, in a VBScript, after export to import results into another system.
 
 
 
Set WSHShell = CreateObject("WScript.Shell")
result = WSHShell.Run("yourProgram.exe yourParams",,True)
 
 
MsgBox "Finish"
 

