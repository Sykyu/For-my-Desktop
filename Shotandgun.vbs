' -------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------

' ------ Real Code Start --------------------------------------------------------------
' ------ Code Runnable Start ----------------------------------------------------------

' ------ �����p�����[�^ ---------------------------------------------------------------
Dim objShell
Dim strDate
Dim jpegname
Dim fs
Dim out
Dim mspaint
Dim strProcName
Set objShell = WScript.CreateObject("WScript.Shell")
Set fs = WScript.CreateObject("Scripting.FileSystemObject")

' ------ �o�̓t�@�C�����ݒ� -----------------------------------------------------------
strDate = Replace(Replace(Replace(Now(), "/", ""), ":", ""), " ", "")
objShell.SendKeys( "% n" )
jpegname = "�摜" + strDate + ".jpeg"
set out = fs.CreateTextFile(jpegname)
out.Close()
objShell.SendKeys( "% n" )
' ------ Output -----------------------------------------------------------------------

strProcName = "mspaint.exe "
mspaint = objShell.Run(strProcName + jpegname, 2)
WScript.sleep(250)
objShell.SendKeys( "^v" )
objShell.SendKeys( "^s" )
WScript.sleep(250)

' ------ Killtask ---------------------------------------------------------------
objShell.run "cmd /c @taskkill /IM mspaint.exe", vbHide 

' ------ Code Runnable End ------------------------------------------------------------
' ------ Real Code End ----------------------------------------------------------------
