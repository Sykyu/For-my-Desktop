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
'out = WScript.CreateObject("ADODB.Stream")
set out = fs.CreateTextFile(jpegname)
out.Close()
'objShell.Type = 1
'objShell.Open()
'objShell.SaveToFile jpegname, 2
objShell.SendKeys( "% n" )
' ------ �摜 ---------------------------------------------------------------

strProcName = "mspaint.exe "
mspaint = objShell.Run(strProcName + jpegname, 2)
'Set window1 = windows(mspaint.exe)
WScript.sleep(250)
'ret = objShell.AppActivate( mspaint )
objShell.SendKeys( "^v" )
objShell.SendKeys( "^s" )
WScript.sleep(250)

' ------ �o�̓t�@�C���� ---------------------------------------------------------------
objShell.run "cmd /c @taskkill /IM mspaint.exe", vbHide 

' ------ Code Runnable End ------------------------------------------------------------
' ------ Real Code End ----------------------------------------------------------------
