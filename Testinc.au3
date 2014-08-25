#include-once
#cs
#include <Constants.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <IE.au3>
#include <Clipboard.au3>
#include <Date.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>
#include <GuiTreeView.au3>
#include <GuiImageList.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <GuiTreeView.au3>
#include <File.au3>
#include <Process.au3>
#include <GuiTab.au3>
#include <GuiListView.au3>
 #include <WinAPI.au3>
#ce

Global $lFile = FileOpen(@ScriptDir & "\" & "Result.log", 1)
Global $str =""
Global $cls = ""
Global $TestName = ""
;Start function to capture the starting time of the scenario
Func Start($testCaseName)
      $TestName = $testCaseName
      Local $Stm = _DateTimeFormat(_NowCalc(), 3)
	  $str =  @CRLF
	  $str ="****************************************START********************************************" & @CRLF & $testCaseName & @CRLF &  "StartTime:"& $Stm &"" & @CRLF

EndFunc

;function tto check for the proper window is active or not
Func wincheck($fun ,$ctrl)
 	  Local $act = WinActive($ctrl)
	  if $act = 0 Then
		 Local $lFile = FileOpen(@ScriptDir & "\" & "Error.log", 1)
 		 Local $wrt = _FileWriteLog(@ScriptDir & "\" & "Error.log", "Error Opening:" & $ctrl, 1)
		 ;MsgBox("","",$wrt)
		 FileClose($lFile)
		 MsgBox($MB_OK,"Error","Error status is recorded in Error.log")
	  EndIf
 EndFunc


 ;OpenEclipse()
 Func OpenEclipse($testCaseEclipseExePath,$testCaseWorkSpacePath)
		 Run($testCaseEclipseExePath)
		 WinWaitActive("Workspace Launcher")
		 Local $win1 = WinActive("Workspace Launcher")
		 Sleep(1500)
		  If $win1 = 0 Then
			$cls = "------Cannot Open Workspace Launcher!--------"
			Close($cls)
			Exit
		 EndIf
		 AutoItSetOption ( "SendKeyDelay", 50)
		 Send($testCaseWorkSpacePath)
		 AutoItSetOption ( "SendKeyDelay", 400)
		 Send("{TAB 3}")
		 Send("{Enter}")
		 ;if $JunoOrKep = "Juno" Then
			;WinWaitActive("[Title:Java EE - Eclipse]")
		 ;Else
		 ; WinWaitActive("[Title:Java EE - Eclipse]")
		 ;EndIf
		 WinWaitActive("[Title:Java EE - Eclipse]")
		 Sleep(2000)
		 Local $win = WinActive("[Title:Java EE - Eclipse]")
		 If $win = 0 Then
			$cls = "------Cannot Open Eclipse!--------"
			Close($cls)
			Exit
		 EndIf

 EndFunc

;Creating Java Project

Func CreateJavaProject($testCaseProjectName)
	  Send("!fnd")
	  WinWaitActive("[Title:New Dynamic Web Project]")
	  Sleep(1500)
	  Local $win2 = WinActive("[Title:New Dynamic Web Project]")
       If $win2 = 0 Then
			$cls = "-----Problem in Creating New Dynamic Web Project!--------"
			Send("{Esc}")
			Close($cls)
			Exit
		 EndIf
	  ; Calling the Winchek Function to validate the proper screen
	  Local $funame, $cntrlname
	  $cntrlname = "[Title:New Dynamic Web Project]"
	  $funame = "CreateJavaProject"
	  wincheck($funame,$cntrlname)
	  AutoItSetOption ( "SendKeyDelay", 50)
	  Send($testCaseProjectName)
	  AutoItSetOption ( "SendKeyDelay", 400)
	  ;Send("{TAB 10}")
	  ;Send("{Enter}")
	  Send("!f")
	  WinWaitActive("[Title:Java EE - Eclipse]")
	   Sleep(2000)
		 Local $win3 = WinActive("[Title:Java EE - Eclipse]")
		 If $win3 = 0 Then
			$cls = "------Cannot Open Eclipse!--------"
			Close($cls)
			Exit
		 EndIf
   EndFunc

;Create JSP file
;***************************************************************
Func CreateJSPFile($testCaseJspName, $testCaseProjectName, $testCaseJspText)
sleep(3000)
Send("{APPSKEY}")
AutoItSetOption ( "SendKeyDelay", 100)
Send("{down}")
Send("{RIGHT}")
Send("{down 14}")
Send("{enter}")
Send($testCaseJspName)
;Send("{TAB 3}")
;Send("{Enter}")
Send("!f")
Local $temp = "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
Sleep(3000)
WinWaitActive($temp)
Sleep(2000)
		 Local $win4 = WinActive($temp)
		 If $win4 = 0 Then
			$cls = "---Error in Opening: "& $temp &"--------"
			Send("{Esc}")
			Close($cls)
			Exit
		 EndIf

; Calling the Winchek Function
Local $funame, $cntrlname
$cntrlname =  "Java EE - " & $testCaseProjectName & "/WebContent/" & $testCaseJspName & " - Eclipse"
$funame = "CreateJSPFile"
wincheck($funame,$cntrlname)
AutoItSetOption ( "SendKeyDelay", 100)
Send("{down 9}")
;Send($testCaseJspText)
ClipPut($testCaseJspText)
Send("^v")
AutoItSetOption ( "SendKeyDelay", 400)
Send("^+s")
EndFunc
;******************************************************************

;***************************************************************
;Function to create Azure project
;***************************************************************
Func CreateAzurePackage($testCaseAzureProjectName, $testCaseCheckJdk, $testCaseJdkPath,$testCaseCheckLocalServer, $testCaseServerPath, $testCaseServerNo)
 WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)


Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")

;MouseClick("primary",119, 490, 1)
Send("^+{NUMPADDIV}}")
Send("{APPSKEY}")
Sleep(1000)
#comments-start
if $JunoOrKep = "Juno" Then
Send("g")
Else
Send("e")
EndIf
#comments-end
Send("e")
Send("{Left}")
Send("{UP}")
;Send("{down 24}")
Send("{right}")
Send("{Enter}")
WinWaitActive("[Title:New Azure Deployment Project]")
Sleep(3000)
 Local $win6 = WinActive("[Title:New Azure Deployment Project]")
 If $win6 = 0 Then
	$cls = "------Error in Creating Azure Package(Cannot Open: New Azure Deployment Project)-------"
	Send("{Esc}")
	Close($cls)
	Exit
 EndIf
; Calling the Winchek Function
Local $funame, $cntrlname
$cntrlname =  "[Title:New Azure Deployment Project]"
$funame = "CreateAzurePackage"
wincheck($funame,$cntrlname)
AutoItSetOption ( "SendKeyDelay", 50)
Send($testCaseAzureProjectName)
AutoItSetOption ( "SendKeyDelay", 150)
Send("{TAB 3}")
Sleep(3000)
Send("{Enter}")

;JDK configuration
sleep(3000)
Local $cmp = StringCompare($testCaseCheckJdk,"Check")
   if $cmp = 0 Then
	   ControlCommand("New Azure Deployment Project","","[CLASSNN:Button5]","UnCheck", "")
	   sleep(2000)
	  ControlCommand("New Azure Deployment Project","","[CLASSNN:Button5]","Check", "")
   EndIf

Send("{TAB}")
Send("+")
Send("{End}")
Send("{BACKSPACE}")
Send($testCaseJdkPath)
Send("!N")

;Server Configuration
sleep(3000)
Local $cmp = StringCompare($testCaseCheckLocalServer,"Check")
   if $cmp = 0 Then
	   ControlCommand("New Azure Deployment Project","","[CLASSNN:Button10]","UnCheck", "")
	   sleep(2000)
	  ControlCommand("New Azure Deployment Project","","[CLASSNN:Button10]","Check", "")
   EndIf
Send("{TAB}")
Send("+")
Send("{END}")
send("{BACKSPACE}")
AutoItSetOption ( "SendKeyDelay", 100)
Send($testCaseServerPath)
AutoItSetOption ( "SendKeyDelay", 200)
Send("{TAB 2}")

 for $count = $testCaseServerNo to 0 step -1
   Send("{Down}")
Next

Send("!F")
EndFunc
;******************************************************************


;*****************************************************************
;Function to publish to cloud
;****************************************************************
Func PublishToCloud($testCaseSubscription, $testCaseStorageAccount, $testCaseServiceName, $testCaseTargetOS, $testCaseTargetEnvironment, $testCaseCheckOverwrite, $testCaseUserName, $testCasePassword)
Sleep(2000)
WinWaitActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
Sleep(3000)
 Local $win6 = WinActive("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 If $win6 = 0 Then
	$cls = "------Error in Publishing to Cloud (Cannot Open: Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse)--------"
	Close($cls)
	Exit
 EndIf
Send("{Up}")
Send("{APPSKEY}")
Sleep(1000)

Send("e")
Send("{Left}")
Send("{UP}")
;Send("{Down 21}")
Send("{Right}")
Send("{Enter}")

WinWaitActive("Publish Wizard")
Sleep(3000)
 Local $win7 = WinActive("Publish Wizard")
 If $win7 = 0 Then
	$cls = "------(Cannot Open: Publish Wizard )--------"
	Send("{Esc}")
	Close($cls)
	Exit
 EndIf
while 1
Dim $hnd =  WinGetText("Publish Wizard","")
StringRegExp($hnd,"Loading Account Settings...",1)
Local $reg = @error
if $reg > 0 Then ExitLoop
WEnd

WinActive("Publish Wizard")
Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:1]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseSubscription)

 Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:2]")
 ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseStorageAccount)


Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:3]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseServiceName)


Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:4]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseTargetOS)

Local $wnd = WinGetHandle("Publish Wizard")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:ComboBox; INSTANCE:5]")
ControlCommand($wnd,"",$wnd1,"SelectString", $testCaseTargetEnvironment)


Local $cmp = StringCompare($testCaseCheckOverwrite,"UnCheck")
   if $cmp = 0 Then
	   ControlCommand("Publish Wizard","","[CLASSNN:Button4]","Check", "")
	   sleep(3000)
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","UnCheck", "")
   Else
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","UnCheck", "")
	   sleep(3000)
	  ControlCommand("Publish Wizard","","[CLASSNN:Button4]","Check", "")
   EndIf

Send("{TAB}")
AutoItSetOption ( "SendKeyDelay", 100)
Send($testCaseUserName)
Send("{TAB}")
Send($testCasePassword)
Send("{TAB 2}")
Send($testCasePassword)
AutoItSetOption ( "SendKeyDelay", 400)
Send("{TAB}")
ControlCommand("Publish Wizard","","[CLASSNN:Button5]","Check", "")
Send("{TAB}")
Send("{Enter}")

EndFunc
;*******************************************************************************

;*****************************************************************
;Function to check the status of RDP and Open Excel
;****************************************************************
Func CheckRDPConnection()
Local $tempTime = _Date_Time_GetLocalTime()
Local $timeDateStamp = _Date_Time_SystemTimeToDateTimeStr($tempTime)
Local $RDPWindow = ControlCommand("Remote Desktop Connection","","[CLASSNN:Button1]","IsVisible", "")
;MsgBox("","",$RDPWindow,3)
if $RDPWindow = 1 Then
$str = $str & "RDP Connection:- YES" & @CRLF
;FileWrite($lFile, "RDP Connection:- YES" & @CRLF)
;_ExcelWriteCell($oExcel1, "Yes", 2, 2)
Send("{TAB 4}")
Send("{Enter}")
Else
   $str = $str & "RDP Connection:- NO" & @CRLF
   ;FileWrite($lFile, "RDP Connection:- NO" & @CRLF)
;_ExcelWriteCell($oExcel1, "No", 2, 2)
EndIf
EndFunc

;Function to check publish key word in Azure activity log and update Result Text
;**************************************************************************
Func ValidateTextAndUpdateExcel($testCaseProjectName, $testCaseValidationText)
;MouseClick("primary",565, 632, 1)
 Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysLink]")
 ControlClick($wnd,"",$wnd1,"left")

Local $string =  ControlGetText("Java EE - MyHelloWorld/WebContent/newsession.jsp - Eclipse","","[CLASS:SysLink]")
;Local $string =  ControlGetText("Web - MyHelloWorld/WebContent/index.jsp - Eclipse","","[CLASS:SysLink]")
$cmp = StringRegExp($string,'<a>Published</a>',0)

;Check in webpage and update excel
Send("{TAB}")
Sleep(3000)
Send("{Enter}")
Sleep(2000)
Send("{F6}")
Sleep(3000)
Send("^c")
Local $url = ClipGet();
Local $temp = $url & $testCaseProjectName
Sleep(5000)
Local $oIE = _IECreate($temp,1,0,1,1)
_IELoadWait($oIE)
Local $readHTML = _IEBodyReadText($oIE)
Local $iCmp = StringRegExp($readHTML,$testCaseValidationText,0)
Sleep(10000)
_IEQuit($oIE)

#cs
;Check for error
If @error = 1 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "Unable to Create the Excel Object")
    Exit
ElseIf @error = 2 Then
    MsgBox($MB_SYSTEMMODAL, "Error!", "File does not exist")
    Exit
 EndIf
#ce
if $iCmp = 1 Then
;MsgBox ($MB_SYSTEMMODAL, "Test Result", "Test Passed")
;_ExcelWriteCell($oExcel, "Test Passed" , 2, 3)
$str = $str & @CRLF & "-----Test Passed-----" & @CRLF
Else
;MsgBox ($MB_SYSTEMMODAL, "Test Result", "Test Failed")
;_ExcelWriteCell($oExcel, "Test Failed" , 2, 3)
$str = $str & @CRLF & "-----Test Failed-----" & @CRLF
EndIf



EndFunc
;*******************************************************************************


Func Close($cls)

   If $cls <> 1 Then
	  Local $Etm = _DateTimeFormat(_NowCalc(), 3)
	  $str = $str & @CRLF & $cls
	  $str = $str & @CRLF & "----Test Failed----"
	  $str = $str & @CRLF & "EndTime:"& $Etm &"" & @CRLF & " " & @CRLF & "*****************************************END*********************************************" & @CRLF
	  FileWrite($lFile, $str)
	  FileClose($lFile)
	  Local $pid = ProcessExists($TestName &".exe")
	  ProcessClose($pid)
	  Local $pid1 = ProcessExists("eclipse.exe")
	  ProcessClose($pid1)
	  ;Exit
   Else
	  Local $Etm = _DateTimeFormat(_NowCalc(), 3)
	  $str = $str & @CRLF & "EndTime:"& $Etm &"" & @CRLF & " " & @CRLF & "*****************************************END********************************************123*"
	  FileWrite($lFile, $str)
	  FileClose($lFile)
	  Local $pid = ProcessExists($TestName &".exe")
	  ProcessClose($pid)
	  Local $pid1 = ProcessExists("eclipse.exe")
	  ProcessClose($pid1)
   EndIf
   EndFunc

Func Delete()

Dim $hWnd = WinGetHandle("[CLASS:SWT_Window0]")
Local $hToolBar = ControlGetHandle($hWnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
WinActivate($hToolBar)

Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")


Local $wnd = WinGetHandle("[CLASS:SWT_Window0]")
  WinActivate($Wnd)
  Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysListView32; INSTANCE:1]")
 Sleep(3000)
 _GUICtrlListView_DeleteAllItems ($wnd1)

;MouseClick("primary",119, 490, 1)
Send("^+{NUMPADDIV}")
for $i = 6 to 1 Step - 1
   Local $chk = _GUICtrlTreeView_GetCount($hToolBar)

if $chk = 0 Then
	ExitLoop
 Else
	;MouseClick("primary",119, 490, 1)
	Local $wnd = WinGetHandle("Java EE - MyHelloWorld/WebContent/index.jsp - Eclipse")
 Local $wnd1 = ControlGetHandle($wnd, "", "[CLASS:SysTreeView32; INSTANCE:1]")
 ControlClick($wnd,"",$wnd1,"left")
   Send("^+{NUMPADDIV}")
   Send("{RIGHT}")
   Send("{DOWN}")
   Send("{UP}")
   Send("{DELETE}")
   Send("{SPACE}")
   Send("{ENTER}")

   EndIf
Next
EndFunc
