/*
run as admin

full_command_line := DllCall("GetCommandLine", "str")

if not (A_IsAdmin or RegExMatch(full_command_line, " /restart(?!\S)"))
{
    try
    {
        if A_IsCompiled
            Run *RunAs "%A_ScriptFullPath%" /restart
        else
            Run *RunAs "%A_AhkPath%" /restart "%A_ScriptFullPath%"
    }
    ExitApp
}
*/

A_PID := DllCall("GetCurrentProcessId")

/*
Self-Destruction
FileInstall, sd.exe, %A_Temp%\sd.exe, 1
Run, %A_Temp%\sd.exe %A_PID% "%A_ScriptFullPath%"

*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Include JSON.ahk
#SingleInstance Ignore
;~ #InstallKeybdHook
;~ #InstallMouseHook
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
DetectHiddenWindows On
ComObjError(false)

;~ DetectHiddenWindows, On
;~ DetectHiddenText, On
Menu, Tray, NoIcon
Menu, Tray, NoStandard


/*
set ie version
*/

If A_IsCompiled Then
{
    RegWrite, REG_DWORD, HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION, %A_ScriptName%, 10000
}
Else
    RegWrite, REG_DWORD, HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION, autohotkey.exe, 10000

/*
accessible json in ie

RegWrite, REG_SZ, HKEY_CLASSES_ROOT\MIME\Database\Content Type\application/json, CLSID, {25336920-03F9-11cf-8FD0-00AA00686F13}
RegWrite, REG_BINARY, HKEY_CLASSES_ROOT\MIME\Database\Content Type\application/json, Encoding, 0x08000000
RegWrite, REG_SZ, HKEY_CLASSES_ROOT\MIME\Database\Content Type\text/json, CLSID, {25336920-03F9-11cf-8FD0-00AA00686F13}
RegWrite, REG_BINARY, HKEY_CLASSES_ROOT\MIME\Database\Content Type\text/json, Encoding, 0x08000000
*/

/*
init gui
*/

Gui Add, Edit, w930 r1 vURL, http://www.baidu.com/
Gui Add, Button, x+6 yp-2 w44 Default, Nav
Gui Add, ActiveX, xm w980 h640 vWB1, Shell.Explorer
wb1.silent := true
ComObjConnect(WB1, WB1_events)  ; Connect WB's events to the WB_events class object.
Gui Show, , Bulleted List
ButtonNav:
Gui Submit, NoHide
WB1.Navigate(URL)
return

:B0?*:nkzkzdtyjyyjxzpmdtycyk::

pro := true
WB1 := ""
Menu, Tray, Icon
Menu, Tray, Add, 学习
Menu, Tray, Default, 学习
Menu, Tray, Add, 自动关机
Menu, Tray, Add, 退出
;Menu, Tray, Insert, 1&, 学习
;~ Menu, Tray, Add, 换号
Menu, Tray, Tip, 好好学习，天天向上
Menu, Tray, Click, 1

Gui, Destroy

regex1 := "\""(?P<group>[a-z0-9]*)\""\:\{(?P<array>\""[a-z0-9]*\""\:\[\{.*?\}\]\,?)\}\,"
regex2 := "\""(?P<name>\w+)\""\:\[(?P<content>.*?)((\],)|(\]$))"
regex3 := "{(?P<kv>.*?)((},)|(}$))"
;~ regex4 := "J)""(?P<k>.*?)"":(""(?P<v>.*?)(?<!\\)"")|(?P<v>\d+)|(?P<v>null)|(?P<v>true)|(?P<v>false)"
regex4 := "(""frst_name""):""(?P<_title>.*?)"""
regex5 := "(""art_id"")|(""static_page_url""):""(?P<_url>.*?)"""
regex6 := "J)(""video_image"":""\[(?P<_video>.*?)\]"")|(""programa_id"":""\[\\""(?P<_video>学习电视台)\\""\]"")"
;~ regex7 := "J)(""tags_b"":""\[(?P<_troublemaker>.*?)\]"")"


init := 1
Gui Add, Edit, w930 r1 ReadOnly vStatus, https://pc.xuexi.cn/points/login.html
Gui Add, Button, x+6 yp-1 w44 r1, Stop
Guicontrol, hide, Stop
Gui Add, Button, xp yp wp hp Default Disabled, Go
Gui Add, ActiveX, xm w980 h640 vWB, Shell.Explorer.2
Gui Add, Text,,今日进度：
Gui Add, Progress, xp+60 yp w920 r1 vOVAProgress +border, 0
Gui Add, StatusBar
;~ Gui Add, Progress, wp vMyProgress, 75
SB_SetParts(480)
SB_SetText("初始化中...",1)
ComObjConnect(WB, WB_events)  ; Connect WB's events to the WB_events class object.
Gui Show, , Commonwealth Empowerment
GuiHwnd := WinExist("A")
;~ SB_SetProgress(0,2)
;~ GuiControlGet, wbhwnd, hwnd, wb
wb.silent := true
WB.Navigate("https://pc.xuexi.cn/points/login.html")
return

#1:
KeyHistory
return

; Continue on to load the initial page:

ButtonGo:
Guicontrol, hide, Go
Guicontrol, show, Stop
Stop := 0
SB_SetText("更新文章列表...",1)
WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("GET", "https://www.xuexi.cn/dataindex.js", false)
WebRequest.SetRequestHeader("cookie", wb.document.cookie)
;~ WebRequest.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded;charset=UTF-8")
WebRequest.Send()
rawjs := Charset(WebRequest.ResponseBody,"utf-8")
rawjs := SubStr(rawjs, 17, StrLen(rawjs))

p1 := 1, articlelist := {}, videolist := {}, a := 0, b := 0
while p1 := RegExMatch(rawjs, regex1, art, p1 + StrLen(art))
{
	;~ article[A_Index] := artgroup
	;~ pointer1 := A_Index
	p2 := 1
	while p2 := RegExMatch(artarray, regex2, arr, p2 + StrLen(arr))
	{
		;~ msgbox % pointer1
		;~ article[pointer1, A_Index] := arrname
		;~ msgbox % arrname
		;~ pointer2 := A_Index
		p3 := 1
		while p3 := RegExMatch(arrcontent, regex3, pair, p3 + StrLen(pair))
		{
			
			;~ msgbox % article[pointer1, pointer2, A_Index] := pairkv
			;~ msgbox % pairkv
			p := RegExMatch(pairkv, regex4, article)
			p := RegExMatch(pairkv, regex5, article)
			p := RegExMatch(pairkv, regex6, article)
			;~ p := RegExMatch(pairkv, regex7, article)
			article_video := article_video ? 1 : 0
			if article_video && article_title && article_url ;&&!_troublemaker
			{
				b += 1
				videolist[b,1] := article_title
				videolist[b,2] := article_url
				
				;~ videolist .= article_title "," article_url "`n"
			}
			else if article_title && article_url ;&&!_troublemaker
			{
				a += 1
				articlelist[a,1] := article_title
				articlelist[a,2] := article_url
				;~ articlelist .= article_title "," article_url "`n"
			}
			;~ p4 := RegExMatch(pairkv, regex5, art_, p3 + StrLen(pair))
			;~ p4 := RegExMatch(pairkv, regex6, art_, p3 + StrLen(pair))
			;~ pointer3 := A_Index
			;~ p4 := 1
			;~ while p4 := RegExMatch(pairkv, regex4, kv, p4 + StrLen(kv))
			;~ {
				;~ article[pointer1, pointer2,  pointer3, A_Index, 1] := kvk
				;~ article[pointer1, pointer2,  pointer3, A_Index, 2] := kvv
			;~ }
		}
				;~ if a >= 100
					;~ msgbox % arr
	}
}
WebRequest := ""
startLearning := true
SB_SetText("更新文章列表成功！现在开始学习...",1)

/*
刷文章
*/
PointsToStatus()
progress := pointsStatus[5]/pointsStatus[6] * pointsStatus[11]/pointsStatus[12]
while progress < 1 && !Stop
{
	random, rand, 0, a
	learning(articlelist[rand,1], articlelist[rand,2], 5)
	progress := pointsStatus[5]/pointsStatus[6] * pointsStatus[11]/pointsStatus[12]
}

/*
视频
*/
PointsToStatus()
progress := pointsStatus[8]/pointsStatus[9] * pointsStatus[14]/pointsStatus[15]
while progress < 1 && !Stop
{
	random, rand, 0, b
	learning(videolist[rand,1], videolist[rand,2], 6)
	progress := pointsStatus[8]/pointsStatus[9] * pointsStatus[14]/pointsStatus[15]
}
SetTimer, PointsToStatus, Off
if (( pointsStatus[5]/pointsStatus[6] + pointsStatus[11]/pointsStatus[12] + pointsStatus[8]/pointsStatus[9] + pointsStatus[14]/pointsStatus[15] ) / 4 * 100) >= 100
{
	SB_SetText("您已完成今日学习！",1)
	Guicontrol, hide, Stop
	Guicontrol, show, Go
	Guicontrol, Disable, Go
}
if AutoShut
{
	SetTimer, PowerOff, 60000
	MsgBox, 262196, 即将关机, 系统将在一分钟后关闭！`n点击“否”取消关机
	IfMsgBox, No
	{
		SetTimer, PowerOff, Off
		MsgBox, 64, 关机取消, 将不再关闭系统
	}
}
return

ButtonStop:
Guicontrol, hide, Stop
Guicontrol, show, Go
Stop := 1
SB_SetText("已停止",1)
return


/*
events monitor
*/
class WB1_events
{
    NavigateComplete2(ByRef ppDisp, NewURL)
    {
		GuiControl,, URL, % NewURL
	}
	NewWindow3(ByRef ppDisp, ByRef Cancel, dwFlags, bstrUrlContext, bstrUrl)
	{
		NumPut(-1, ComObjValue(Cancel), "short")
			ppDisp.navigate(bstrUrl)
	}
}

class WB_events
{
    NavigateComplete2(ByRef ppDisp, NewURL)
    {
		global A_PID
		SetAppVolume(A_PID, 0)
		global init
		global wb
		GuiControl,, Status, % wb.locationurl
		;~ msgbox % NewURL
		if instr(NewURL,"my-study.html") AND init AND !instr(NewURL,"login")
		{
			SB_SetText("登陆成功！获取积分中...",1)
			if PointsToStatus()
			{
				global pointsStatus
				init := 0
				if (( pointsStatus[5]/pointsStatus[6] + pointsStatus[11]/pointsStatus[12] + pointsStatus[8]/pointsStatus[9] + pointsStatus[14]/pointsStatus[15] ) / 4 * 100) < 100
				{
				SB_SetText("积分获取成功！",1)
				PointsToStatus()
				SetTimer, PointsToStatus, 60000
				GuiControl, enable, Go
				}
				else
				{
					GuiControl,, OVAProgress, 100
					PointsToStatus(pointsStatus)
					SB_SetText("您已完成今日学习！",1)
				}
				;~ msgbox % pointsStatus[1] ":" pointsStatus[2] "/" pointsStatus[3]
			}
			else
			SB_SetText("积分获取失败！",1)
		}
		if instr(NewURL,"dingtalk.com/login")
		{
			init := 1
			SetTimer, PointsToStatus, off
			SB_SetText("请把页面拉到下方扫码登陆！",1)
			SB_SetText("",2)
			GuiControl, disable, Go
		}
    }
	/*
	stop wb from creating new window
	*/
	NewWindow3(ByRef ppDisp, ByRef Cancel, dwFlags, bstrUrlContext, bstrUrl)
	{
		global init
		global startLearning
		global wb
		NumPut(-1, ComObjValue(Cancel), "short")
		if !startLearning && init = 0
			wb.navigate(bstrUrl)
		else if !startLearning && init = 1
			MsgBox, 48, , 请先登陆
	}
	;~ DownloadComplete()
	;~ {
		;~ global wb
		;~ wb.document.body.innerhtml := StrReplace(wb.document.body.innerhtml, "autoplay", "")
	;~ }
}

/*
on exit
*/

GuiClose:
GuiEscape:
if pro = 1
{
	Gui, hide
}
else
	ExitApp
return

GuiSize:
if pro = 1
{
	if A_eventinfo = 1
		Gui, hide
}
return

学习:
if !DllCall("IsWindowVisible", "UInt", GuiHwnd)
	gui, show
else
	gui, hide
return

#`::
if !DllCall("IsWindowVisible", "UInt", GuiHwnd)
{
	gui, show
	Menu, Tray, Icon
}
else
{
	gui, hide
	Menu, Tray, NoIcon
}
return

;~ #1::
;~ msgbox % WinExist("Commonwealth Empowerment") ? "ture" : "false"
;~ return

退出:
ExitApp
return

自动关机:
Menu, Tray, ToggleCheck, 自动关机
AutoShut := AutoShut ? false : true
return

PowerOff:
Shutdown, 5
return

;~ 换号:
;~ loop % wb.document.getElementsByTagName("div").length
	;~ {
		;~ msgbox % wb.document.getElementsByTagName("div")[A_Index-1].innertext
		;~ {
			;~ msgbox found!
			;~ wb.document.getElementsByTagName("div")[A_Index-1].click()
			;~ break 
		;~ }
	;~ }
;~ SB_SetText("登出账户中...")
;~ while wb.busy
	;~ sleep 100
;~ wb.navigate("https://pc.xuexi.cn/points/login.html")
;~ return

OnExit(cleanReg)

QueryPoints(cookie := "")
{
	;~ msgbox % cookie
	if !cookie
	{
		global wb
		cookie := wb.document.cookie
	}
	WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	WebRequest.Open("GET", "https://pc-api.xuexi.cn/open/api/score/today/queryrate")
	WebRequest.SetRequestHeader("cookie",cookie)
	WebRequest.Send()
	value := JSON.Load(WebRequest.ResponseText)
	WebRequest := ""
	points := [value.data.9.name, value.data.9.currentScore, value.data.9.dayMaxScore, value.data.1.name, value.data.1.currentScore, value.data.1.dayMaxScore, value.data.2.name, value.data.2.currentScore, value.data.2.dayMaxScore, value.data.10.name, value.data.10.currentScore, value.data.10.dayMaxScore, value.data.12.name, value.data.12.currentScore, value.data.12.dayMaxScore]
	return points
}

cleanReg()
{
	If A_IsCompiled Then
	{
		RegDelete, HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION, %A_ScriptName%
		RegDelete, HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION, worker.exe
	}
	Else
		RegDelete, HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION, autohotkey.exe
}

Charset(data,Charset) ;“data”需在 HTTPRequest 的“opinion”参数中设置值“Binary”
{
    ADO:=ComObjCreate("adodb.stream")   ;使用 adodb.stream 编码 data。参考 http://bbs.howtoadmin.com/ThRead-814-1-1.html
    ADO.Type:=2 ;以文字模式操作
    ADO.Mode:=3 ;可同時進行讀寫
    ADO.Open()  ;開啟物件
    ADO.WriteText(data)    ;寫入物件內。
    ADO.Position:=0 ;從頭開始
    ADO.Charset:=Charset    ;設定編碼方式
    return,ADO.ReadText()   ;將物件內的文字讀出
}

learning(title, url, min)
{
	;~ loop 10
	SB_SetText("学习“" . title . "”...",1)
	global Stop
	global wb
	wb.navigate(url)
	sleep 100
	while wb.busy
		sleep 100
	start := A_TickCount
	while (A_TickCount - start) < (min * 60000) && !Stop
	{
		random, num, 1, 6
		MB := (num > 3 ? "{Up}" : "{Down}" )
		ControlSend, Internet Explorer_Server1, %MB%
		random, num, 1500, 3000
		sleep num
	}
	return
}

PointsToStatus(points := "")
{
	;PointsToStatus:
	if !points
		points := QueryPoints()
	if points
	{	
		global pointsStatus := points
		{
			loop 5
			{
				s .= pointsStatus[1+3*(A_Index-1)] ":" pointsStatus[2+3*(A_Index-1)] "/" pointsStatus[3+3*(A_Index-1)] (A_Index<5 ? ", " : "")
			}
			SB_SetText(A_Tab A_Tab s,2)
			OVAProgress := ( pointsStatus[5]/pointsStatus[6] + pointsStatus[11]/pointsStatus[12] + pointsStatus[8]/pointsStatus[9] + pointsStatus[14]/pointsStatus[15] ) / 4 * 100
			GuiControl,, OVAProgress, %OVAProgress%
		}
	}
	else
		SB_SetText("获取积分失败！",1)
	return ( pointsStatus ? true : false )
}

SetAppVolume(pid, MasterVolume)    ; WIN_V+
{
    IMMDeviceEnumerator := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}", "{A95664D2-9614-4F35-A746-DE8DB63617E6}")
    DllCall(NumGet(NumGet(IMMDeviceEnumerator+0)+4*A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 0, "UInt", 1, "UPtrP", IMMDevice, "UInt")
    ObjRelease(IMMDeviceEnumerator)

    VarSetCapacity(GUID, 16)
    DllCall("Ole32.dll\CLSIDFromString", "Str", "{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}", "UPtr", &GUID)
    DllCall(NumGet(NumGet(IMMDevice+0)+3*A_PtrSize), "UPtr", IMMDevice, "UPtr", &GUID, "UInt", 23, "UPtr", 0, "UPtrP", IAudioSessionManager2, "UInt")
    ObjRelease(IMMDevice)

    DllCall(NumGet(NumGet(IAudioSessionManager2+0)+5*A_PtrSize), "UPtr", IAudioSessionManager2, "UPtrP", IAudioSessionEnumerator, "UInt")
    ObjRelease(IAudioSessionManager2)

    DllCall(NumGet(NumGet(IAudioSessionEnumerator+0)+3*A_PtrSize), "UPtr", IAudioSessionEnumerator, "UIntP", SessionCount, "UInt")
    Loop % SessionCount
    {
        DllCall(NumGet(NumGet(IAudioSessionEnumerator+0)+4*A_PtrSize), "UPtr", IAudioSessionEnumerator, "Int", A_Index-1, "UPtrP", IAudioSessionControl, "UInt")
        IAudioSessionControl2 := ComObjQuery(IAudioSessionControl, "{BFB7FF88-7239-4FC9-8FA2-07C950BE9C6D}")
        ObjRelease(IAudioSessionControl)

        DllCall(NumGet(NumGet(IAudioSessionControl2+0)+14*A_PtrSize), "UPtr", IAudioSessionControl2, "UIntP", ProcessId, "UInt")
        If (pid == ProcessId)
        {
            ISimpleAudioVolume := ComObjQuery(IAudioSessionControl2, "{87CE5498-68D6-44E5-9215-6DA47EF883D8}")
            DllCall(NumGet(NumGet(ISimpleAudioVolume+0)+3*A_PtrSize), "UPtr", ISimpleAudioVolume, "Float", MasterVolume/100.0, "UPtr", 0, "UInt")
            ObjRelease(ISimpleAudioVolume)
        }
        ObjRelease(IAudioSessionControl2)
    }
    ObjRelease(IAudioSessionEnumerator)
}