;;; Jelec IT Remote Assistance.ahk v0.3.1
;;; The goal is to create a Windows Remote Assistance ticket, and then email that ticket and its password to the IT people.


;; initial setup, including variable creation and AHK behavioral preferences. 
#SingleInstance force

settitlematchmode,2
CoordMode,click,relative

isThisMachineLikelyOnTheCompanyNetwork := false
frequentFlyer := false

;; Now on to the bulk of the script. 
;; This next bit determines if the user is on our company network. If not, then this method of remote support will fail; making the rest of this script a waste of their time.
;; In the next version, I plan to check for the existence of a VPN program, and offer to download and run it for them.

Loop, 20 {
	;; This is looped 20 times. Although a user is likely to only have 2 to 4 network adapters, I'm not yet sure how the VPN network adapter or Microsoft's "MiniPort Adaptors" will factor into this count. I guessing they won't, but maybe they have VirtualBox installed, or HyperV, or NeoRouter, or who knows what else. This whole loop can be done 100 times in about 10 seconds, so I don't feel that 20 is excessive. Not that it matters, because most users will already be on our network using either of their first two adapters, and the loop will stop and rest of the script will continue immediately after finding a single IP address that suggests they're in our network.
	ip := A_IPAddress%A_Index%
	
	ipInteger := Str2Addr(ip)
	
	if (ipInteger [OMITTED])  ; [OMITTED]
	{
		isThisMachineLikelyOnTheCompanyNetwork = true
		break
	}
}

if (isThisMachineLikelyOnTheCompanyNetwork === false) {
	msgbox,Sorry`, but you are beyond our reach right now,You seem to be at home`, off shore`, or otherwise out of the reach of Jelec IT.`n`nPlease connect your computer to the Jelec network either by plugging it in to an Ethernet port inside of the building`, or by signing into a VPN client. `Click OK to open our VPN client.
	[OMITTED]
	[OMITTED]
	winwait,C:\Windows\system32\cmd.exe
	winclose,C:\Windows\system32\cmd.exe
}
else
{

	progress,FS10 W600 Y650,%A_Space% `n %A_Space%,Please wait,Jelec IT Remote Assistance,Segoe UI   ;; This creates a nice progress window at the bottom of screens with 768 vertical pixels. I didn't want to get it any lower, or else it might be invisible on laptops, which we have a lot of.
	sleep 2000

	ifwinnotexist,Windows Remote Assistance    ;; if it's already open, don't bother opening it again.
	{
		run %windir%\system32\msra.exe
		sleep 1000
	}
	progress,10,%A_Space% `n %A_Space%,Please `Click "Yes",Jelec IT Remote Assistance     ;; before running as admin, I can't send any input to 'privileged' windows, like msra or UAC. Maybe I could find a keyboard driver that would allow me to more closely emulate human input... like Hak5's USB Rubber Ducky. Until someone smarter than that writes such a driver, I'll just keep asking the user to escalate it for me.

;;; 	standard AHK privilege escalation pasted in the following 15 lines.
	Loop, %0%  ; For each parameter:
	{
		param := %A_Index%  ; Fetch the contents of the variable whose name is contained in A_Index.
		params .= A_Space . param
	}
	ShellExecute := A_IsUnicode ? "shell32\ShellExecute":"shell32\ShellExecuteA"
		
	if not A_IsAdmin
	{
		If A_IsCompiled
			DllCall(ShellExecute, uint, 0, str, "RunAs", str, A_ScriptFullPath, str, params , str, A_WorkingDir, int, 1)
		Else
			DllCall(ShellExecute, uint, 0, str, "RunAs", str, A_AhkPath, str, """" . A_ScriptFullPath . """" . A_Space . params, str, A_WorkingDir, int, 1)
		ExitApp
	}


	; winwait,Windows Remote Assistance
	winactivate,Windows Remote Assistance
	Send {Tab}
	sleep 500
	Send +{Tab} ;; this tab and shift-tab seems to create a cursor-object and then move said cursor to the button I want to push. Thanks to my boss for suggesting it.
	sleep 500
	Send {Enter}
	progress,20,Creating a new Remote Assistance Ticket,Please wait,Jelec IT Remote Assistance
	sleep 2000
	winactivate,Windows Remote Assistance
	send {enter}
	progress,30,Creating a new Remote Assistance Ticket,Please wait,Jelec IT Remote Assistance
	winwait,Save As
	sleep 2000
	progress,40,Saving ticket,Please wait,Jelec IT Remote Assistance
	winactivate,Save As
	send `%userprofile`%\Invitation.msrcIncident
	sleep 1000
	send {enter}
	sleep 2000
	ifwinexist,Confirm Save As ;; if there exists an existing ticket in their home directory, we can assume we've dome this before.
	{
		winactivate,Confirm Save As
		send !y
		frequentFlyer = true
	}
	progress,50,Collecting ticket information,Please wait,Jelec IT Remote Assistance
	sleep 3000
	winactivate,Windows Remote Assistance
	send {tab}
	sleep 1000
	winactivate,Windows Remote Assistance
	send ^c
	password = %Clipboard%

	progress,60,Preparing Ticket for your fab IT department,Please wait for Microsoft Outlook to open,Jelec IT Remote Assistance
	;; despite all of the preceding and following code, this next line is what really makes this whole thing worth doing. The mailto URI scheme doesn't officially support attachments, and neither does Windows' implementation of of it, apparently. Thankfully, Outlook can be passed a surprisingly robust list of flags and arguments, and everyone who is likely to have access to this script is just as likely to have Outlook installed and configured. The likelihood is in fact so great, that I doubt I'll ever need to test for its installation and configuration. So I won't.
	run "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE" -c IPM.Note /a %USERPROFILE%/Invitation.msrcIncident /m "helpdesk@jelec.com?subject=Remote`%20Assistance`%20Ticket`,`%20Password`%20`%20%password%&body=Please`%20define`%20the`%20trouble`%20here.`%0A`%0A`%0AIT`%20folk`%20are`%20notoriously`%20lousy`%20guessers`%2C`%20so`%20I`%20would`%20recommend`%20a`%20brief`%20description`%20of`%20the`%20trouble.`%0AFor`%20example`%2C`%20what`%20do`%20you`%20want`%20your`%20computer`%20to`%20do`%2C`%20and`%20what`%20is`%20it`%20doing`%20instead`%3F"

	;% ;; My usual editor doesn't understand that a backtick tells AHK to treat the following character literately; in this case, NOT assuming that a percentage sign means to resolve the enclosed text to a variable. This percent sign exists only to keep the rest of this code from turning an unhelpful shade of 'variable color'.
	winwait,Remote Assistance Ticket`, Password
	progress,FS10 W600 Y50,Waiting on Microsoft Outlook,Please wait,Jelec IT Remote Assistance    ;; Move the progress window up, giving the user space to type their hopefully descriptive email. The drastic change in location also draws attention, encouraging them to read the upcoming message telling them to type. 
	progress,70,Waiting on Microsoft Outlook,Please wait,Jelec IT Remote Assistance
	sleep 5000
	winactivate,Remote Assistance Ticket`, Password
	send ^a
	progress,80,A better description usually leads to a quicker solution,Please type,Jelec IT Remote Assistance
	
	loop
	{
		ifwinexist,Remote Assistance Ticket`, Password
		{
			winwait,Microsoft Outlook,Sorry`, we're having trouble starting Outlook,10
			ifwinexist,Microsoft Outlook,Sorry`, we're having trouble starting Outlook
			{
				winactivate
				send {enter}
			}
		}
		else
		{
			break
		}
		sleep 2000
		progress,80,Click `send whenever you are ready,Please type,Jelec IT Remote Assistance
	}
	if (7 < %a_hour% < 18)
	{
;;			Maybe in a future version I'll take timezone into account.
;;			I could use something like:
				; dif := A_Now - A_NowUTC
				; hourdif := dif / 1000 ; this bit is broken. Apparently you can't divide a negative.
				; msgbox,0,0,Hour difference between here and UTC is %hourdiff%.`n`nUTC is %A_NowUTC%
		if (1 < %a_wday% < 7)
		{
			msgbox,0,Thanks for the email,Thank you for the email.`n`nPlease call help (4357) and let us know a ticket is on its way. We would hate to keep you long.
		}
		else
		{
			msgbox,0,Oh come on`, it's our day off!,Thank you for the email.`n`nSorry we can't come to the computer right now, but we are probably busy partying or otherwise enjoying our day off. `If this is an emergency`, you can try calling our cell phones.`nHopefully we haven't left them in the other room`, or with our towels on the other end of a water park.
		}	
	}
	else
	{
		msgbox,0,What time is it?,Depending on how you want to look at it, it's either too early or too late for tech support.`n`nIf it is an emergency`, you could try our cell phones. We might be busy sleeping`, spending time with family`, or fixing our own computers. `If the latter`, we would probably answer.
	}
	progress,90,Waiting for a tech to respond to your email,Please wait,Jelec IT Remote Assistance
	winwait,Windows Remote Assistance,Yes
	progress,99,A tech has connected,Almost done!,Jelec IT Remote Assistance
	sleep 1000
	send {tab}
	sleep 300
	send {tab}
	sleep 300
	send {enter}
	progress,100,,Done!,Jelec IT Remote Assistance
	sleep 2000
}

exitapp
esc::exitapp  ;; pressing the Esc key at any time during this script will immediately halt execution and close it.


;;; Functions, go-tos, and other such non-linear bits. 


;; 		Str2Addr function curtsey of author "just me", who published it at
;; 		http://www.autohotkey.com/board/topic/71490-class-ipaddress-control-support-for-ip-address-controls/

;; 		"just me" seems to have derived inspiration from author "shajul" and their work published at 
;; 		http://www.autohotkey.com/board/topic/69586-ip-address-control-classlib-ahk-l-1101/

;; 		"shajul" based their work on a project by author "PhiLho" who's project is published at 
;; 		http://www.autohotkey.com/board/topic/15998-use-the-ip-address-control-in-your-gui/

;; 		"PhiLho" uses Microsoft's SysIPAddress32 library, and "PhiLho" apparently got help from "majknetor", who is presumably a part of the same community at autohotkey.com/boards

;;; 	Just wanted to make sure I pay my dues, and pat the head of the giant who's shoulders I stand upon 
;;;		(as well as the giants below them that I can easily see from my vantage point) to convert an IP address into a 32-bit integer. 

Str2Addr(ip) {
	If !RegExMatch(ip, "^(?:\d{1,3}\.){3}\d{1,3}$")
		Return False
	Result := 0
	Loop, Parse, ip, .
	{
		If ((A_LoopField & 0xFF) <> A_LoopField)
			Return False
		Result += A_LoopField << ((4 - A_Index) * 8)
	}
	Return Result
}
