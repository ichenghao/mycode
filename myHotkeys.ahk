IfExist, myhotkey.ico
Menu, Tray, Icon, myhotkey.ico

;-------------------- To run some programs --------------------

#F2::Run D:\Portable\antrenamer2\Renamer.exe
#a::Run D:\Portable\AutoHotkey\AutoHotkeyU64.exe D:\Portable\AutoHotkey\Automation.ahk
#c::Run "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
#f::
if WinExist("chenghao@10.163.231.175 - WinSCP") or WinExist("chenghao@10.163.231.208 - WinSCP")
    WinActivate
else
    Run D:\Portable\WinSCP\WinSCP.exe "chenghao@10.163.231.175"
return
#g::Run D:\Portable\foobar2000\foobar2000.exe
#j::Run D:\Portable\FSCapture90\FSCapture.exe
#k::Run "C:\Program Files\Calibre2\calibre.exe"
#m::Run charmap.exe
#n::Run D:\Portable\Notepad++\notepad++.exe
#p::Run "D:\Portable\PDF-XChange Viewer\PDFXCview.exe"
#s::Run D:\Portable\Surface-Engineering-Calculation-Platform\Surface-Engineering-Calculation-Platform.exe

#t::          ; This is a timer
InputBox, time, Timer, Set the timer for ? minute(s)., , , 140
if ErrorLevel
    return
else
    timer:= time*60000
    Sleep,%timer%
    SoundBeep, 750, 500
    MsgBox, ,Timer, %time% minute(s) is up! Take a break.
return
#v::Run "C:\Program Files\Oracle\VirtualBox\VirtualBox.exe"
#w::Run "D:\Portable\Wrench Production Client-New\WrenchENT.exe"
#x::Run "C:\Program Files (x86)\XMind\XMind.exe"
#z::Run D:\Portable

#WheelUp::Send {Volume_Up}
#WheelDown::Send {Volume_Down}

#!h::Run "%windir%\system32\notepad.exe" %A_ScriptFullPath%          ;Edit this script
#!r::Reload                                                          ;Reload

;-------------------- To open some folds --------------------

!c::Run ::{21ec2020-3aea-1069-a2dd-08002b30309d}                     ;Control Panel
!e::Run ::{20d04fe0-3aea-1069-a2d8-08002b30309d}                     ;My Computer
!l::
;Run D:\Documents\LeadingSoft\LZMsg\HU130703000017\MsgFiles          ;LeadingSoft MsgFiles
Run D:\Documents\WeChat Files\smartermouse\FileStorage\File          ;WeChat File
Run D:\Documents\WXWork\1688853132579778\Cache\File                  ;WeChat Work File
return
!n::Run ::{7007acc7-3202-11d1-aad2-00805fc1270e}                     ;Network Connections
!q::                                                                 ;USB flash drive
DriveGet, u_pan, List, REMOVABLE
if ErrorLevel
    MsgBox, , Info, The USB flash drive does not exist., 1
else
    Run "%u_pan%:"
return
!r::Run ::{645ff040-5081-101b-9f08-00aa002f954e}                     ;Recycle Bin
!w::Run D:\Work\GPP
!x::
IfWinExist, Q-Dir
    WinActivate
else
    Run D:\Portable\Q-Dir\Q-Dir_x64.exe
return
!z::Run E:\Downloads                                                 ;Downloads

;-------------------- To run CMD.exe in current path --------------------

#!^c::                                                               ;Run cmd as admin in current path
WinGetText, text, A
word_array := StrSplit(text, "`n")
Loop % word_array.MaxIndex()
{
    FoundPos := RegExMatch(word_array[A_Index], "^Address: *")
    if FoundPos = 1
    {
        full_path := RegExReplace(word_array[A_Index],"^Address: ")
        break
    }
}
run *RunAs cmd.exe /k cd /d %full_path%
return
#!c::                                                                ;Run cmd in current path
WinGetText, text, A
word_array := StrSplit(text, "`n")
Loop % word_array.MaxIndex()
{
    FoundPos := RegExMatch(word_array[A_Index], "^Address: *")
    if FoundPos = 1
    {
        full_path := RegExReplace(word_array[A_Index],"^Address: ")
        break
    }
}
run cmd.exe /k cd /d %full_path%
return
