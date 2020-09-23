Attribute VB_Name = "Module1"
'TechMenus
'Copyright (C) 2000 Lewis Anthony wilson
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

'Mail:
'Coments@TechSun.co.uk
'Bugs@TechSun.co.uk
'Employment@TechSun.co.uk


Rem ------------- WindowProc Delcrations.
Option Explicit
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal wndrpcPrev As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const GWL_WNDPROC = (-4)
Public OldWndProc As Long
Rem ------------- WindowProc Delcrations.

Public Popup As Boolean
Public RoundMenus As Boolean
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function WaitMessage Lib "user32" () As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Type POINTAPI
    X As Long
    y As Long
End Type
Public Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type


'// speed things up
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long

Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Public Const REALTIME_PRIORITY_CLASS = &H100
'// speed things up


Public Hack As Long                     'Pointer Hack for MenuCls ,used to raise events
Public Xm() As New menuObj
Public MenuID() As Variant
Public MenuCount As Long

Public Trans As Boolean

Private Const PM_REMOVE = &H1

Private Const WM_HOTKEY = &H312
Public Const WM_TIMER = &H113
Public Const WM_KEYDOWN = &H100
Public Const WM_ACTIVATE = &H6
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_SETFOCUS = &H7

Public Vdcengine As New VirtualDC

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Select Case uMsg

Case WM_ACTIVATE: MenuCheak: Exit Function
'Case WM_KEYDOWN: KeyInp Wparam
'Case WM_COMPACTING: Debug.Print "Warning System Resources Are Low. This computer may not be Capable Of Operating this Menu"
Case WM_TIMER: KeyInp
End Select

WindowProc = CallWindowProc(OldWndProc, hwnd, uMsg, wParam, ByVal lParam)   'Tell Vb To Handle the Msg if its not for us.
End Function

Function KeyInp()

WaitMessage
        
Dim X As Long
Dim message As Msg
For X = 1 To MenuCount
  
    If PeekMessage(message, Xm(X).gethwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
    If Popup = False Then Xm(X).Click: Exit Function
    If Popup = True Then Xm(X).Pop: Exit Function
    End If

Next


End Function
Sub MenuCheak()
Dim Menu As Long
For Menu = 1 To MenuCount

    If Xm(Menu).poped = True Then
        Xm(Menu).Destroy
    End If

Next Menu
End Sub

Function NewMenu(id As Variant, MenuCaption As String, MenuTextColor As Long, OwnerHwnd As Long, Hotkey As Long, keyModifier As keyMode) As Boolean

    Dim Chk As Long
    For Chk = 1 To MenuCount
        If UCase(Trim(id)) = UCase(Trim(MenuID(Chk))) Then NewMenu = False: Exit Function
    Next

MenuCount = MenuCount + 1
ReDim Preserve Xm(MenuCount)
ReDim Preserve MenuID(MenuCount)
MenuID(MenuCount) = id
Xm(MenuCount).OwnerHwnd = OwnerHwnd
Xm(MenuCount).Caption = MenuCaption
Xm(MenuCount).MenuID = id
Xm(MenuCount).TxtRGB = MenuTextColor
    

    Xm(MenuCount).Hotkey = Hotkey
    Xm(MenuCount).HotkeyMod = keyModifier


    
    NewMenu = True

End Function

Function GetMenu(ByVal id As Variant) As Long
Dim X As Long
For X = 1 To MenuCount
    If id = MenuID(X) Then GetMenu = X: Exit Function
Next X
End Function
Function GetID(Hdc As Long) As Long
Dim X As Long
For X = 1 To MenuCount
    If Xm(X).isin(Hdc) = True Then GetID = X: Exit Function
Next X
End Function
