VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "Tech Menus Version 5.9"
   ClientHeight    =   6315
   ClientLeft      =   1170
   ClientTop       =   1275
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   10125
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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

'greetz to toidyman #visualbasic for all he awsome help and design input
' and putting up with my persistant Bitching :>

'greets to bigal #visualbasic for helping me with various gdi related things

'greets to sam from #VbChat irc.vbchat.com
'who supplyed to the awsome Alphablending Apis (Win98/2K/NT only) unfortunatly
'Big Thanks to Tony kemper For His Virtual DC Class submited via planetsource code






Private WithEvents Xmenu As menuCls
Attribute Xmenu.VB_VarHelpID = -1

Private Sub Form_Load()
Set Xmenu = New menuCls
Xmenu.OwnerHwnd = Me.hWnd       '// This can be 0 it will work, but subclassing is lost !~
Xmenu.Popupmode = True
Xmenu.TransLucient = True

Xmenu.AddMenu "File", "TMenu", RGB(255, 255, 255), vbKeyF1, MOD_ALT
Xmenu.Set_Extents "TMenu", 0, 0, 15, 100
Xmenu.Set_textures "TMenu", App.Path + "\TMP3.jpg", App.Path + "\tmp5.jpg", 126

Xmenu.AddItem "TMenu", "&AddUser", "Opt1"
Xmenu.AddItem "TMenu", "&EditUser", "Opt2"
Xmenu.AddItem "TMenu", "&DeleteUser", "Opt3"





'// Menu 2
Xmenu.AddMenu "Edit", "TMenu2", RGB(255, 255, 255), vbKeyF2, MOD_ALT
Xmenu.Set_Extents "TMenu2", 100, 0, 15, 100
Xmenu.Set_textures "TMenu2", App.Path + "\tmp3.jpg", App.Path + "\tmp5.jpg"

Xmenu.AddItem "TMenu2", "&Undo", "Opt1"
Xmenu.AddItem "TMenu2", "&Copy", "Opt2"
Xmenu.AddItem "TMenu2", "&Paste", "Opt3"
Xmenu.AddItem "TMenu2", "&Cut-", "Opt4"


'////////////
Xmenu.Init '/                       'Initilise Everything... hello world
'////////////







End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Xmenu.Pop "TMenu"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Xmenu = Nothing
End Sub

Private Sub Xmenu_SelectItem(MenuID As Variant, id As Variant)
Debug.Print MenuID & " " & id
End Sub

'Sub CreateTimeLine()
'Dim days As Long
'Dim daywidth As Long
'Dim X As Long
'Dim lp As Long
'Picture1.Cls
'days = getnumdays(Month(Now))

'daywidth = Picture1.ScaleWidth / days


'For X = 1 To days
'Debug.Print days
'Picture1.Line (lp, 0)-(10, lp + daywidth), QBColor(1), B
'Picture1.Line (lp, 0)-(lp + daywidth, 20), QBColor(1), B
'If days = Day(Now) Then Picture1.Line (lp, 0)-(lp + daywidth, 20), QBColor(1), BF
'Picture1.CurrentX = lp
'Picture1.CurrentY = 0
'Picture1.Print CStr(X)
'lp = lp + daywidth
'Next X

'End Sub
Function getnumdays(mounth As String)
If mounth = 2 Then getnumdays = 28: Exit Function

If mounth = 9 Then getnumdays = 30: Exit Function
If mounth = 4 Then getnumdays = 30: Exit Function
If mounth = 6 Then getnumdays = 30: Exit Function
If mounth = 11 Then getnumdays = 30: Exit Function

getnumdays = 31

End Function
