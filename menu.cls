VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "menuCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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


Public Enum keyMode
MOD_ALT = &H1
MOD_CONTROL = &H2
MOD_SHIFT = &H4
End Enum
Public AlphaBlendLevel As Long
'Public Event Key(MenuID As Variant, id As KeyCodeConstants)
Public Event SelectItem(MenuID As Variant, id As Variant)
Public Event Error(MenuID As Variant, ErrorNum As Long, Error As String)

Public OwnerHwnd As Long


'Public ImageItem As String
'Public Hlightimage As String
'Public Creadits As String
Private ret As Long
Public Sub AddMenu(MenuName As String, MenuID As Variant, MenuTextColor As Long, Optional Hotkey As KeyCodeConstants, Optional HotkeyMod As keyMode)
If NewMenu(MenuID, MenuName, MenuTextColor, OwnerHwnd, Hotkey, HotkeyMod) = False Then RaiseEvent Error(MenuID, 3, "Menu Already Exists"):

End Sub
Public Sub AddItem(MenuID As Variant, Item As String, id As Variant)

ret = GetMenu(MenuID)


If ret = 0 Then RaiseEvent Error(MenuID, 1, "Unable to Add Menu Item"): Exit Sub
Xm(ret).Additems Item, id
End Sub
Public Sub Addsubitem(MenuID, Item As String, ToID As Variant, id As Variant)


ret = GetMenu(MenuID)
If ret = 0 Then RaiseEvent Error(MenuID, 2, "Unable to Add Menu Sub Item"): Exit Sub
Xm(ret).Addsubitem Item, ToID, id

End Sub



Public Function Init() As Boolean
Creadits = "ToidyMan BigAl ThornBlade"
For ret = 1 To MenuCount
'Xm(ret).Hotkey = Hotkey
'Xm(ret).HotkeyMod = HotkeyMod

Xm(ret).Init
Next ret


If OldWndProc <> 0 Then Exit Function
'//////////////////////////////////////////////////////////////////////////
OldWndProc = SetWindowLong(OwnerHwnd, GWL_WNDPROC, AddressOf WindowProc) '/
'//////////////////////////////////////////////////////////////////////////
SetTimer OwnerHwnd, 1, 350, 0
'MsgBox "This a Product is in Preview mode !!!!" + vbCrLf + "TechSun Technologys. www.TechSun.co.uk"
End Function


Sub Set_Extents(MenuID As Variant, Left As Long, Top As Long, Hieght As Long, Width As Long)

ret = GetMenu(MenuID)
Xm(ret).Left = Left
Xm(ret).Top = Top
Xm(ret).ItemHeight = Hieght
Xm(ret).ItemWidth = Width

Xm(ret).SetFont

End Sub
Private Sub Class_Initialize()
Hack = ObjPtr(Me)

End Sub
Sub I_eventS(MenuID As Variant, id As Variant)
RaiseEvent SelectItem(MenuID, id)
End Sub
Sub Set_textures(MenuID, Image_Background As String, Image_Selected As String, Optional AlphaBlendLevel As Long = 160)

ret = GetMenu(MenuID)
Xm(ret).ItemImage = Image_Background
Xm(ret).Hlightimage = Image_Selected
Xm(ret).SetTextures
Xm(ret).alphablendLev = AlphaBlendLevel

End Sub
Function Pop(MenuID As Variant)

If Popup = True Then
    ret = GetMenu(MenuID)
    Xm(ret).Pop
End If

End Function
'Property Let RoundedMenus(round As Boolean)
'RoundMenus = round
'End Property
Property Let Popupmode(PopupMenu As Boolean)
Popup = PopupMenu
End Property

Property Let TransLucient(Transparent As Boolean)
Trans = Transparent
End Property

Private Sub Class_Terminate()

KillTimer OwnerHwnd, 1
For X = 1 To MenuCount
Call UnregisterHotKey(Xm(X).gethwnd, Xm(X).gethwnd)
Next X
End Sub

