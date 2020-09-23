VERSION 5.00
Begin VB.Form XmenuObj 
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00C00000&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   16
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   167
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "XmenuObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Dpos As Long
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.Tag = "Footer" Or Me.Tag = "Main" Then Exit Sub
Dim ret As Long
ret = GetID(Me.Hdc)





'Dim Tmp As menuCls
'Dim Kn As Long
 '   Call CopyMemory(Tmp, Hack, 4)
  '  Kn = KeyCode
   ' Tmp.I_eventK Ret, Kn

'Call CopyMemory(Tmp, 0&, 4)
'Set Tmp = Nothing

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Me.Tag = "Footer" Then Exit Sub
Dim ret As Long
ret = GetID(Me.Hdc)

If Me.Tag = "Main" Then Xm(ret).Click: Exit Sub
Xm(ret).Item_Mousedown Me.Hdc

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim ret As Long
ret = GetID(Me.Hdc)
If ret = 0 Then MsgBox "Shit!!"
If Me.Tag <> "Main" And Me.Tag <> "Footer" Then Xm(ret).GetFocus Me.Hdc

End Sub

Private Sub Form_Paint()
Dim ret As Long
ret = GetID(Me.Hdc)

If Me.Tag = "Main" Then Xm(ret).paintmain: Exit Sub
If Me.Tag = "Footer" Then Xm(ret).Draw_Footer: Exit Sub

Xm(ret).Paint Me.Hdc


End Sub
