VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "menuObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function GetDesktopWindow Lib "user32" () As Integer

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
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



Option Explicit



Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
'// awsome Alpha Blending by xxxxxx

Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

'//
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'//

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_TRANSPARENT = &H20&
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)

Public alphablendLev As Long

Private Type POINTAPI
        X As Long
        y As Long
End Type
Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long



'// Font Creation
Private Const OUT_TT_PRECIS = 4
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal op As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
'// Font Creation

'//Draw Text
Private Const DT_VCENTER = &H4
Private Const DT_CENTER = &H1
Private Const DT_CALCRECT = &H400
Private Const DT_SINGLELINE = &H20
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal Hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'//Draw Text


'/Strecth Blt
Private Declare Function StretchBlt Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal Hdc As Long, ByVal nStretchMode As Long) As Long
Private Const SRCCOPY = &HCC0020
'/Strecth Blt

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'// Window Set Calls
Private Declare Function SetTextColor Lib "gdi32" (ByVal Hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'// Window Set Calls
'Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40



'//Frame Control
Private Declare Function DrawFrameControl Lib "user32" (ByVal Hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

Private Enum Draw_Control
    DFC_CAPTION = 1
    DFC_MENU = 2
    DFC_SCROLL = 3
    DFC_BUTTON = 4
End Enum

Private Enum Draw_Control_Styles
    DFCS_BUTTONPUSH = &H10
    DFCS_FLAT = &H4000
    DFCS_MONO = &H8000
    DFCS_PUSHED = &H200
    DFCS_SCROLLCOMBOBOX = &H5
    DFCS_SCROLLDOWn = &H1
    DFCS_SCROLLLEFT = &H2
    DFCS_SCROLLRight = &H3
    DFCS_SCROLLSIZEGRIP = &H8
    DFCS_SCROLLSIZEGRIPRIGHT = &H10
    DFCS_SCROLLUP = &H0
End Enum
'// End of FramControl

Private Type Size
        cx As Long
        cy As Long
End Type

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal Hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal Hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long


Private Declare Function FillRect Lib "user32" (ByVal Hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long


Rem ------- Edge
Private Enum edge
    BDR_RAISEDOUTER = &H1
    BDR_SUNKENOUTER = &H2
    BDR_RAISEDINNER = &H4
    BDR_SUNKENINNER = &H8
    BDR_OUTER = &H3
    BDR_INNER = &HC
    BDR_RAISED = &H5
    BDR_SUNKEN = &HA
    EDGE_raised = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_Sunken = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHEd = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
End Enum

'Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const BF_DIAGONAL = &H10
Private Const BF_MIDDLE = &H800    ' Fill in the middle.
Private Const BF_SOFT = &H1000     ' Use for softer buttons.
Private Const BF_ADJUST = &H2000   ' Calculate the space left over.
Private Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Private Const BF_MONO = &H8000     ' For monochrome borders.
Rem --edge
Private Declare Function DrawEdge Lib "user32" (ByVal Hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
'Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal Hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal Hdc As Long) As Long


Public OwnerHwnd As Long


Private Type Item
SubPoped As Boolean
SubItems() As Long
Caption As String
Form As Form
id As Variant
selected As Boolean
IsLink As Boolean
Isub As Boolean
OwnerPos As Long
imageHandle As Long
CopyDc As Long
End Type

Private MenuBrush As Long
Private MenuFont As Long
Private MainForm As Form
Private FooTForm As Form '// the Bottom Form. what a waste, i will fix this.

Private Items() As Item
Private ItemCount As Long
Private Footheight As Long
'Private Linc As Long
'Private Tinc As Long


'//Hlight Source picture  tmp stuff
Private Hlpic As New StdPicture
Private HlMemDC As Long
Private HlBM As BITMAP
Public Hlightimage As String
'//

'// Source Picture
Private memDc_Object As Long            '/Device Context Last Selected Object
Private memDc As Long                   '/memory Device Context
Private Bm As BITMAP
Private pic As New StdPicture
Public ItemImage As String
'// Source Picture

Private LoadOK As Boolean

Public Caption As String
Public Left As Long
Public Top As Long

Public ItemHeight As Long
Public ItemWidth As Long

Public HotkeyMod As Long
Public Hotkey As Long

Public poped As Boolean
Public MenuID As Variant

'//Text Colors
'Public SelRGB As Long
Public TxtRGB As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal Hdc As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal Hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private OwnerDc As Long
Private deskHdc As Long
Private tempDskhdc As Long



'//Text Colors


Sub Additems(Caption As String, id As Variant, Optional PicHDC As Long)
'If Not (LoadOK) Then Exit Sub

ItemCount = ItemCount + 1

ReDim Preserve Items(ItemCount)
ReDim Items(ItemCount).SubItems(0)
    Items(ItemCount).imageHandle = PicHDC
    Items(ItemCount).Caption = Caption
    Items(ItemCount).id = id
    Items(ItemCount).OwnerPos = 0
    
Set Items(ItemCount).Form = New XmenuObj
Items(ItemCount).Form.Tag = Me.MenuID
SetWindowLong Items(ItemCount).Form.hwnd, GWL_EXSTYLE, WS_EX_TOPMOST


    
    
Call SelectObject(Items(ItemCount).Form.Hdc, MenuFont)

    
End Sub
Sub Addsubitem(Caption As String, ToID As Variant, id As Variant)
Dim obj As Long
'If Not (LoadOK) Then Exit Sub
Dim X As Long
Dim SubCount As Long

ItemCount = ItemCount + 1

For X = 1 To ItemCount
        
        If ToID = Items(X).id Then
        Items(X).IsLink = True
        SubCount = UBound(Items(X).SubItems) + 1
        ReDim Preserve Items(ItemCount)
        ReDim Preserve Items(X).SubItems(SubCount)
        ReDim Items(ItemCount).SubItems(0)
        
        Items(X).SubItems(SubCount) = ItemCount
        Items(ItemCount).OwnerPos = X
        Items(ItemCount).id = id
        Items(ItemCount).Caption = Caption
        Items(ItemCount).Isub = True
        
        Set Items(ItemCount).Form = New XmenuObj
        Items(ItemCount).Form.Tag = Me.MenuID
        Items(ItemCount).Form.Width = Me.ItemWidth * Screen.TwipsPerPixelX
        SetWindowLong Items(ItemCount).Form.hwnd, GWL_EXSTYLE, WS_EX_TOPMOST
        
        
        'SetParent Items(ItemCount).Form.hwnd, OwnerHwnd
        'SetWindowPos Items(ItemCount).Form.hwnd, 0, Left, Top, Me.ItemWidth, Me.ItemHeight, 0
        SetWindowLong Items(ItemCount).Form.hwnd, GWL_EXSTYLE, WS_EX_TOPMOST
        obj = SelectObject(Items(ItemCount).Form.Hdc, MenuFont)
        End If

Next

End Sub
Sub Click()


Dim Kl As Long

    For Kl = 1 To MenuCount
        If Xm(Kl).poped = True And Xm(Kl).MenuID <> Me.MenuID Then Xm(Kl).Destroy
    Next Kl

    
    If poped = True Then
        poped = False
        Me.Destroy
        Exit Sub
    End If


If Trans = True Then Me.CaptureOwner
Debug.Print "Screen grabed"
Dim wr As RECT, pnt As POINTAPI, inc As Long
Dim X As Long
GetWindowRect MainForm.hwnd, wr
pnt.X = wr.Left
pnt.y = wr.Bottom
'ScreenToClient OwnerHwnd, pnt

inc = pnt.y


For X = 1 To ItemCount
    If Items(X).Isub = False Then
        SetWindowPos Items(X).Form.hwnd, 0, pnt.X, inc, Me.ItemWidth, Me.ItemHeight, SWP_SHOWWINDOW

        inc = inc + (Me.ItemHeight)
        DoEvents
    End If
Next


SetParent FooTForm.hwnd, 0
SetWindowPos FooTForm.hwnd, 0, pnt.X, inc, 0, 0, SWP_SHOWWINDOW Or SWP_NOSIZE

'Me.Draw_Footer
'SetFocusAPI Me.OwnerHwnd

poped = True
End Sub
Sub Draw_Footer()

Dim FR As RECT

FR.Right = FooTForm.ScaleWidth
FR.Bottom = FooTForm.ScaleHeight
DrawFrameControl FooTForm.Hdc, FR, Draw_Control.DFC_BUTTON, Draw_Control_Styles.DFCS_BUTTONPUSH

End Sub

Sub SetTextures()
On Error GoTo err
If Me.Hlightimage = "" Then Me.Hlightimage = "Fial"
If Me.ItemImage = "" Then Me.ItemImage = "Fial"

Set pic = LoadPicture(Me.ItemImage)
Set Hlpic = LoadPicture(Me.Hlightimage)

err:
If Error <> "" Then
    Unload XmenuObj
    'Set Xm = Nothing
    MsgBox "Load Fialed!"
    LoadOK = False
    Exit Sub
End If
LoadOK = True

GetObjectAPI pic.Handle, Len(Bm), Bm
GetObjectAPI Hlpic.Handle, Len(HlBM), HlBM

memDc = CreateCompatibleDC(MainForm.Hdc)
HlMemDC = CreateCompatibleDC(MainForm.Hdc)
Call SelectObject(memDc, pic.Handle)
Call SelectObject(HlMemDC, Hlpic.Handle)
End Sub
Sub CaptureOwner()
If Trans = False Then Exit Sub
Dim Desktop As Long
Dim wdth As Long
Dim hegt As Long
Desktop = GetDC(GetDesktopWindow)

wdth = Screen.Width / Screen.TwipsPerPixelX
hegt = Screen.Height / Screen.TwipsPerPixelY

BitBlt OwnerDc, 0, 0, wdth, hegt, Desktop, 0, 0, &HCC0020
End Sub
Function Pop()
    
    Me.setpos
    Me.Click
'    SetWindowPos MainForm.hwnd, 0, left, top, Me.ItemWidth, Me.ItemHeight, SWP_SHOWWINDOW
End Function
Function Init()




MainForm.Tag = "Main"
FooTForm.Tag = "Footer"
Footheight = (FooTForm.Height / Screen.TwipsPerPixelX) / 3


SetParent MainForm.hwnd, Me.OwnerHwnd


If Popup = False Then
    SetWindowPos MainForm.hwnd, 0, Left, Top, Me.ItemWidth, Me.ItemHeight, SWP_SHOWWINDOW
End If

    
    
    '////
    Dim ret As Long
    Dim ThreadID As Long
    ThreadID = GetCurrentThread
    ret = SetThreadPriority(ThreadID, THREAD_PRIORITY_HIGHEST)
    Debug.Print ret
    
    Dim TID As Long
    TID = GetCurrentProcess
    SetPriorityClass TID, REALTIME_PRIORITY_CLASS
    '////
    '
    
'    SetWindowLong MainForm.hwnd, GWL_EXSTYLE, WS_EX_TOPMOST
    Call SelectObject(MainForm.Hdc, MenuFont)
    Call RegisterHotKey(MainForm.hwnd, MainForm.hwnd, Me.HotkeyMod, Hotkey)
    


    'Dim wr As RECT
    'GetClientRect MainForm.hwnd, wr
    
    OwnerDc = Vdcengine.Create(MainForm.Hdc, 1024, 768)
    
    deskHdc = GetDC(OwnerHwnd)
    
End Function

Sub setpos()
If poped = True Then Exit Sub
Dim pnt As POINTAPI
GetCursorPos pnt
ScreenToClient OwnerHwnd, pnt
'ClientToScreen OwnerHwnd, pnt
Me.Left = pnt.X
Me.Top = pnt.y
SetWindowPos MainForm.hwnd, 0, Left, Top, Me.ItemWidth, Me.ItemHeight, SWP_SHOWWINDOW
End Sub

Sub hide()
MainForm.Visible = False
FooTForm.Visible = False
End Sub



Function gethwnd() As Long
gethwnd = MainForm.hwnd
End Function
Private Sub Frame(Rec As RECT, Color As Long, Framesize As Long, Hdc As Long)
Dim FramePen As Long, ret As Long
FramePen = CreatePen(0, Framesize, Color)
ret = SelectObject(Hdc, FramePen)
FillRect Hdc, Rec, FramePen

'Rectangle hdc, Rec.Left, Rec.Top, Rec.Right, Rec.Bottom

SelectObject Hdc, ret
DeleteObject FramePen
End Sub

Private Sub Class_Initialize()
Set MainForm = New XmenuObj
Set FooTForm = New XmenuObj

SetWindowLong MainForm.hwnd, GWL_EXSTYLE, WS_EX_TOPMOST
SetWindowLong FooTForm.hwnd, GWL_EXSTYLE, WS_EX_TOPMOST

End Sub

Private Sub Class_Terminate()
Dim X As Long
For X = 1 To ItemCount
Unload Items(X).Form
Next

DeleteObject MenuBrush
DeleteObject MenuFont
DeleteDC memDc
DeleteDC HlMemDC
Vdcengine.Destroy

Set FooTForm = Nothing
Set MainForm = Nothing
Set pic = Nothing
Set Hlpic = Nothing
SelectObject memDc, memDc_Object
SetWindowLong OwnerHwnd, GWL_WNDPROC, OldWndProc
End Sub
Sub Paint(XHdc As Long)

'Dim X As Long
Dim Tr As RECT
Dim X As Long
X = Me.GetItemByhdc(XHdc)
                       

            
            'Me.Transulute Items(X).Form.Hwnd, XHdc, X
            
            'Tr.Bottom = Items(X).Form.Height
            'Tr.Right = Items(X).Form.Width
            'Frame Tr, 0, 2, Items(X).Form.hdc

            FillDC2 Items(X).Form.Hdc, X
            Printf3 Items(X).Form.Hdc, Items(X).Caption, 1, 5, Me.TxtRGB, 0
            If UBound(Items(X).SubItems) > 0 Then Me.Draw_Nextarrow X
            





End Sub


Sub GetFocus(Hdc As Long)
Dim Tr As RECT, X As Long
Dim Tmp As Long
If poped = False Then Debug.Print "ERROR": Exit Sub
X = GetItemByhdc(Hdc)
            
            If Items(X).selected = False Then
                Tr.Bottom = Items(X).Form.ScaleHeight
                Tr.Right = Items(X).Form.ScaleWidth
        
                Items(X).selected = True
                'Items(X).Form.BackColor = SelRGB
                'Items(X).Form.Refresh
                'Hlight24 Items(X).Form.hdc
                'DrawEdge Items(X).Form.hdc, Tr, edge.EDGE_raised, BF_RIGHT Or BF_BOTTOM
                
                Blendex HlMemDC, Items(X).Form.Hdc, Me.ItemWidth, Me.ItemHeight, 0, 0, Me.ItemWidth, Me.ItemHeight
'                Call StretchBlt(Items(X).Form.hDC, 0, 0, Me.ItemWidth, Me.ItemHeight, HlMemDC, 0, 0, HlBM.bmWidth, HlBM.bmHeight, SRCCOPY)
                'Call StretchBlt(Items(X).Form.hDC, 0, 0, Me.ItemWidth, Me.ItemHeight, HlMemDC, 0, 0, HlBM.bmWidth, HlBM.bmHeight, SRCCOPY)
                Printf3 Items(X).Form.Hdc, Items(X).Caption, 1, 5, Me.TxtRGB, 0
                
                
                Else
            Exit Sub
            End If
        
        
        If UBound(Items(X).SubItems) > 0 Then Me.Draw_Nextarrow X
            Dim t As Long
            For t = 1 To ItemCount
        
                If Items(t).Form.Hdc <> Hdc And Items(t).selected = True Then
                Items(t).selected = False
                
                Me.Paint Items(t).Form.Hdc
                Exit For
                End If
                            
            Next
        
        
'// Root Tracker Mechnisim

For Tmp = 1 To ItemCount
If UBound(Items(Tmp).SubItems) > 0 Then

    If Items(Tmp).id <> Items(X).id And Items(X).Isub = False Then
        Me.Kill_Array CLng(Tmp)
    End If

End If
Next Tmp
'//


'// Sub Tracker Mechnisim
Dim owner As Long
owner = Items(X).OwnerPos
If owner <> 0 Then
    For Tmp = 1 To UBound(Items(owner).SubItems)

        With Items(Items(owner).SubItems(Tmp))
        If .IsLink = True And .SubPoped = True And .id <> Items(X).id Then Me.KillAssosiates Items(owner).SubItems(Tmp)
        End With

    Next
End If
'//


If Not (Items(X).IsLink) Then Exit Sub
If Items(X).SubPoped = True Then Exit Sub

Dim SubInc As Long
Dim wr As RECT
'Dim Pnt As POINTAPI


    GetWindowRect Items(X).Form.hwnd, wr
    
    'ScreenToClient OwnerHwnd, Pnt
    SubInc = wr.Top

Dim Xsub As Long
    For Xsub = 1 To UBound(Items(X).SubItems)
        
        With Items(Items(X).SubItems(Xsub))
            If (wr.Right + Me.ItemWidth) > Screen.Width / Screen.TwipsPerPixelY Then
            SetWindowPos .Form.hwnd, 0, wr.Left - Me.ItemWidth, SubInc, Me.ItemWidth, Me.ItemHeight, SWP_SHOWWINDOW Or SWP_NOSIZE
            Else
            SetWindowPos .Form.hwnd, 0, wr.Right, SubInc, Me.ItemWidth, Me.ItemHeight, SWP_SHOWWINDOW
            End If
        End With
        'DoEvents
        SubInc = SubInc + Me.ItemHeight
        

    Next
Items(X).SubPoped = True

End Sub

Sub Printf3(Hdc As Long, Txt As String, Xpos As Long, Ypos As Long, Color As Long, bkC As Long)
'Dim obj As Long

Dim R As RECT, sz As Size

SetTextColor Hdc, Color
GetTextExtentPoint Hdc, Txt, Len(Txt), sz
    
    R.Top = Xpos
    R.Left = Ypos
    R.Right = Ypos + sz.cx
    R.Bottom = R.Top + sz.cy

DrawText Hdc, Txt, Len(Txt), R, DT_SINGLELINE Or DT_VCENTER



End Sub
Sub paintmain()

Dim wr As RECT, pnt As POINTAPI
Dim FR As RECT
FR.Right = MainForm.ScaleWidth
FR.Bottom = MainForm.ScaleHeight
GetWindowRect MainForm.hwnd, wr
pnt.X = wr.Left
pnt.y = wr.Bottom
ScreenToClient OwnerHwnd, pnt

        DrawFrameControl MainForm.Hdc, FR, Draw_Control.DFC_BUTTON, Draw_Control_Styles.DFCS_BUTTONPUSH
        blend_main
        Printf3 MainForm.Hdc, Caption, 0, 2, 0, 0
        
        If poped = False Then
            SetParent FooTForm.hwnd, OwnerHwnd
            SetWindowPos FooTForm.hwnd, 0, pnt.X, pnt.y, Me.ItemWidth, Footheight, SWP_SHOWWINDOW
            FooTForm.Visible = True
        End If



End Sub
Sub Item_Mousedown(Hdc As Long)
MenuCheak
Dim Tmp As menuCls
    
    Call CopyMemory(Tmp, Hack, 4)
    Tmp.I_eventS Me.MenuID, Items(GetItemByhdc(Hdc)).id

Call CopyMemory(Tmp, 0&, 4)
Set Tmp = Nothing

End Sub
Function GetItemByhdc(Hdc As Long) As Long
Dim X As Long
For X = 1 To ItemCount
    If Items(X).Form.Hdc = Hdc Then GetItemByhdc = X: Exit Function
Next
End Function
Private Sub FillDC2(Hdc As Long, M As Long)

'Dim pixY As Long, pixX As Long

    'pixX = Bm.bmHeight
    'pixY = Bm.bmWidth

'Dim Ret As Long


If Trans = True Then

Dim wr As RECT
GetWindowRect Items(M).Form.hwnd, wr
Call StretchBlt(Hdc, 0, 0, Me.ItemWidth, Me.ItemHeight, OwnerDc, wr.Left, wr.Top, Me.ItemWidth, Me.ItemHeight, vbSrcCopy)
Call Blendex(memDc, Hdc, Bm.bmWidth, Bm.bmHeight, 0, 0, Me.ItemWidth, Me.ItemHeight)
Exit Sub
End If

Call StretchBlt(Hdc, 0, 0, Me.ItemWidth, Me.ItemHeight, memDc, 0, 0, Bm.bmWidth, Bm.bmHeight, SRCCOPY)
Call Blendex(memDc, Hdc, Bm.bmWidth, Bm.bmHeight, 0, 0, Me.ItemWidth, Me.ItemHeight)
'Debug.Print ret
'If ret <= 0 Then Me.Destroy: MsgBox "Aieeeee !!! Error return from **FillDC2** Function ?!?"

End Sub

Sub Destroy()
Dim X As Long
For X = 1 To ItemCount
    Items(X).Form.Visible = False
    Items(X).selected = False
    Items(X).SubPoped = False
    
Next

poped = False

If Popup = False Then Me.paintmain
If Popup = True Then Me.hide

'SetFocusAPI OwnerHwnd

End Sub
 
Sub Kill_Array(Xa As Long)

    Items(Xa).SubPoped = False
    Dim Xs As Long
    For Xs = 1 To UBound(Items(Xa).SubItems)
        Items(Items(Xa).SubItems(Xs)).Form.Visible = False
        Items(Items(Xa).SubItems(Xs)).selected = False
    Next

End Sub
Sub KillAssosiates(Item As Long)
Dim X As Long
For X = 1 To UBound(Items(Item).SubItems)
    Items(Items(Item).SubItems(X)).Form.Visible = False
    Items(Items(Item).SubItems(X)).selected = False
    Items(Item).SubPoped = False
    Me.KillAssosiates Items(Item).SubItems(X)
Next
End Sub

Sub Draw_Nextarrow(X As Long)
Dim testrect As RECT
    
    testrect.Left = Items(X).Form.ScaleWidth - 15
    testrect.Right = testrect.Left + 15
    testrect.Bottom = Items(X).Form.ScaleHeight
    'testrect.Left = testrect.Right - 15
    DrawFrameControl Items(X).Form.Hdc, testrect, Draw_Control.DFC_SCROLL, Draw_Control_Styles.DFCS_SCROLLRight Or Draw_Control_Styles.DFCS_BUTTONPUSH
    Blendex memDc, Items(X).Form.Hdc, HlBM.bmWidth, HlBM.bmHeight, testrect.Top, testrect.Left, testrect.Right, testrect.Bottom
End Sub

Function isin(Hdc As Long) As Boolean
Dim X As Long
For X = 1 To ItemCount
If Hdc = Items(X).Form.Hdc Then isin = True: Exit Function
If Hdc = MainForm.Hdc Or Hdc = FooTForm.Hdc Then isin = True: Exit Function
Next X
End Function
Sub SetFont(Optional FaceName As String)
MenuFont = CreateFont(Me.ItemHeight - 2, 0, 0, 0, 700, 0, 0, 0, 0, 0, 55, 0, 0, FaceName)

'MenuFont = CreateFont(0, 0, 0, 0, 700, 0, 0, 0, 0, 5, 0, 0, 0, FaceName)
End Sub

Sub Blendex(SrcDc As Long, dstdc As Long, SrcWidth As Long, srcHeight As Long, dsttop As Long, dstleft As Long, DstWidth As Long, DstHeight As Long)

        Dim Blend As BLENDFUNCTION, BlendLng As Long
        Blend.SourceConstantAlpha = 255 - Me.alphablendLev
        CopyMemory BlendLng, Blend, 4
        AlphaBlend dstdc, dstleft, dsttop, DstWidth, DstHeight, SrcDc, 0, 0, SrcWidth, srcHeight, BlendLng
        



End Sub

Sub blend_main()
Blendex memDc, MainForm.Hdc, Bm.bmWidth, Bm.bmHeight, 0, 0, Me.ItemWidth, Me.ItemHeight
End Sub
Sub blend_foot()
Blendex memDc, FooTForm.Hdc, FooTForm.Width, FooTForm.Height, 0, 0, FooTForm.Width, FooTForm.Height

End Sub



