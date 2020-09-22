Attribute VB_Name = "ModMisc"
Option Explicit
'to colorize then menu's
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuInfo Lib "user32" (ByVal hMenu As Long, lpcmi As bMENUINFO) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, lpcmi As bMENUINFO) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewitem As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'to keep form on top
#If Win16 Then 'Conditional Compile statements


Declare Sub SetWindowPos Lib "User" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
#Else


Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If
Public Declare Function SetWindowLong Lib "user32" _
        Alias "SetWindowLongA" (ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

Public Declare Function GetWindowLong Lib "user32" _
        Alias "GetWindowLongA" (ByVal hwnd As Long, _
        ByVal nIndex As Long) As Long

Public Declare Sub DragFinish Lib "shell32.dll" _
        (ByVal HDROP As Long)

Public Declare Function DragQueryFile Lib "shell32.dll" _
        Alias "DragQueryFileA" (ByVal HDROP As Long, _
        ByVal UINT As Long, _
        ByVal lpStr As String, _
        ByVal ch As Long) As Long

Public Declare Function CallWindowProc Lib "user32" _
        Alias "CallWindowProcA" (ByVal lplPhWnd As Long, _
        ByVal hwnd As Long, _
        ByVal Msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long



Public Const GWL_WNDPROC = (-4)
Public Const WM_DROPFILES = &H233

Public lPhWnd As Long
Public obj    As ClsDrag
Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type
Private Type bMENUINFO
    cbSize As Long
    fMask As Long
    dwStyle As Long
    cyMax As Long
    hbrBack As Long
    dwContextHelpID As Long
    dwMenuData As Long
End Type
Private Const MIM_APPLYTOSUBMENUS = &H80000000
Private Const MIM_BACKGROUND = &H2
Private Const LR_LOADFROMFILE = &H10
Private Const MF_BITMAP = &H4
Private Const MF_BYPOSITION = &H400


Public Function WndProc(ByVal lHwnd As Long, _
                        ByVal lMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long

    Dim lFileInfo As Long
    Dim sBuffer As String * 256
    Dim lRet As Long
    '--------------------
    'Look For WM_DROPFILES
    '---------------------
    If lMsg = WM_DROPFILES Then
       lFileInfo = wParam
       lRet = DragQueryFile(lFileInfo, 0, ByVal sBuffer, 256)
       obj.ReturnFile Trim(sBuffer)
       DragFinish lFileInfo
       WndProc = 0
    Else
        WndProc = CallWindowProc(lPhWnd, lHwnd, _
                                 lMsg, wParam, lParam)
    End If
    
End Function

Public Sub SetMenuBackColor(mHwnd As Long, mColor As Long, mStyle As Long, mHatch As Long, SingleMnu As Boolean, Optional MmnuItem As Long)
'modified code from PSC
    Dim Ret As Long, nCnt As Long, mnu As Integer
    Dim hMenu As Long, hSubMenu As Long
    Dim hBrush As Long
    Dim BI As LOGBRUSH
    Dim MI As bMENUINFO
    SetDefs mHwnd
    BI.lbStyle = mStyle
    BI.lbHatch = mHatch
    BI.lbColor = mColor
    hBrush = CreateBrushIndirect(BI)
    hMenu = GetMenu(mHwnd)
    nCnt = GetMenuItemCount(hMenu)
    If Not SingleMnu Then
        For mnu = 0 To nCnt - 1
            hSubMenu = GetSubMenu(hMenu, mnu)
            MI.cbSize = Len(MI)
            Ret = GetMenuInfo(hSubMenu, MI)
            MI.fMask = MIM_APPLYTOSUBMENUS Or MIM_BACKGROUND
            MI.hbrBack = hBrush
            Ret = SetMenuInfo(hSubMenu, MI)
        Next mnu
    Else
        hSubMenu = GetSubMenu(hMenu, MmnuItem)
        MI.cbSize = Len(MI)
        Ret = GetMenuInfo(hSubMenu, MI)
        MI.fMask = MIM_APPLYTOSUBMENUS Or MIM_BACKGROUND
        MI.hbrBack = hBrush
        Ret = SetMenuInfo(hSubMenu, MI)
    End If
    If fMain.chkTrue.Value = 1 Then
        Ret = GetMenuInfo(hMenu, MI)
        MI.fMask = MIM_BACKGROUND
        MI.hbrBack = hBrush
        Ret = SetMenuInfo(hMenu, MI)
        DrawMenuBar mHwnd
    End If
End Sub

Public Sub SetDefs(mHwnd As Long)
'Sets things back to normal
    Dim Ret As Long, nCnt As Long, mnu As Integer
    Dim hMenu As Long, hSubMenu As Long
    Dim hBrush As Long
    Dim MI As bMENUINFO
    hMenu = GetMenu(mHwnd)
    nCnt = GetMenuItemCount(hMenu)
    For mnu = 0 To nCnt - 1
        hSubMenu = GetSubMenu(hMenu, mnu)
        MI.cbSize = Len(MI)
        Ret = GetMenuInfo(hSubMenu, MI)
        MI.fMask = MIM_APPLYTOSUBMENUS Or MIM_BACKGROUND
        MI.hbrBack = hBrush
        Ret = SetMenuInfo(hSubMenu, MI)
    Next mnu
        Ret = GetMenuInfo(hMenu, MI)
        MI.fMask = MIM_APPLYTOSUBMENUS Or MIM_BACKGROUND
        MI.hbrBack = hBrush
        Ret = SetMenuInfo(hMenu, MI)
        DrawMenuBar mHwnd
End Sub

Public Sub SetMenuBitmap(mHwnd As Long, mImage As String, Mainmnu As Long, mSubmnu As Long)
'Not used in the example - fiddle with this one
'can work quite well if you choose the right pics
Dim hMenu As Long
Dim hSubMenu As Long
Dim lMenuID As Long
Dim hImage As Long
hMenu = GetMenu(mHwnd)
hSubMenu = GetSubMenu(hMenu, Mainmnu)
lMenuID = GetMenuItemID(hSubMenu, mSubmnu)
hImage = LoadImage(0, mImage, 0, 0, 0, LR_LOADFROMFILE)
ModifyMenu hSubMenu, mSubmnu, MF_BITMAP Or MF_BYPOSITION, lMenuID, hImage
End Sub
Sub KeepOffTop(F As Form)
    'sets the given form Off TopMost
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    SetWindowPos F.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


Sub KeepOnTop(F As Form)
    'sets the given form On TopMost
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    SetWindowPos F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
