VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub DragAcceptFiles Lib "shell32.dll" _
                    (ByVal Hwnd As Long, ByVal fAccept As Long)


Public Event GetFileName(ByVal sFile As String)
Private lHwnd     As Long


'-----#  Raise Event and return File Name #-----
Friend Sub ReturnFile(ByVal sFileName As String)
   RaiseEvent GetFileName(sFileName)
End Sub

'-----# Start looking for message WM_DROPFILES # -----------
Public Function Start_SubClass(ByVal lParam As Long) As Long
    lHwnd = lParam
    DragAcceptFiles lParam, True
    lPhWnd = SetWindowLong(lParam, GWL_WNDPROC, AddressOf WndProc)
End Function

'-----# Stop looking for message WM_DROPFILES # -----------
Public Function Stop_SubClass(ByVal lParam As Long) As Long
    SetWindowLong lParam, GWL_WNDPROC, lPhWnd
    DragAcceptFiles lParam, False
End Function


Private Sub Class_Initialize()
    Set obj = Me
End Sub

Private Sub Class_Terminate()
    Stop_SubClass (lHwnd)
End Sub


