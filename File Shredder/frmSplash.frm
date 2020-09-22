VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00D6A981&
   BorderStyle     =   0  'None
   ClientHeight    =   4020
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7275
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr1 
      Interval        =   1000
      Left            =   5640
      Top             =   600
   End
   Begin VB.Label r1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading.."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2280
      TabIndex        =   9
      Top             =   3240
      Width           =   675
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CompanyProduct"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E8EFF2&
      Height          =   435
      Left            =   2475
      TabIndex        =   8
      Top             =   705
      Width           =   3000
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LicenseTo"
      ForeColor       =   &H00E8EFF2&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   6855
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EFA88A&
      Height          =   780
      Left            =   2640
      TabIndex        =   6
      Top             =   1140
      Width           =   2550
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Platform"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E8EFF2&
      Height          =   375
      Left            =   5595
      TabIndex        =   5
      Top             =   2340
      Width           =   1380
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E8EFF2&
      Height          =   285
      Left            =   6060
      TabIndex        =   4
      Top             =   2700
      Width           =   915
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning....Im not held responsible for any data loss or damage created!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   195
      Left            =   270
      TabIndex        =   3
      Top             =   3660
      Width           =   5295
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      BackColor       =   &H00F8DED7&
      Caption         =   "RD-Software"
      ForeColor       =   &H00A86D26&
      Height          =   225
      Left            =   4680
      TabIndex        =   2
      Top             =   3270
      Width           =   2415
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright: 2002 RD-soft"
      ForeColor       =   &H00E8EFF2&
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   3060
      Width           =   2415
   End
   Begin VB.Image imgLogo 
      Height          =   3300
      Left            =   240
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1890
   End
   Begin VB.Label lbInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00F8DED7&
      Caption         =   $"frmSplash.frx":65E6
      ForeColor       =   &H00A86D26&
      Height          =   975
      Left            =   2280
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Shape sMain 
      BackColor       =   &H00A86D26&
      BackStyle       =   1  'Opaque
      Height          =   3975
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   7215
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim s As Integer
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
KeepOnTop frmSplash

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblLicenseTo = "LicensedTo: Pscode User's! http://www.pscode.com"
    lblCompanyProduct = App.CompanyName
    
    frmSplash.Show
    frmSplash.Refresh
    
    Load fMain
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub tmr1_Timer()
On Error Resume Next
Dim lod As String
s = s + 1
lod = "Loading in.." & s & " seconds."
    r1.Caption = lod
        If s = 6 Then
            r1.Caption = "Bye... && Enjoy! =)"
        ElseIf s = 7 Then
            tmr1.Enabled = False
            fMain.Show
            Unload Me
        End If
        
End Sub
