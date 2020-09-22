VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form fMain 
   BackColor       =   &H00D6A981&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Shredder"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6960
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbOver 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8EFF2&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4030
      Width           =   735
   End
   Begin fleShred.FlatButton cmdClear 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Clear the list without deleting!"
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ForeColor       =   16777215
      BackColor       =   11037990
      Caption         =   "Clear List"
      HasFocusRect    =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   16309975
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   4630
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtFileName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E8EFF2&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "File Name"
      Top             =   3240
      Width           =   6495
   End
   Begin VB.ListBox lstFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8EFF2&
      ForeColor       =   &H00800000&
      Height          =   2370
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
   Begin fleShred.FlatButton cmdDelete 
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      ToolTipText     =   "Delete all files in the list!"
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ForeColor       =   16777215
      BackColor       =   11037990
      Caption         =   "Delete Files"
      Enabled         =   0   'False
      HasFocusRect    =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   16309975
   End
   Begin VB.CheckBox chkTrue 
      BackColor       =   &H00D6A981&
      Caption         =   "test"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   5280
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label inf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Security level"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2760
      TabIndex        =   12
      Top             =   4080
      Width           =   960
   End
   Begin VB.Label lFC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FileCount"
      Height          =   195
      Left            =   1440
      TabIndex        =   10
      Top             =   4650
      Width           =   2865
   End
   Begin VB.Shape s1 
      BackColor       =   &H00E8EFF2&
      BackStyle       =   1  'Opaque
      Height          =   300
      Index           =   1
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   4610
      Width           =   3075
   End
   Begin VB.Label lStat 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   4650
      Width           =   795
   End
   Begin VB.Shape s1 
      BackColor       =   &H00E8EFF2&
      BackStyle       =   1  'Opaque
      Height          =   300
      Index           =   0
      Left            =   165
      Shape           =   4  'Rounded Rectangle
      Top             =   4610
      Width           =   1095
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FileList:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   540
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A86D26&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   2880
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   15
      Width           =   6735
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Controls:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   660
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected FileName:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   1365
   End
   Begin VB.Shape sppb1 
      BackColor       =   &H00A86D26&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   6735
   End
   Begin VB.Shape sp1 
      BackColor       =   &H00A86D26&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   6735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A86D26&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   6735
   End
   Begin VB.Image icon2 
      Height          =   480
      Left            =   360
      Picture         =   "fMain.frx":08CA
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icon1 
      Height          =   480
      Left            =   0
      Picture         =   "fMain.frx":0D0C
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuCL 
         Caption         =   "Clear List"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuDeleteFiles 
         Caption         =   "Delete Files"
         Shortcut        =   ^D
      End
      Begin VB.Menu SPACE1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu MnuWhat 
         Caption         =   "How to Add File..."
      End
      Begin VB.Menu SPACE2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyList As New Collection 'for the file count

Dim WithEvents obj As ClsDrag
Attribute obj.VB_VarHelpID = -1

Dim i, x As Integer
Dim TheColor As Long
Dim TheStyle As Long



Private Sub cmdClear_Click()
lstFiles.Clear 'clear the listbox
    pb1.Visible = False
    For x = 1 To MyList.Count 'remove the items 1 by 1
            MyList.Remove 1
            pb1.Value = x
            pb1.Visible = True
            txtFileName = ""
    Next
    If pb1.Value = pb1.Max Then
        pb1.Visible = False
    Else
        pb1.Visible = True
    End If
        lStat = ""
    Check
End Sub

Private Sub cmdDelete_Click()
Me.MousePointer = 11
    
    If MyList.Count = 0 Then
        MsgBox "Error: Enter a file for shredding!", vbCritical
        Me.MousePointer = 0
        pb1.Visible = False
        Exit Sub
    Else
    pb1.Visible = True
        For x = 1 To MyList.Count
        
            DeleteFile (MyList.Item(x))
            pb1.Value = x
        Next
                For x = 1 To MyList.Count
                    MyList.Remove 1
                Next
                
                    Me.MousePointer = 0
                    lstFiles.Clear
                    pb1.Visible = False
                    lFC = ""
                    txtFileName = ""
                    Check
    End If
End Sub

Private Sub Form_Load()
Set obj = New ClsDrag 'set the class
    obj.Start_SubClass (Me.lstFiles.Hwnd) 'for the dragfile function
    
    TheColor = &HD6A981 'the menu's color
    TheStyle = 0 'the style
        SetMenuBackColor Me.Hwnd, TheColor, TheStyle, 0, False 'set menu and style
    lStat = "Waiting..."
    lFC = ""
    
    Check ' the sub to make the status bar
            cmbOver.AddItem "10"
            cmbOver.AddItem "20"
            cmbOver.AddItem "30"
            cmbOver.AddItem "40"
            cmbOver.AddItem "50"
            cmbOver.AddItem "60"
            cmbOver.AddItem "70"
            cmbOver.AddItem "80"
            cmbOver.AddItem "90"
            cmbOver.AddItem "100"
            cmbOver.AddItem "110"
            cmbOver.ListIndex = 4
End Sub

Private Sub lstFiles_Click()
    txtFileName.Text = lstFiles.Text
End Sub

Private Sub mnuAbout_Click()
    MsgBox App.Comments & vbNewLine & _
            App.CompanyName & "2002" & vbNewLine & _
            "File Shredder Version:" & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine & _
            App.path
End Sub

Private Sub mnuCL_Click()
    cmdClear_Click
End Sub

Private Sub mnuDeleteFiles_Click()
    cmdDelete_Click
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub MnuWhat_Click()
    MsgBox "Too add a file simply drag & drop the file to the list box!", vbInformation
End Sub

Private Sub obj_GetFileName(ByVal sFile As String)
    lstFiles.Clear
    
    lFC = "There are " & MyList.Count + 1 & " files ready for deleting!"
    pb1.Max = MyList.Count + 1 'to make the progress bar work properly
    MyList.Add sFile 'add contents to mylist
    For i = 1 To MyList.Count
        lstFiles.AddItem MyList.Item(i)
    Next
    Check
End Sub

Public Sub Check()
    If MyList.Count = 0 Then
        lStat = "Waiting..."
        Me.Icon = icon2.Picture
        cmdDelete.Enabled = False
    Else
        lStat = "Ready..."
        Me.Icon = icon1.Picture
        cmdDelete.Enabled = True
    End If
End Sub
Function DeleteFile(path As String) 'this corrupts's files if they are
                                    'text files or sumthing like that it
                                    'overwrites the text!
    Dim i As Integer, OW As String  'variable For times To overwrite
    Dim Data1 As String, Data2 As String, Data3 As String, Data4 As String, Data5 As String, Data6 As String, Data7 As String, Data8 As String, Data9 As String, Data10 As String, Data11 As String, Data12 As String, Data13 As String, Data14 As String, Data15 As String, Data16 As String, Data17 As String, Data18 As String, Data19 As String, Data20 As String, Data21 As String, Data22 As String, Data23 As String
    Dim d1 As Long, d2 As Long, d3 As Long, d4 As Long
    Dim FinalByte As Byte 'just a byte To Do the final overwrite With
    OW = cmbOver.ListIndex * 100
    'xp-optimizer
    Data1 = Chr(234) & Chr(222) ' the variables information
    Data2 = Chr(49) & Chr(231) ' the variables information
    Data3 = Chr(49) & Chr(48) ' the variables information
    Data4 = Chr(49) & Chr(48) ' the variables information
    Data5 = Chr(49) & Chr(48) ' the variables information
    Data6 = Chr(49) & Chr(48) ' the variables information
    Data7 = Chr(49) & Chr(48) 'the variables information
    Data8 = Chr(49) & Chr(48) 'the variables information
    Data9 = Chr(49) & Chr(48) ' the variables information
    Data10 = Chr(49) & Chr(48) ' the variables information
    Data11 = Chr(49) & Chr(48) ' the variables information
    Data12 = Chr(49) & Chr(48) ' the variables information
    Data13 = Chr(46) '. the variables information
    Data14 = Chr(46) '. the variables information
    Data15 = Chr(86) 'V the variables information
    Data16 = Chr(101) 'e the variables information
    Data17 = Chr(114) 'r the variables information
    Data18 = Chr(151) '_the variables information
    Data19 = Chr(49) '1 the variables information
    Data20 = Chr(46) '. the variables information
    Data21 = Chr(48) '0 the variables information
    Data22 = Chr(46) '. the variables information
    Data23 = Chr(48) '0 the variables information
    d1 = Chr(49) & Chr(48) 'the variables information
    d2 = Chr(49) & Chr(48) 'the variables information
    d3 = Chr(49) & Chr(48) 'the variables information
    d4 = Chr(49) & Chr(48) 'the variables information
    
    Open path For Binary Access Write As #1 'open the path so we can overwrite it

    For i = 1 To OW 'a Loop
        Put #1, , d2 'overwrite
    Next i 'stop Loop

    
    For i = 1 To OW 'a Loop
        Put #1, , Data1 'overwrite
    Next i 'stop Loop


    For i = 1 To OW 'another Loop
        Put #1, , Data2 'overwrite
    Next i 'stop Loop


    For i = 1 To OW 'another Loop
        Put #1, , Data3 'overwrite
    Next i 'stop Loop


    For i = 1 To OW 'another Loop
        Put #1, , Data4 'overwrite
    Next i 'stop Loop


    For i = 1 To OW 'another Loop
        Put #1, , Data5 'overwrite
    Next i 'stop Loop


    For i = 1 To OW 'Im sure you Get the point from here on!
        'that this is just the overwriting stage
        '     !
        Put #1, , Data6
    Next i


    For i = 1 To OW
        Put #1, , Data7
    Next i


    For i = 1 To OW
        Put #1, , Data8
    Next i


    For i = 1 To OW
        Put #1, , Data9
    Next i


    For i = 1 To OW
        Put #1, , Data10
    Next i


    For i = 1 To OW
        Put #1, , Data11
    Next i


    For i = 1 To OW
        Put #1, , Data12
    Next i


    For i = 1 To OW
        Put #1, , Data13
    Next i


    For i = 1 To OW
        Put #1, , Data14
    Next i


    For i = 1 To OW
        Put #1, , Data15
    Next i


    For i = 1 To OW
        Put #1, , Data16
    Next i


    For i = 1 To OW
        Put #1, , Data17
    Next i


    For i = 1 To OW
        Put #1, , Data18
    Next i


    For i = 1 To OW
        Put #1, , Data19
    Next i


    For i = 1 To OW
        Put #1, , Data20
    Next i
    
    For i = 1 To OW
        Put #1, , Data21
    Next i
    
    For i = 1 To OW
        Put #1, , Data22
    Next i
    
    For i = 1 To OW
        Put #1, , Data23
    Next i
    
    For i = 1 To OW
        Put #1, , d1
    Next i
    
    For i = 1 To OW
        Put #1, , d3
    Next i
    
    For i = 1 To OW
        Put #1, , d4
    Next i
    
    For i = 1 To 10
        Put #1, , d4
    Next i
    
    For i = 1 To 100
        Put #1, , d4
    Next i
    
    For i = 1 To OW 'the final Loop
        Put #1, , FinalByte 'the final overwrite
    Next i 'stop final Loop
    Close #1 'close the file
    Kill path 'delete it
    

End Function


