VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "KoJD REMOTECHAR v1.2"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "                        Bot                      "
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label9"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label10"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label11"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label12"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label13"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label16"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label17"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label18"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label19"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label20"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label14"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbchattype"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbst"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtop(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtop(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtop(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtop(3)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtop(4)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtop(5)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtop(6)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtop(7)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtop(8)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Command3"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtop(9)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtop(10)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtop(11)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "chall"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "chop(0)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "chop(1)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "chop(2)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "chop(3)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "chop(4)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "chop(5)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "chop(11)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "chop(10)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "chop(9)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "chop(8)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "chop(7)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "chop(6)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "chEnable"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "tmMain"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).ControlCount=   50
      TabCaption(1)   =   "onlinehile"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label15"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "snoxd"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label21"
      Tab(2).ControlCount=   1
      Begin VB.Timer tmMain 
         Interval        =   1000
         Left            =   3000
         Top             =   360
      End
      Begin VB.CheckBox chEnable 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1920
         TabIndex        =   52
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chop 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Index           =   6
         Left            =   3820
         TabIndex        =   48
         Top             =   2280
         Width           =   255
      End
      Begin VB.CheckBox chop 
         Caption         =   "Check1"
         Height          =   195
         Index           =   7
         Left            =   3820
         TabIndex        =   47
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox chop 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Index           =   8
         Left            =   3820
         TabIndex        =   46
         Top             =   2760
         Width           =   255
      End
      Begin VB.CheckBox chop 
         Caption         =   "Check1"
         Height          =   195
         Index           =   9
         Left            =   3820
         TabIndex        =   45
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox chop 
         Caption         =   "Check1"
         Height          =   195
         Index           =   10
         Left            =   3820
         TabIndex        =   44
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox chop 
         Caption         =   "Check1"
         Height          =   195
         Index           =   11
         Left            =   3820
         TabIndex        =   43
         Top             =   3480
         Width           =   255
      End
      Begin VB.CheckBox chop 
         Caption         =   "Check1"
         Height          =   195
         Index           =   5
         Left            =   3820
         TabIndex        =   42
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chop 
         Caption         =   "Check1"
         Height          =   195
         Index           =   4
         Left            =   3820
         TabIndex        =   41
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox chop 
         Caption         =   "Check1"
         Height          =   195
         Index           =   3
         Left            =   3820
         TabIndex        =   40
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox chop 
         Caption         =   "Check1"
         Height          =   195
         Index           =   2
         Left            =   3820
         TabIndex        =   39
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chop 
         Caption         =   "Check1"
         Height          =   195
         Index           =   1
         Left            =   3820
         TabIndex        =   38
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chop 
         Caption         =   "Check1"
         Height          =   195
         Index           =   0
         Left            =   3820
         TabIndex        =   37
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chall 
         Caption         =   "Check1"
         Height          =   195
         Left            =   3820
         TabIndex        =   35
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtop 
         Height          =   285
         Index           =   11
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   34
         Text            =   "tell"
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox txtop 
         Height          =   285
         Index           =   10
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   32
         Text            =   "oktrade"
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtop 
         Height          =   285
         Index           =   9
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   30
         Text            =   "confirmpt"
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   2160
         TabIndex        =   28
         Top             =   2760
         Width           =   255
      End
      Begin VB.TextBox txtop 
         Height          =   285
         Index           =   8
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   27
         Text            =   "chat"
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtop 
         Height          =   285
         Index           =   7
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   24
         Text            =   "disband"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtop 
         Height          =   285
         Index           =   6
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   21
         Text            =   "add"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtop 
         Height          =   285
         Index           =   5
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   18
         Text            =   "goto"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtop 
         Height          =   285
         Index           =   4
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   16
         Text            =   "sitstand"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtop 
         Height          =   285
         Index           =   3
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   14
         Text            =   "spawn!"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtop 
         Height          =   285
         Index           =   2
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   12
         Text            =   "offyourpc"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtop 
         Height          =   285
         Index           =   1
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   10
         Text            =   "townoff"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtop 
         Height          =   285
         Index           =   0
         Left            =   2400
         MaxLength       =   9
         TabIndex        =   8
         Text            =   "gotown"
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lbst 
         AutoSize        =   -1  'True
         Caption         =   "DISABLED"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2400
         TabIndex        =   53
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lbchattype 
         Alignment       =   2  'Center
         Caption         =   "[ Normal ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   51
         Top             =   2835
         Width           =   735
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "www.snoxd.net"
         Height          =   195
         Left            =   -73560
         TabIndex        =   50
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "www.onlinehile.com"
         Height          =   195
         Left            =   -73680
         TabIndex        =   49
         Top             =   1800
         Width           =   1425
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Enable Remote Char       ->"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   2235
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Tell your hp,mp,lvl,exp:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   3525
         Width           =   1710
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Accept and Apply the Trade:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   3315
         Width           =   2070
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Accept Party Request:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   3075
         Width           =   1635
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "(msg)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   915
         TabIndex        =   26
         Top             =   2835
         Width           =   510
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Chat (chat"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2835
         Width           =   765
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Quit/Disband Party:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   2595
         Width           =   1425
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "(nick)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1245
         TabIndex        =   22
         Top             =   2355
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Add Party (add             )"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   2355
         Width           =   1740
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "(X,Y)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1545
         TabIndex        =   19
         Top             =   2115
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Run location ( goto            )"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2115
         Width           =   1965
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Sit/Stand up:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1875
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Respawn:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1635
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Town+Exit+Off PC:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1395
         Width           =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Town+Exit:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1140
         Width           =   840
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   4320
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         X1              =   3480
         X2              =   3480
         Y1              =   360
         Y2              =   3840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Enabled"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Town:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   885
         Width           =   450
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Knight OnLine Client"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Attach"
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "KoJD ~ onlinehile.com, snoxd.net"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   0
      TabIndex        =   4
      Top             =   4200
      Width           =   4365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nid As NOTIFYICONDATA ' trayicon variable
Private Sub chall_Click()
If chall.value = 1 Then
For i = 0 To 11
If i <> 6 And i <> 8 Then
chop(i).value = 1
End If
Next
Else
For i = 0 To 11
chop(i).value = 0
Next
End If
End Sub

Private Sub chEnable_Click()
If chEnable.value = 1 Then
tmMain.Enabled = True
lbst.Caption = "ENABLED"
lbst.ForeColor = vbGreen
Debug.Print "Hack Enabled"
Else
tmMain.Enabled = False
lbst.Caption = "DISABLED"
lbst.ForeColor = vbRed
Debug.Print "Hack Disabled"
End If
End Sub

Private Sub Command1_Click()
LoadOffsets
If AttachKO = False Then
Exit Sub
End If
Me.Show
KO_ADR_CHR = ReadLong(KO_PTR_CHR)
KO_ADR_DLG = ReadLong(KO_PTR_DLG)
FIXSNDFNC
Debug.Print "Attached KO"
End Sub

Private Sub Command2_Click()
Me.WindowState = 1
minimize_to_tray
End Sub

Private Sub Command3_Click()
Form2.Show
End Sub



Private Sub Form_Load()
ChatType = 1
Dim x As Integer
For x = 0 To 11
SetLength x
Next
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
Shell "explorer http://www.onlinehile.com/"
ElseIf SSTab1.Tab = 2 Then
Shell "explorer http://www.snoxd.net/"
End If
End Sub
Sub minimize_to_tray()
Me.Hide
nid.cbSize = Len(nid)
nid.hwnd = Me.hwnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Me.Icon ' the icon will be your Form1 project icon
nid.szTip = vbNullChar
Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim msg As Long
Dim sFilter As String
msg = x / Screen.TwipsPerPixelX
Select Case msg
Case WM_LBUTTONDOWN
Me.Show ' show form
Me.WindowState = 0
Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
Case WM_LBUTTONUP
Case WM_LBUTTONDBLCLK
Case WM_RBUTTONDOWN
Case WM_RBUTTONUP
Me.Show
Shell_NotifyIcon NIM_DELETE, nid
Case WM_RBUTTONDBLCLK
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
End Sub

Public Sub DC()
Dim pStr As String, pBytes() As Byte
pStr = "5101"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub TOWN()
Dim pStr As String, pBytes() As Byte
pStr = "48"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub OFFKO()
On Error Resume Next
Shell ("TASKKILL /F /IM KnightOnline.exe")
End Sub
Public Sub OFFPC()
On Error Resume Next
Shell ("shutdown -s -t 1")
End Sub
Public Sub RESPAWN()
Dim pStr As String, pBytes() As Byte
pStr = "1201"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub JOINPARTY()
Dim pStr As String, pBytes() As Byte
pStr = "2f0201"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub ACCEPTTRADE()
Dim pStr As String, pBytes() As Byte
pStr = "300201"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub APPLYTRADE()
Dim pStr As String, pBytes() As Byte
pStr = "3005"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub DISBAND()
Dim pStr As String, pBytes() As Byte
pStr = "2f05"
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Sub
Public Sub GOTOXY(x As Integer, y As Integer)
Dim num_1 As Single, num_2 As Single
num_1 = CSng(x + 1)
num_2 = CSng(y + 1)
WriteLong KO_ADR_CHR + OffsetKO_OFF_MOVTYPE, 1
WriteFloat KO_ADR_CHR + OffsetKO_OFF_X, num_1
WriteFloat KO_ADR_CHR + OffsetKO_OFF_Y, num_2
WriteLong KO_ADR_CHR + OffsetKO_OFF_MVCHRTYP, 2
End Sub

Private Sub tmMain_Timer()
On Error Resume Next
Dim pBytes() As Byte
Dim pSize As Long

For i = 0 To 11
If chop(i).value = 1 Then
If i <> 5 And i <> 6 And i <> 8 Then
ReceivedMSG(i) = ""
pSize = msgLength(i)
ReDim pBytes(1 To pSize)
ReadByteArray ADR_LASTMSG, pBytes, pSize
Dim x As Long
For x = 1 To pSize
ReceivedMSG(i) = ReceivedMSG(i) & Hex(pBytes(x))
Next
ReceivedMSG(i) = HexToStr(ReceivedMSG(i))
Debug.Print "-" & i & "-:" & txtop(i).Text & "(" & msgLength(i) & ") <=> " & ReceivedMSG(i) & "(" & Len(ReceivedMSG(i)) & ")"
If ReceivedMSG(i) = txtop(i).Text Then
Debug.Print "Executed:" & i
Action (i)
End If

ElseIf i = 5 Then
ReceivedMSG(i) = ""
pSize = msgLength(i) + 12 'goto (xxxx,yyyy) => (fix+12)
ReDim pBytes(1 To pSize)
ReadByteArray ADR_LASTMSG, pBytes, pSize
For x = 1 To pSize
ReceivedMSG(i) = ReceivedMSG(i) & Hex(pBytes(x))
Next
ReceivedMSG(i) = HexToStr(ReceivedMSG(i))
ReceivedX = Mid(ReceivedMSG(i), Int(msgLength(i) + 3), 4) 'abcd (xxxx,yyyy) -> 61 62 63 64 20 28
ReceivedY = Mid(ReceivedMSG(i), Int(msgLength(i) + 8), 4) 'abcd (xxxx,yyyy) -> 61 62 63 64 20 28
Debug.Print txtop(i).Text & "(" & msgLength(i) & ") <=> " & ReceivedMSG(i) & "(" & Len(ReceivedMSG(i)) & ")"
If Left(ReceivedMSG(i), msgLength(i)) = txtop(i).Text Then
Debug.Print "Executed:" & i
Action (i)
End If
End If
End If
Next

End Sub

Public Sub SetLength(Index As Integer)
msgLength(Index) = Hex(Len(txtop(Index).Text))
Debug.Print "msgLength(" & Index & "):" & msgLength(Index) & " | (" & txtop(Index).Text & ")"
End Sub

Private Sub txtop_Change(Index As Integer)
SetLength Index
End Sub

Public Sub FIXSNDFNC()
'1931689751 => &H474A20
'3531683606 => &H474780
WriteLong &HB8AB00, 1931689751
KO_SND_FNC = &H474A20
End Sub
