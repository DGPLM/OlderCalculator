VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form2 
   Caption         =   "三线摆测刚体的转动惯量实验数据处理计算器"
   ClientHeight    =   7500
   ClientLeft      =   8685
   ClientTop       =   3450
   ClientWidth     =   10680
   LinkTopic       =   "Form2"
   ScaleHeight     =   7500
   ScaleWidth      =   10680
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   13361
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "直接测量量计算"
      TabPicture(0)   =   "Form2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(6)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label2(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label2(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Image1(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text16"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text10(6)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text10(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text10(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text10(3)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text10(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text10(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text10(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Command6"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(4)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(5)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Command5"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Command4"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Command3"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Command2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Command1"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text9"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text8"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text7"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text6(6)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text6(5)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text6(4)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text6(3)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text6(2)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text6(1)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text6(0)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text5(6)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text5(5)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text5(4)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text5(3)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text5(2)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text5(1)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text5(0)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text4(6)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Text4(5)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Text4(4)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Text4(3)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Text4(2)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Text4(1)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Text4(0)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Text3(6)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Text3(5)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Text3(4)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Text3(3)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Text3(2)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Text3(1)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Text3(0)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "Text2(6)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "Text2(5)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "Text2(4)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "Text2(3)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "Text2(2)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "Text2(1)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "Text2(0)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "Text1(6)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "Text1(3)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "Text1(2)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "Text1(1)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "Text1(0)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).ControlCount=   72
      TabCaption(1)   =   "间接测量量计算及转动惯量的计算"
      TabPicture(1)   =   "Form2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text18"
      Tab(1).Control(1)=   "Command8"
      Tab(1).Control(2)=   "Text17"
      Tab(1).Control(3)=   "Text15"
      Tab(1).Control(4)=   "Text14"
      Tab(1).Control(5)=   "Text13"
      Tab(1).Control(6)=   "Text12"
      Tab(1).Control(7)=   "Text11"
      Tab(1).Control(8)=   "Command7"
      Tab(1).Control(9)=   "Label2(15)"
      Tab(1).Control(10)=   "Label2(14)"
      Tab(1).Control(11)=   "Label2(13)"
      Tab(1).Control(12)=   "Label2(12)"
      Tab(1).Control(13)=   "Label2(11)"
      Tab(1).Control(14)=   "Label2(9)"
      Tab(1).Control(15)=   "Label2(8)"
      Tab(1).Control(16)=   "Image1(1)"
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "平行轴定理的验证"
      TabPicture(2)   =   "Form2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text21(0)"
      Tab(2).Control(1)=   "Text24"
      Tab(2).Control(2)=   "Command11"
      Tab(2).Control(3)=   "Command10"
      Tab(2).Control(4)=   "Command9"
      Tab(2).Control(5)=   "Text23(5)"
      Tab(2).Control(6)=   "Text23(4)"
      Tab(2).Control(7)=   "Text23(3)"
      Tab(2).Control(8)=   "Text23(2)"
      Tab(2).Control(9)=   "Text23(1)"
      Tab(2).Control(10)=   "Text23(0)"
      Tab(2).Control(11)=   "Text22(5)"
      Tab(2).Control(12)=   "Text22(4)"
      Tab(2).Control(13)=   "Text22(3)"
      Tab(2).Control(14)=   "Text22(2)"
      Tab(2).Control(15)=   "Text22(1)"
      Tab(2).Control(16)=   "Text22(0)"
      Tab(2).Control(17)=   "Text21(5)"
      Tab(2).Control(18)=   "Text21(4)"
      Tab(2).Control(19)=   "Text21(3)"
      Tab(2).Control(20)=   "Text21(2)"
      Tab(2).Control(21)=   "Text21(1)"
      Tab(2).Control(22)=   "Text20(5)"
      Tab(2).Control(23)=   "Text20(4)"
      Tab(2).Control(24)=   "Text20(3)"
      Tab(2).Control(25)=   "Text20(2)"
      Tab(2).Control(26)=   "Text20(1)"
      Tab(2).Control(27)=   "Text20(0)"
      Tab(2).Control(28)=   "Text19(5)"
      Tab(2).Control(29)=   "Text19(4)"
      Tab(2).Control(30)=   "Text19(3)"
      Tab(2).Control(31)=   "Text19(2)"
      Tab(2).Control(32)=   "Text19(1)"
      Tab(2).Control(33)=   "Text19(0)"
      Tab(2).Control(34)=   "Label2(20)"
      Tab(2).Control(35)=   "Label2(19)"
      Tab(2).Control(36)=   "Label2(18)"
      Tab(2).Control(37)=   "Label2(17)"
      Tab(2).Control(38)=   "Label2(16)"
      Tab(2).Control(39)=   "Label2(10)"
      Tab(2).Control(40)=   "Image1(2)"
      Tab(2).ControlCount=   41
      Begin VB.TextBox Text1 
         Height          =   350
         Index           =   0
         Left            =   1920
         TabIndex        =   102
         Text            =   "25.393"
         Top             =   1320
         Width           =   750
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Index           =   1
         Left            =   2880
         TabIndex        =   101
         Text            =   "25.347"
         Top             =   1320
         Width           =   750
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Index           =   2
         Left            =   3840
         TabIndex        =   100
         Text            =   "25.356"
         Top             =   1320
         Width           =   750
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Index           =   3
         Left            =   4800
         TabIndex        =   99
         Text            =   "25.326"
         Top             =   1320
         Width           =   750
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFC0&
         Height          =   350
         Index           =   6
         Left            =   8040
         TabIndex        =   98
         Top             =   1320
         Width           =   750
      End
      Begin VB.TextBox Text2 
         Height          =   350
         Index           =   0
         Left            =   1920
         TabIndex        =   97
         Top             =   1800
         Width           =   750
      End
      Begin VB.TextBox Text2 
         Height          =   350
         Index           =   1
         Left            =   2880
         TabIndex        =   96
         Top             =   1800
         Width           =   750
      End
      Begin VB.TextBox Text2 
         Height          =   350
         Index           =   2
         Left            =   3840
         TabIndex        =   95
         Top             =   1800
         Width           =   750
      End
      Begin VB.TextBox Text2 
         Height          =   350
         Index           =   3
         Left            =   4800
         TabIndex        =   94
         Top             =   1800
         Width           =   750
      End
      Begin VB.TextBox Text2 
         Height          =   350
         Index           =   4
         Left            =   5760
         TabIndex        =   93
         Top             =   1800
         Width           =   750
      End
      Begin VB.TextBox Text2 
         Height          =   350
         Index           =   5
         Left            =   6720
         TabIndex        =   92
         Top             =   1800
         Width           =   750
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFF80&
         Height          =   350
         Index           =   6
         Left            =   8040
         TabIndex        =   91
         Top             =   1800
         Width           =   750
      End
      Begin VB.TextBox Text3 
         Height          =   350
         Index           =   0
         Left            =   1920
         TabIndex        =   90
         Text            =   "10.58"
         Top             =   2400
         Width           =   750
      End
      Begin VB.TextBox Text3 
         Height          =   350
         Index           =   1
         Left            =   2880
         TabIndex        =   89
         Text            =   "10.62"
         Top             =   2400
         Width           =   750
      End
      Begin VB.TextBox Text3 
         Height          =   350
         Index           =   2
         Left            =   3840
         TabIndex        =   88
         Text            =   "10.58"
         Top             =   2400
         Width           =   750
      End
      Begin VB.TextBox Text3 
         Height          =   350
         Index           =   3
         Left            =   4800
         TabIndex        =   87
         Text            =   "10.62"
         Top             =   2400
         Width           =   750
      End
      Begin VB.TextBox Text3 
         Height          =   350
         Index           =   4
         Left            =   5760
         TabIndex        =   86
         Text            =   "10.60"
         Top             =   2400
         Width           =   750
      End
      Begin VB.TextBox Text3 
         Height          =   350
         Index           =   5
         Left            =   6720
         TabIndex        =   85
         Text            =   "10.59"
         Top             =   2400
         Width           =   750
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFF00&
         Height          =   350
         Index           =   6
         Left            =   8040
         TabIndex        =   84
         Top             =   2400
         Width           =   750
      End
      Begin VB.TextBox Text4 
         Height          =   300
         Index           =   0
         Left            =   1920
         TabIndex        =   83
         Text            =   "16.12"
         Top             =   3000
         Width           =   750
      End
      Begin VB.TextBox Text4 
         Height          =   350
         Index           =   1
         Left            =   2880
         TabIndex        =   82
         Text            =   "16.14"
         Top             =   3000
         Width           =   750
      End
      Begin VB.TextBox Text4 
         Height          =   350
         Index           =   2
         Left            =   3840
         TabIndex        =   81
         Text            =   "16.20"
         Top             =   3000
         Width           =   750
      End
      Begin VB.TextBox Text4 
         Height          =   350
         Index           =   3
         Left            =   4800
         TabIndex        =   80
         Text            =   "16.18"
         Top             =   3000
         Width           =   750
      End
      Begin VB.TextBox Text4 
         Height          =   350
         Index           =   4
         Left            =   5760
         TabIndex        =   79
         Text            =   "16.00"
         Top             =   3000
         Width           =   750
      End
      Begin VB.TextBox Text4 
         Height          =   350
         Index           =   5
         Left            =   6720
         TabIndex        =   78
         Text            =   "16.00"
         Top             =   3000
         Width           =   750
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFF00&
         Height          =   350
         Index           =   6
         Left            =   8040
         TabIndex        =   77
         Top             =   3000
         Width           =   750
      End
      Begin VB.TextBox Text5 
         Height          =   350
         Index           =   0
         Left            =   1920
         TabIndex        =   76
         Text            =   "48.35"
         Top             =   3600
         Width           =   750
      End
      Begin VB.TextBox Text5 
         Height          =   350
         Index           =   1
         Left            =   2880
         TabIndex        =   75
         Text            =   "48.33"
         Top             =   3600
         Width           =   750
      End
      Begin VB.TextBox Text5 
         Height          =   350
         Index           =   2
         Left            =   3840
         TabIndex        =   74
         Text            =   "48.35"
         Top             =   3600
         Width           =   750
      End
      Begin VB.TextBox Text5 
         Height          =   350
         Index           =   3
         Left            =   4800
         TabIndex        =   73
         Text            =   "48.32"
         Top             =   3600
         Width           =   750
      End
      Begin VB.TextBox Text5 
         Height          =   350
         Index           =   4
         Left            =   5760
         TabIndex        =   72
         Text            =   "48.32"
         Top             =   3600
         Width           =   750
      End
      Begin VB.TextBox Text5 
         Height          =   350
         Index           =   5
         Left            =   6720
         TabIndex        =   71
         Text            =   "48.35"
         Top             =   3600
         Width           =   750
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFFF80&
         Height          =   350
         Index           =   6
         Left            =   8040
         TabIndex        =   70
         Top             =   3600
         Width           =   750
      End
      Begin VB.TextBox Text6 
         Height          =   350
         Index           =   0
         Left            =   1920
         TabIndex        =   69
         Text            =   "22.02"
         Top             =   4200
         Width           =   750
      End
      Begin VB.TextBox Text6 
         Height          =   350
         Index           =   1
         Left            =   2880
         TabIndex        =   68
         Text            =   "22.04"
         Top             =   4200
         Width           =   750
      End
      Begin VB.TextBox Text6 
         Height          =   350
         Index           =   2
         Left            =   3840
         TabIndex        =   67
         Text            =   "22.04"
         Top             =   4200
         Width           =   750
      End
      Begin VB.TextBox Text6 
         Height          =   350
         Index           =   3
         Left            =   4800
         TabIndex        =   66
         Text            =   "22.02"
         Top             =   4200
         Width           =   750
      End
      Begin VB.TextBox Text6 
         Height          =   350
         Index           =   4
         Left            =   5760
         TabIndex        =   65
         Text            =   "22.02"
         Top             =   4200
         Width           =   750
      End
      Begin VB.TextBox Text6 
         Height          =   350
         Index           =   5
         Left            =   6720
         TabIndex        =   64
         Text            =   "22.02"
         Top             =   4200
         Width           =   750
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FFFFC0&
         Height          =   350
         Index           =   6
         Left            =   8040
         TabIndex        =   63
         Top             =   4200
         Width           =   750
      End
      Begin VB.TextBox Text7 
         Height          =   500
         Left            =   1440
         TabIndex        =   62
         Top             =   5760
         Width           =   1200
      End
      Begin VB.TextBox Text8 
         Height          =   500
         Left            =   3960
         TabIndex        =   61
         Top             =   5760
         Width           =   1200
      End
      Begin VB.TextBox Text9 
         Height          =   500
         Left            =   6360
         TabIndex        =   60
         Top             =   5760
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "点击计算平均值"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   240
         TabIndex        =   59
         Top             =   6600
         Width           =   1200
      End
      Begin VB.CommandButton Command2 
         Caption         =   "点击计算周期T"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1920
         TabIndex        =   58
         Top             =   6600
         Width           =   1200
      End
      Begin VB.CommandButton Command3 
         Caption         =   "点击计算上孔a"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3720
         TabIndex        =   57
         Top             =   6600
         Width           =   1200
      End
      Begin VB.CommandButton Command4 
         Caption         =   "点击计算下孔b"
         Height          =   600
         Left            =   5520
         TabIndex        =   56
         Top             =   6600
         Width           =   1200
      End
      Begin VB.CommandButton Command5 
         Caption         =   "点击计算悬线长l"
         Height          =   600
         Left            =   7200
         TabIndex        =   55
         Top             =   6600
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Index           =   5
         Left            =   6720
         TabIndex        =   54
         Text            =   "25.330"
         Top             =   1320
         Width           =   750
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Index           =   4
         Left            =   5760
         TabIndex        =   53
         Text            =   "25.327"
         Top             =   1320
         Width           =   750
      End
      Begin VB.CommandButton Command6 
         Caption         =   "点击计算圆柱体半价r"
         Height          =   600
         Left            =   9120
         TabIndex        =   52
         Top             =   6600
         Width           =   1200
      End
      Begin VB.TextBox Text10 
         Height          =   350
         Index           =   0
         Left            =   1920
         TabIndex        =   51
         Top             =   4800
         Width           =   750
      End
      Begin VB.TextBox Text10 
         Height          =   350
         Index           =   1
         Left            =   2880
         TabIndex        =   50
         Top             =   4800
         Width           =   750
      End
      Begin VB.TextBox Text10 
         Height          =   350
         Index           =   2
         Left            =   3840
         TabIndex        =   49
         Top             =   4800
         Width           =   750
      End
      Begin VB.TextBox Text10 
         Height          =   350
         Index           =   3
         Left            =   4800
         TabIndex        =   48
         Top             =   4800
         Width           =   750
      End
      Begin VB.TextBox Text10 
         Height          =   350
         Index           =   4
         Left            =   5760
         TabIndex        =   47
         Top             =   4800
         Width           =   750
      End
      Begin VB.TextBox Text10 
         Height          =   350
         Index           =   5
         Left            =   6720
         TabIndex        =   46
         Top             =   4800
         Width           =   750
      End
      Begin VB.TextBox Text10 
         Height          =   350
         Index           =   6
         Left            =   8040
         TabIndex        =   45
         Top             =   4800
         Width           =   750
      End
      Begin VB.CommandButton Command7 
         Caption         =   "点击计算间接量H"
         Height          =   750
         Left            =   -72240
         TabIndex        =   44
         Top             =   6000
         Width           =   1500
      End
      Begin VB.TextBox Text11 
         Height          =   700
         Left            =   -72960
         TabIndex        =   43
         Top             =   3000
         Width           =   1200
      End
      Begin VB.TextBox Text12 
         Height          =   700
         Left            =   -69120
         TabIndex        =   42
         Top             =   3000
         Width           =   1200
      End
      Begin VB.TextBox Text13 
         Height          =   700
         Left            =   -65760
         TabIndex        =   41
         Top             =   3000
         Width           =   1200
      End
      Begin VB.TextBox Text14 
         Height          =   700
         Left            =   -72960
         TabIndex        =   40
         Top             =   4800
         Width           =   1200
      End
      Begin VB.TextBox Text15 
         Height          =   700
         Left            =   -69000
         TabIndex        =   39
         Top             =   4680
         Width           =   1200
      End
      Begin VB.TextBox Text16 
         Height          =   500
         Left            =   9000
         TabIndex        =   38
         Top             =   5760
         Width           =   1200
      End
      Begin VB.TextBox Text17 
         Height          =   700
         Left            =   -65760
         TabIndex        =   37
         Top             =   4680
         Width           =   1200
      End
      Begin VB.CommandButton Command8 
         Caption         =   "点击计算转动惯量I"
         Height          =   750
         Left            =   -68640
         TabIndex        =   36
         Top             =   6000
         Width           =   1500
      End
      Begin VB.TextBox Text18 
         Height          =   400
         Left            =   -69840
         TabIndex        =   35
         Text            =   "1234"
         Top             =   1200
         Width           =   1000
      End
      Begin VB.TextBox Text19 
         Height          =   500
         Index           =   0
         Left            =   -72120
         TabIndex        =   34
         Text            =   "6.95"
         Top             =   1800
         Width           =   750
      End
      Begin VB.TextBox Text19 
         Height          =   500
         Index           =   1
         Left            =   -70920
         TabIndex        =   33
         Text            =   "8.98"
         Top             =   1800
         Width           =   750
      End
      Begin VB.TextBox Text19 
         Height          =   500
         Index           =   2
         Left            =   -69720
         TabIndex        =   32
         Text            =   "10.95"
         Top             =   1800
         Width           =   750
      End
      Begin VB.TextBox Text19 
         Height          =   500
         Index           =   3
         Left            =   -68520
         TabIndex        =   31
         Text            =   "12.92"
         Top             =   1800
         Width           =   750
      End
      Begin VB.TextBox Text19 
         Height          =   500
         Index           =   4
         Left            =   -67320
         TabIndex        =   30
         Text            =   "14.94"
         Top             =   1800
         Width           =   750
      End
      Begin VB.TextBox Text19 
         Height          =   500
         Index           =   5
         Left            =   -66120
         TabIndex        =   29
         Text            =   "16.95"
         Top             =   1800
         Width           =   750
      End
      Begin VB.TextBox Text20 
         Height          =   500
         Index           =   0
         Left            =   -72120
         TabIndex        =   28
         Text            =   "1.199"
         Top             =   2640
         Width           =   750
      End
      Begin VB.TextBox Text20 
         Height          =   500
         Index           =   1
         Left            =   -70920
         TabIndex        =   27
         Text            =   "1.211"
         Top             =   2640
         Width           =   750
      End
      Begin VB.TextBox Text20 
         Height          =   500
         Index           =   2
         Left            =   -69720
         TabIndex        =   26
         Text            =   "1.243"
         Top             =   2640
         Width           =   750
      End
      Begin VB.TextBox Text20 
         Height          =   500
         Index           =   3
         Left            =   -68520
         TabIndex        =   25
         Text            =   "1.275"
         Top             =   2640
         Width           =   750
      End
      Begin VB.TextBox Text20 
         Height          =   500
         Index           =   4
         Left            =   -67320
         TabIndex        =   24
         Text            =   "1.306"
         Top             =   2640
         Width           =   750
      End
      Begin VB.TextBox Text20 
         Height          =   500
         Index           =   5
         Left            =   -66120
         TabIndex        =   23
         Text            =   "1.329"
         Top             =   2640
         Width           =   750
      End
      Begin VB.TextBox Text21 
         Height          =   500
         Index           =   1
         Left            =   -70920
         TabIndex        =   22
         Top             =   3480
         Width           =   750
      End
      Begin VB.TextBox Text21 
         Height          =   500
         Index           =   2
         Left            =   -69720
         TabIndex        =   21
         Top             =   3480
         Width           =   750
      End
      Begin VB.TextBox Text21 
         Height          =   500
         Index           =   3
         Left            =   -68520
         TabIndex        =   20
         Top             =   3480
         Width           =   750
      End
      Begin VB.TextBox Text21 
         Height          =   500
         Index           =   4
         Left            =   -67320
         TabIndex        =   19
         Top             =   3480
         Width           =   750
      End
      Begin VB.TextBox Text21 
         Height          =   500
         Index           =   5
         Left            =   -66120
         TabIndex        =   18
         Top             =   3480
         Width           =   750
      End
      Begin VB.TextBox Text22 
         Height          =   500
         Index           =   0
         Left            =   -72120
         TabIndex        =   17
         Top             =   4320
         Width           =   750
      End
      Begin VB.TextBox Text22 
         Height          =   500
         Index           =   1
         Left            =   -70920
         TabIndex        =   16
         Top             =   4320
         Width           =   750
      End
      Begin VB.TextBox Text22 
         Height          =   500
         Index           =   2
         Left            =   -69720
         TabIndex        =   15
         Top             =   4320
         Width           =   750
      End
      Begin VB.TextBox Text22 
         Height          =   500
         Index           =   3
         Left            =   -68520
         TabIndex        =   14
         Top             =   4320
         Width           =   750
      End
      Begin VB.TextBox Text22 
         Height          =   500
         Index           =   4
         Left            =   -67320
         TabIndex        =   13
         Top             =   4320
         Width           =   750
      End
      Begin VB.TextBox Text22 
         Height          =   500
         Index           =   5
         Left            =   -66120
         TabIndex        =   12
         Top             =   4320
         Width           =   750
      End
      Begin VB.TextBox Text23 
         Height          =   500
         Index           =   0
         Left            =   -72120
         TabIndex        =   11
         Top             =   5160
         Width           =   750
      End
      Begin VB.TextBox Text23 
         Height          =   500
         Index           =   1
         Left            =   -70920
         TabIndex        =   10
         Top             =   5160
         Width           =   750
      End
      Begin VB.TextBox Text23 
         Height          =   500
         Index           =   2
         Left            =   -69720
         TabIndex        =   9
         Top             =   5160
         Width           =   750
      End
      Begin VB.TextBox Text23 
         Height          =   500
         Index           =   3
         Left            =   -68520
         TabIndex        =   8
         Top             =   5160
         Width           =   750
      End
      Begin VB.TextBox Text23 
         Height          =   500
         Index           =   4
         Left            =   -67320
         TabIndex        =   7
         Top             =   5160
         Width           =   750
      End
      Begin VB.TextBox Text23 
         Height          =   500
         Index           =   5
         Left            =   -66120
         TabIndex        =   6
         Top             =   5160
         Width           =   750
      End
      Begin VB.CommandButton Command9 
         Caption         =   "点击计算转动惯量I的实验值"
         Height          =   600
         Left            =   -74400
         TabIndex        =   5
         Top             =   6360
         Width           =   1600
      End
      Begin VB.CommandButton Command10 
         Caption         =   "点击计算转动惯量I的公式值"
         Height          =   615
         Left            =   -71040
         TabIndex        =   4
         Top             =   6360
         Width           =   1600
      End
      Begin VB.CommandButton Command11 
         Caption         =   "点击计算转动惯量I的百分误差"
         Height          =   600
         Left            =   -67920
         TabIndex        =   3
         Top             =   6360
         Width           =   1600
      End
      Begin VB.TextBox Text24 
         Height          =   500
         Left            =   -69480
         TabIndex        =   2
         Text            =   "118"
         Top             =   840
         Width           =   2000
      End
      Begin VB.TextBox Text21 
         Height          =   500
         Index           =   0
         Left            =   -72120
         TabIndex        =   1
         Top             =   3480
         Width           =   750
      End
      Begin VB.Image Image1 
         Height          =   9090
         Index           =   2
         Left            =   -75000
         Picture         =   "Form2.frx":0054
         Top             =   360
         Width           =   14535
      End
      Begin VB.Image Image1 
         Height          =   9075
         Index           =   1
         Left            =   -75120
         Picture         =   "Form2.frx":16A04
         Top             =   360
         Width           =   13530
      End
      Begin VB.Image Image1 
         Height          =   10005
         Index           =   0
         Left            =   0
         Picture         =   "Form2.frx":32212
         Top             =   360
         Width           =   15000
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "周期T/s"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   127
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "上孔a/cm"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   126
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "下孔b/cm"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   125
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "悬线长l/cm"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   480
         TabIndex        =   124
         Top             =   3600
         Width           =   1365
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "圆柱体直径2r/cm"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   5
         Left            =   360
         TabIndex        =   123
         Top             =   4080
         Width           =   1530
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A类不确定度"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   360
         TabIndex        =   122
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "B类不确定度"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   2880
         TabIndex        =   121
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "总不确定度"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   5400
         TabIndex        =   120
         Top             =   5880
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "平均值"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   8040
         TabIndex        =   119
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "圆柱体半径r/cm"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   7
         Left            =   360
         TabIndex        =   118
         Top             =   4800
         Width           =   1530
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "相对不确定度"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   7800
         TabIndex        =   117
         Top             =   5880
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入下圆盘质量m/g"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   660
         Index           =   8
         Left            =   -72600
         TabIndex        =   116
         Top             =   1200
         Width           =   2370
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "间接测量量H的平均值/cm"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Index           =   9
         Left            =   -74760
         TabIndex        =   115
         Top             =   3120
         Width           =   1650
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "间接测量量H的总不确定度/cm"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   11
         Left            =   -71040
         TabIndex        =   114
         Top             =   3120
         Width           =   1770
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转动惯量I*E-4/kg*m^2"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   12
         Left            =   -74880
         TabIndex        =   113
         Top             =   4800
         Width           =   1770
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转动惯量I的总不确定度*E-4/kg*m^2"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   13
         Left            =   -71160
         TabIndex        =   112
         Top             =   4680
         Width           =   2130
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转动惯量I的相对不确定度"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   14
         Left            =   -67560
         TabIndex        =   111
         Top             =   4680
         Width           =   1770
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "间接测量量H的相对不确定度"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   15
         Left            =   -67560
         TabIndex        =   110
         Top             =   3120
         Width           =   1770
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "孔距2d/cm"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   -74400
         TabIndex        =   109
         Top             =   1800
         Width           =   1530
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "周期T2/s"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   -74280
         TabIndex        =   108
         Top             =   2640
         Width           =   1530
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转动惯量I的实验值*E-4/kg*m^2"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   17
         Left            =   -74640
         TabIndex        =   107
         Top             =   3360
         Width           =   2010
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转动惯量I的公式值*E-4/kg*m^2"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   18
         Left            =   -74640
         TabIndex        =   106
         Top             =   4200
         Width           =   2010
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转动惯量I的百分误差"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   19
         Left            =   -74640
         TabIndex        =   105
         Top             =   5160
         Width           =   2010
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入单个圆柱体质量m/g"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   330
         Index           =   20
         Left            =   -72600
         TabIndex        =   104
         Top             =   840
         Width           =   3210
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "时间t/s"
         BeginProperty Font 
            Name            =   "造字工房尚黑 G0v1 粗体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   103
         Top             =   1320
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Double, B As Double, L As Double, aa As Double

Private Sub command1_click()

 Dim A As Double, B As Double, c As Double, d As Double, e As Double, f As Double, m As Double

 For i = 0 To 5

  Text2(i) = Format(Val(Text1(i).Text) / 20, "0.000") '求出六个周期时间

  Text10(i) = Format(Val(Text6(i).Text) / 2, "0.00")

  A = A + Val(Text2(i).Text)

  B = B + Val(Text1(i).Text)

  c = c + Val(Text3(i).Text)

  d = d + Val(Text4(i).Text)

  e = e + Val(Text5(i).Text)

  f = f + Val(Text6(i).Text)

  m = m + Val(Text10(i).Text)

 Next i

 Text1(6) = Round(B / 6, 3)  '计算平均值

 Text2(6) = Round(A / 6, 3)

 Text3(6) = Format(c / 6, "0.00")

 Text4(6) = Format(d / 6, "0.00")

 Text5(6) = Format(e / 6, "0.00")

 Text6(6) = Format(f / 6, "0.00")

 Text10(6) = Format(m / 6, "0.00")

End Sub



Private Sub command2_click()

 Dim g As Double, T As Double

 For i = 0 To 5

  g = g + (Val(Text2(i)) - Text2(6)) ^ 2 '计算周期T的A类不确定度1

 Next i

 Text7.Text = Format(Sqr(g / 5), "0.000") '计算周期T的A类不确定度2

 Text8.Text = Format(0.002 / Sqr(3), "0.000") '计算周期T的B类不确定度

 T = Format(Sqr(Val(Text7.Text) ^ 2 + Val(Text8.Text) ^ 2), "0.000") '计算周期T的总不确定度
 
 Text9.Text = T

 Text16 = Format(T * 100 / Text2(6), "0.00") & "%" '计算周期T的相对不确定度

End Sub

Private Sub Command3_Click()

 Dim h As Double

 For i = 0 To 5

  h = h + (Val(Text3(i)) - Text3(6)) ^ 2 '计算上孔a的A类不确定度1

 Next i

 Text7.Text = Format(Sqr(h / 5), "0.00") '计算上孔a的A类不确定度2

 Text8.Text = Format(0.02 / Sqr(3), "0.00") '计算上孔a的B类不确定度

 A = Format(Sqr(Val(Text7.Text) ^ 2 + Val(Text8.Text) ^ 2), "0.00") '计算上孔a的总不确定度

 Text9.Text = A

 Text16 = Format(A * 100 / Text3(6), "0.00") & "%" '计算上孔a的相对不确定度

End Sub

Private Sub Command4_Click()

 Dim j As Double

 For i = 0 To 5

  j = j + (Val(Text4(i)) - Text4(6)) ^ 2 '计算下孔b的A类不确定度1

 Next i
 
 Text7.Text = Format(Sqr(j / 5), "0.00") '计算下孔b的A类不确定度2

 Text8.Text = Format(0.02 / Sqr(3), "0.00") '计算下孔b的B类不确定度

 B = Format(Sqr(Val(Text7.Text) ^ 2 + Val(Text8.Text) ^ 2), "0.00") '计算下孔b的总不确定度

 Text9.Text = B

 Text16 = Format(B * 100 / Text4(6), "0.00") & "%" '计算下孔b的相对不确定度

End Sub

Private Sub Command5_Click()

 Dim k As Double

 For i = 0 To 5

  k = k + (Val(Text5(i)) - Text5(6)) ^ 2 '计算悬线长的A类不确定度1

 Next i

 Text7.Text = Format(Sqr(k / 5), "0.00") '计算悬线长的A类不确定度2

 Text8.Text = Format(0.02 / Sqr(3), "0.00") '计算悬线长的B类不确定度

 L = Format(Sqr(Val(Text7.Text) ^ 2 + Val(Text8.Text) ^ 2), "0.00") '计算悬线长的总不确定度

 Text9.Text = L

 Text16 = Format(L * 100 / Text5(6), "0.00") & "%" '计算悬线长的相对不确定度

End Sub

Private Sub Command6_Click()

 Dim z As Double, r As Double

 For i = 0 To 5

  z = z + (Val(Text10(i)) - Text10(6)) ^ 2  '计算圆柱体半径的A类不确定度1

 Next i

 Text7.Text = Format(Sqr(z / 5), "0.00") '计算圆柱体半径的A类不确定度1

 Text8.Text = Format(0.02 / Sqr(3), "0.00") '计算圆柱体半径的B类不确定度

 r = Format(Sqr(Val(Text7.Text) ^ 2 + Val(Text8.Text) ^ 2), "0.00") '计算圆柱体半径的总不确定度

 Text9.Text = r

 Text16 = Format(r * 100 / Text10(6), "0.00") & "%" '计算圆柱体半径的相对不确定度

End Sub

Private Sub Command7_Click()

 Dim n As Double, o As Double, p As Double, q As Double, r As Double

 n = Format(Sqr(Text5(6) ^ 2 - (1 / 3) * (Text4(6) - Text3(6)) ^ 2), "0.00")   '计算间接测量量H

 o = Format(Sqr(((Text4(6) - Text3(6)) ^ 2 * (A ^ 2 + B ^ 2) + 9 * Text5(6) ^ 2 * L * L) / (9 * Text5(6) ^ 2 - 3 * (Text4(6) - Text3(6)) ^ 2)), "0.00") '根据不确定度传递公式测量H的总不确定度

 Text11 = n

 Text12 = o

 Text13 = Format(o * 100 / Text11, "0.00") & "%" '测量间接测量量H的相对不确定度

End Sub

Private Sub Command8_Click()

 Dim bb As Double

 aa = Format((Text18 * 9.8 * Text2(6) * Text2(6) * Text3(6) * Text4(6) * 0.1) / (12 * 3.14 * 3.14 * Text11), "0.00") '计算转动惯量I

 Text14 = aa

 bb = Sqr((A ^ 2 / Text3(6) ^ 2) + (B ^ 2 / Text4(6) ^ 2) + (4 * T * T / Text2(6) ^ 2) - (Text12 ^ 2 / Text11 ^ 2)) '根据相对不确定传递公式计算转动惯量I的相对不确定度

 Text17 = Format(bb * 100, "0.00") & "%"

 Text15 = Format(aa * bb, "0.00")  '计算转动惯量I的总不确定度
End Sub

Private Sub Command9_Click()

 For i = 0 To 5

  Text21(i) = Format(((((Text18 + 2 * Text24) * 9.8 * Text3(6) * Text4(6) * Text20(i) ^ 2) * 0.1 / (12 * 3.14 * 3.14 * Text11)) - aa) * 0.5, "0.00") '计算转动惯量I的实验值

 Next i

End Sub

Private Sub Command10_Click()

 For i = 0 To 5

  Text22(i) = Format((0.5 * Text24 * Text10(6) ^ 2 * 0.01 + Text24 * ((Text19(i) / 2) ^ 2)) * 0.001, "0.00") '计算转动惯量I的公式值

 Next i

End Sub

Private Sub Command11_Click()

 For i = 0 To 5

  Text23(i) = Format((Abs(Text21(i) - Text22(i)) * 100 / Text22(i)), "0") & "%" '计算百分误差

 Next i

End Sub



