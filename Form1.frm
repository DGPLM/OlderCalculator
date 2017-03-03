VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "三线摆测刚体的转动惯量实验数据处理计算器"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8205
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6105
   ScaleWidth      =   8205
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "造字工房尚黑 G0v1 粗体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   4200
      TabIndex        =   3
      Top             =   4560
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登录"
      BeginProperty Font 
         Name            =   "造字工房尚黑 G0v1 粗体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   1920
      Picture         =   "Form1.frx":5937F
      TabIndex        =   2
      Top             =   4560
      Width           =   1500
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   2640
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "造字工房尚黑 G0v1 粗体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "账号"
      BeginProperty Font 
         Name            =   "造字工房尚黑 G0v1 粗体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public c As Integer
Private Sub command1_click()
If Text1 = "abc" Then
If Text2 = "123" Then
Form2.Show
Unload Me
Else
If c > 1 Then
c = c - 1
MsgBox "用户名或密码错误！还有" & c & "次机会", vbExclamation, "错误"
Text2 = ""
Text2.SetFocus
Else
End
End If
End If
Else
If c > 1 Then
c = c - 1
MsgBox "用户名或密码错误！还有" & c & "次机会", vbExclamation, "错误"
Text2 = ""
Text2.SetFocus
Else
End
End If
End If
End Sub
Private Sub command2_click()
End
End Sub
Private Sub form_load()
c = 3
End Sub
