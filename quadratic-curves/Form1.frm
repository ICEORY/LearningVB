VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   9765
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "y=a*x^2+b*x+c"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   6960
      TabIndex        =   2
      Top             =   600
      Width           =   2655
      Begin VB.CommandButton Command1 
         Caption         =   "确定"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "c"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "b"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6500
      Left            =   120
      ScaleHeight     =   6435
      ScaleWidth      =   6435
      TabIndex        =   1
      Top             =   600
      Width           =   6500
   End
   Begin VB.Label Label1 
      Caption         =   "曲线显示"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x, xt, y, yt As Double
Dim a, b, c As Double
Public Sub draw()
Picture1.Line (50, 3250)-(6450, 3250)
Picture1.Line (3250, 50)-(3250, 6450)
Picture1.Line (6450, 3250)-(6400, 3200)
Picture1.Line (6450, 3250)-(6400, 3300)
Picture1.Line (3250, 50)-(3200, 100)
Picture1.Line (3250, 50)-(3300, 100)
x = -3250
xt = x
For x = -325 To 325
y = a * x * x + b * x + c
yt = a * xt * xt + b * xt + c
y = y
yt = yt
If xt * 10 + 3250 <= 6500 And xt * 10 + 3250 >= 0 And x * 10 + 3250 <= 6500 And x * 10 + 3250 >= 0 And 3250 - yt * 10 <= 6500 And 3250 - yt * 10 >= 0 And 3250 - y * 10 <= 6500 And 3250 - y * 10 >= 0 Then
Picture1.Line (xt * 10 + 3250, 3250 - yt * 10)-(x * 10 + 3250, 3250 - y * 10)
End If
xt = x
Next
End Sub
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox "请输入参数a,b,c"
Else
Picture1.Cls
a = Text1.Text
b = Text2.Text
c = Text3.Text
Call draw
End If
End Sub

Private Sub Form_Load()
Form1.Height = 8000
Form1.Width = 10000
Picture1.Height = 6500
Picture1.Width = 6600
Picture1.AutoRedraw = True
Picture1.Line (50, 3250)-(6450, 3250)
Picture1.Line (3250, 50)-(3250, 6450)
Picture1.Line (6450, 3250)-(6400, 3200)
Picture1.Line (6450, 3250)-(6400, 3300)
Picture1.Line (3250, 50)-(3200, 100)
Picture1.Line (3250, 50)-(3300, 100)
End Sub
