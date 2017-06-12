VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   -1710
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   13755
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   9600
      TabIndex        =   18
      Top             =   4200
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   8880
      Top             =   5160
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   1300
      Left            =   0
      ScaleHeight     =   1245
      ScaleWidth      =   7935
      TabIndex        =   2
      Top             =   2280
      Width           =   8000
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   1300
      Left            =   0
      ScaleHeight     =   1245
      ScaleWidth      =   7935
      TabIndex        =   1
      Top             =   480
      Width           =   8000
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   8000
      Left            =   0
      ScaleHeight     =   7935
      ScaleWidth      =   7935
      TabIndex        =   0
      Top             =   4080
      Width           =   8000
   End
   Begin VB.Label Label11 
      Caption         =   "π"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   17
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "π"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   16
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Hz/s"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   15
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Hz/s"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   14
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "相位"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   13
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "相位"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   12
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "频率"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   11
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "频率"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "合成波形"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Y轴波形"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "X轴波形"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x_1, x_11, x_2, x_22 As Double
Dim y_1, y_11, y_2, y_22 As Double
Dim f_1, f_2 As Integer
Dim h_1, h_2 As Double

Private Sub Command1_Click()
f_1 = Text1.Text
f_2 = Text3.Text
h_1 = Text2.Text
h_2 = Text4.Text
x_1 = 0
x_2 = 0
x_11 = 0
x_22 = 0
Picture2.Cls
Picture3.Cls
Picture1.Cls
Picture2.Line (0, 650)-(8000, 650), vbRed
Picture3.Line (0, 650)-(8000, 650), vbRed
End Sub

Private Sub Form_Load()
x_1 = 0
x_2 = 0
x_11 = 0
x_22 = 0

f_1 = 10
f_2 = 10
h_1 = 0
h_2 = 0
Text1.Text = f_1
Text2.Text = h_1
Text3.Text = f_2
Text4.Text = h_2

Picture2.AutoRedraw = True
Picture2.Line (0, 650)-(8000, 650), vbRed

Picture3.AutoRedraw = True
Picture3.Line (0, 650)-(8000, 650), vbRed

Timer1.Interval = 1
Timer1.Enabled = True
End Sub
'定义函数并画线
Private Sub Timer1_Timer()
y_1 = 650 - Cos(x_1 * 2 * f_1 * 3.1415 + h_1 * 3.1415) * 500
y_11 = 650 - Cos(x_11 * 2 * f_1 * 3.1415 + h_1 * 3.1415) * 500
y_22 = 650 - Cos(x_22 * 2 * f_2 * 3.1415 + h_2 * 3.1415) * 500
y_2 = 650 - Cos(x_2 * 2 * f_2 * 3.1415 + h_2 * 3.1415) * 500
Picture2.Line (x_11, y_11)-(x_1, y_1), vbGreen
Picture3.Line (x_22, y_22)-(x_2, y_2), vbBlue
Picture1.Line (5 * y_22, 5 * y_11)-(5 * y_2, 5 * y_1), vbRed
x_22 = x_2
x_2 = x_2 + 10

x_11 = x_1
x_1 = x_1 + 10
If x_1 >= 6000 Then
x_1 = x_1 - 6000
x_11 = x_1
Picture2.Cls
Picture2.Line (0, 650)-(8000, 650), vbRed

End If
If x_2 >= 6000 Then
x_2 = x_2 - 6000
x_22 = x_2
Picture3.Cls
Picture3.Line (0, 650)-(8000, 650), vbRed

End If
End Sub
