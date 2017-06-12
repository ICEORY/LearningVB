VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   14500
      Left            =   120
      ScaleHeight     =   14445
      ScaleMode       =   0  'User
      ScaleWidth      =   14500
      TabIndex        =   0
      Top             =   0
      Width           =   14500
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   2  'Dash
         Height          =   400
         Left            =   1320
         Top             =   2880
         Visible         =   0   'False
         Width           =   398
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'定义两个全局变量
Dim coun As Integer '用于表示下棋的总步数
Dim pill(0 To 18, 0 To 18) As Integer '用于储存下棋的坐标，例如pill(0,1)=1表示棋子落在第一行第二列位置，且棋子为黑子
'若pill(0,1)=2,表示下了白子，若值为零则该位置为空
Private Sub Form_Load() '棋盘初始化，显示19条横线以及19条竖线并在相应位置绘制五个黑点
Dim i, j As Integer '定义两个临时变量，用于计数用
coun = 0 '步数归零
Pic1.AutoRedraw = True '开启picture控件的自动重绘功能，使图线能正确绘制
Form1.Height = 10500 '设置窗体大小，高为10500
Form1.Width = 11500 '宽为10500
Form1.Caption = "五子棋" '设置窗体名称

For i = 0 To 18 '将储存棋子位置的数组归零
    For j = 0 To 18
    pill(i, j) = 0
    Next
Next

For i = 0 To 18 '绘制棋盘上的黑线
Pic1.Line (500, 500 + 500 * i)-(9500, 500 + 500 * i)
Pic1.Line (500 + 500 * i, 500)-(500 + i * 500, 9500)
Next

'以下五个语句调用绘制圆的过程，在棋盘上相应位置绘制小黑点
Call draw(3, 3, 0)
Call draw(9, 9, 0)
Call draw(15, 15, 0)
Call draw(3, 15, 0)
Call draw(15, 3, 0)
End Sub


Public Sub draw(ByVal x As Integer, ByVal y As Integer, ByVal n As Integer) '绘制圆，若绘制小圆则作为标志点，绘制大圆为棋子
Dim r As Integer '定义半径
'n用于选择绘制圆的类型
'n=0表示绘制标志点
If n = 0 Then
Pic1.FillColor = vbBlack
Pic1.FillStyle = 0
r = 50
End If
'绘制黑棋
If n = 1 Then
Pic1.FillColor = vbBlack
Pic1.FillStyle = 0
r = 200
End If
'绘制白棋
If n = 2 Then
Pic1.FillColor = vbWhite
Pic1.FillStyle = 0
r = 200
End If

Pic1.Circle (500 * x + 500, 500 * y + 500), r
End Sub
Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) '获取鼠标按下的位置
Dim qx, qy As Integer '定义两个变量用于表示棋子的坐标

For qx = 0 To 17 '将鼠标位置与棋盘上位置对应
    If x < 500 Then '当鼠标不在棋盘上的情况，将qx置零
    qx = 0
    Exit For
    End If
    '将鼠标横坐标与棋盘坐标在一定的范围内对应
    If x > 500 * (qx + 1) - 250 And x < 500 * (qx + 1) + 250 Then
    Exit For
    End If
Next
'将鼠标的纵坐标与棋盘坐标对应，方法同上
For qy = 0 To 17
    If y < 500 Then
    qy = 0
    Exit For
    End If
    
    If y > 500 * (qy + 1) - 250 And y < 500 * (qy + 1) + 250 Then
    Exit For
    End If
Next
'判断是否该位置是否已经有落子了，有的话跳到结束，等待下次鼠标的点击
If pill(qx, qy) <> 0 Then
GoTo end_1
End If
'如果该位置没有落子，则在该位置绘制棋子
Call draw(qx, qy, (coun Mod 2) + 1)
pill(qx, qy) = coun Mod 2 + 1
'判断输赢，只需要判断本次落子的周围是否连成五颗，根据递推原理可以知道该方法的可行性
If win(qx, qy) = 0 Then
coun = coun + 1
End If
'对输赢情况的结果执行相应的操作
If win(qx, qy) = 1 Then
MsgBox "黑子赢了"
Pic1.Cls
Call Form_Load
End If

If win(qx, qy) = 2 Then
MsgBox "白子赢了"
Pic1.Cls
Call Form_Load
End If

end_1: '该位置对应之前的goto 语句
End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) '获取鼠标移动过程中的坐标
Dim qx, qy As Integer '定义两个变量储存坐标
'该子过程主要是显示鼠标在棋盘上对于的位置，并显示光标，可以避免落子位置判断失误的情况
'以下的原理与获取棋子坐标的方法一样，如果上面的操作原理懂的话这个也懂了
For qx = 0 To 17
    If x < 500 Then
    qx = 0
    Exit For
    End If
    If x > 500 * (qx + 1) - 250 And x < 500 * (qx + 1) + 250 Then
    Exit For
    End If
Next

For qy = 0 To 17
    If y < 500 Then
    qy = 0
    Exit For
    End If
    If y > 500 * (qy + 1) - 250 And y < 500 * (qy + 1) + 250 Then
    Exit For
    End If
Next

Shape1.Left = 500 * (qx + 1) - 200
Shape1.Top = 500 * (qy + 1) - 200
Shape1.Visible = True '使用shape控件，并使其可见，从而显示出光标
End Sub
Public Function win(ByVal x As Integer, ByVal y As Integer)
'该子过程用于判断输赢
Dim i, j As Integer
Dim k(0 To 7) As Integer '定义一个八维的数组用于储存落子周围八个方向的棋子数目
Dim fun(0 To 3) As Integer '定义四个数组，储存四个方向总棋子数目，与上面那个的关系是：fun(0)=k(0)+k(4)
'对定义的数组进行初始化
For i = 0 To 7
k(i) = 0
Next

'以下八个片段的原理完全相同，获取八个方向同颜色的棋子数目
i = x - 1
j = y - 1
Do While i >= 0 And j >= 0
    If pill(i, j) = pill(x, y) Then '如果该方向连下去的棋子颜色与目标棋子颜色一致，则判断下一位的棋子颜色
    k(0) = k(0) + 1
    i = i - 1
    j = j - 1
    Else '否则的话终止判断，返回该方向同种棋子数目
    Exit Do
    End If
Loop

i = x
j = y - 1
Do While j >= 0
    If pill(i, j) = pill(x, y) Then
    k(1) = k(1) + 1
    j = j - 1
    Else
    Exit Do
    End If
Loop

i = x + 1
j = y - 1
Do While i <= 18 And j >= 0
    If pill(i, j) = pill(x, y) Then
    k(2) = k(2) + 1
    i = i + 1
    j = j - 1
    Else
    Exit Do
    End If
Loop

i = x + 1
j = y
Do While i <= 18
    If pill(i, j) = pill(x, y) Then
    k(3) = k(3) + 1
    i = i + 1
    Else
    Exit Do
    End If
Loop

i = x + 1
j = y + 1
Do While i <= 18 And j <= 18
    If pill(i, j) = pill(x, y) Then
    k(4) = k(4) + 1
    i = i + 1
    j = j + 1
    Else
    Exit Do
    End If
Loop

i = x
j = y + 1
Do While j <= 18
    If pill(i, j) = pill(x, y) Then
    k(5) = k(5) + 1
    j = j + 1
    Else
    Exit Do
    End If
Loop

i = x - 1
j = y + 1
Do While i >= 0 And j <= 18
    If pill(i, j) = pill(x, y) Then
    k(6) = k(6) + 1
    i = i - 1
    j = j + 1
    Else
    Exit Do
    End If
Loop

i = x - 1
j = y
Do While i >= 0
    If pill(i, j) = pill(x, y) Then
    k(7) = k(7) + 1
    i = i - 1
    Else
    Exit Do
    End If
Loop

fun(0) = k(0) + k(4)
fun(1) = k(1) + k(5)
fun(2) = k(2) + k(6)
fun(3) = k(3) + k(7)
'对四个方向的棋子数目进行判断，如果有一个方向棋子达到4个以上则获胜，已经将中央棋子去除
If fun(0) >= 4 Or fun(1) >= 4 Or fun(2) >= 4 Or fun(3) >= 4 Then
win = pill(x, y)
Else
win = 0
End If

End Function
