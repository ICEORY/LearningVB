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
   StartUpPosition =   3  '����ȱʡ
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
'��������ȫ�ֱ���
Dim coun As Integer '���ڱ�ʾ������ܲ���
Dim pill(0 To 18, 0 To 18) As Integer '���ڴ�����������꣬����pill(0,1)=1��ʾ�������ڵ�һ�еڶ���λ�ã�������Ϊ����
'��pill(0,1)=2,��ʾ���˰��ӣ���ֵΪ�����λ��Ϊ��
Private Sub Form_Load() '���̳�ʼ������ʾ19�������Լ�19�����߲�����Ӧλ�û�������ڵ�
Dim i, j As Integer '����������ʱ���������ڼ�����
coun = 0 '��������
Pic1.AutoRedraw = True '����picture�ؼ����Զ��ػ湦�ܣ�ʹͼ������ȷ����
Form1.Height = 10500 '���ô����С����Ϊ10500
Form1.Width = 11500 '��Ϊ10500
Form1.Caption = "������" '���ô�������

For i = 0 To 18 '����������λ�õ��������
    For j = 0 To 18
    pill(i, j) = 0
    Next
Next

For i = 0 To 18 '���������ϵĺ���
Pic1.Line (500, 500 + 500 * i)-(9500, 500 + 500 * i)
Pic1.Line (500 + 500 * i, 500)-(500 + i * 500, 9500)
Next

'������������û���Բ�Ĺ��̣�����������Ӧλ�û���С�ڵ�
Call draw(3, 3, 0)
Call draw(9, 9, 0)
Call draw(15, 15, 0)
Call draw(3, 15, 0)
Call draw(15, 3, 0)
End Sub


Public Sub draw(ByVal x As Integer, ByVal y As Integer, ByVal n As Integer) '����Բ��������СԲ����Ϊ��־�㣬���ƴ�ԲΪ����
Dim r As Integer '����뾶
'n����ѡ�����Բ������
'n=0��ʾ���Ʊ�־��
If n = 0 Then
Pic1.FillColor = vbBlack
Pic1.FillStyle = 0
r = 50
End If
'���ƺ���
If n = 1 Then
Pic1.FillColor = vbBlack
Pic1.FillStyle = 0
r = 200
End If
'���ư���
If n = 2 Then
Pic1.FillColor = vbWhite
Pic1.FillStyle = 0
r = 200
End If

Pic1.Circle (500 * x + 500, 500 * y + 500), r
End Sub
Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) '��ȡ��갴�µ�λ��
Dim qx, qy As Integer '���������������ڱ�ʾ���ӵ�����

For qx = 0 To 17 '�����λ����������λ�ö�Ӧ
    If x < 500 Then '����겻�������ϵ��������qx����
    qx = 0
    Exit For
    End If
    '����������������������һ���ķ�Χ�ڶ�Ӧ
    If x > 500 * (qx + 1) - 250 And x < 500 * (qx + 1) + 250 Then
    Exit For
    End If
Next
'�����������������������Ӧ������ͬ��
For qy = 0 To 17
    If y < 500 Then
    qy = 0
    Exit For
    End If
    
    If y > 500 * (qy + 1) - 250 And y < 500 * (qy + 1) + 250 Then
    Exit For
    End If
Next
'�ж��Ƿ��λ���Ƿ��Ѿ��������ˣ��еĻ������������ȴ��´����ĵ��
If pill(qx, qy) <> 0 Then
GoTo end_1
End If
'�����λ��û�����ӣ����ڸ�λ�û�������
Call draw(qx, qy, (coun Mod 2) + 1)
pill(qx, qy) = coun Mod 2 + 1
'�ж���Ӯ��ֻ��Ҫ�жϱ������ӵ���Χ�Ƿ�������ţ����ݵ���ԭ�����֪���÷����Ŀ�����
If win(qx, qy) = 0 Then
coun = coun + 1
End If
'����Ӯ����Ľ��ִ����Ӧ�Ĳ���
If win(qx, qy) = 1 Then
MsgBox "����Ӯ��"
Pic1.Cls
Call Form_Load
End If

If win(qx, qy) = 2 Then
MsgBox "����Ӯ��"
Pic1.Cls
Call Form_Load
End If

end_1: '��λ�ö�Ӧ֮ǰ��goto ���
End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) '��ȡ����ƶ������е�����
Dim qx, qy As Integer '��������������������
'���ӹ�����Ҫ����ʾ����������϶��ڵ�λ�ã�����ʾ��꣬���Ա�������λ���ж�ʧ������
'���µ�ԭ�����ȡ��������ķ���һ�����������Ĳ���ԭ���Ļ����Ҳ����
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
Shape1.Visible = True 'ʹ��shape�ؼ�����ʹ��ɼ����Ӷ���ʾ�����
End Sub
Public Function win(ByVal x As Integer, ByVal y As Integer)
'���ӹ��������ж���Ӯ
Dim i, j As Integer
Dim k(0 To 7) As Integer '����һ����ά���������ڴ���������Χ�˸������������Ŀ
Dim fun(0 To 3) As Integer '�����ĸ����飬�����ĸ�������������Ŀ���������Ǹ��Ĺ�ϵ�ǣ�fun(0)=k(0)+k(4)
'�Զ����������г�ʼ��
For i = 0 To 7
k(i) = 0
Next

'���°˸�Ƭ�ε�ԭ����ȫ��ͬ����ȡ�˸�����ͬ��ɫ��������Ŀ
i = x - 1
j = y - 1
Do While i >= 0 And j >= 0
    If pill(i, j) = pill(x, y) Then '����÷�������ȥ��������ɫ��Ŀ��������ɫһ�£����ж���һλ��������ɫ
    k(0) = k(0) + 1
    i = i - 1
    j = j - 1
    Else '����Ļ���ֹ�жϣ����ظ÷���ͬ��������Ŀ
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
'���ĸ������������Ŀ�����жϣ������һ���������Ӵﵽ4���������ʤ���Ѿ�����������ȥ��
If fun(0) >= 4 Or fun(1) >= 4 Or fun(2) >= 4 Or fun(3) >= 4 Then
win = pill(x, y)
Else
win = 0
End If

End Function
