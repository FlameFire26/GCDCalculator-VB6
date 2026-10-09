VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "求最公数"
   ClientHeight    =   2985
   ClientLeft      =   5235
   ClientTop       =   5205
   ClientWidth     =   10755
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   10755
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "方正苏新诗柳楷简体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "方正毡笔黑简体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4043
      TabIndex        =   8
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "方正苏新诗柳楷简体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "方正苏新诗柳楷简体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "方正苏新诗柳楷简体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "方正苏新诗柳楷简体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "计算最大公约数的步骤次数"
      BeginProperty Font 
         Name            =   "方正水柱简体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   9
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "最小公倍数"
      BeginProperty Font 
         Name            =   "方正水柱简体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "最大公约数"
      BeginProperty Font 
         Name            =   "方正水柱简体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "输入2"
      BeginProperty Font 
         Name            =   "方正水柱简体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "输入1"
      BeginProperty Font 
         Name            =   "方正水柱简体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If IsNumeric(Text1.Text) And IsNumeric(Text2.Text) Then
If Text1.Text = 0 Or Text2.Text = 0 Then
Text3.Text = 0
Text4.Text = 0
Text5.Text = 0
Else
a = CInt(Text1.Text)
b = CInt(Text2.Text)

Do
If b > a Then
b = a + b
a = b - a
b = b - a
End If
c = a Mod b
a = b
b = c
d = d + 1
If b = 0 Then Exit Do
Loop Until a Mod b = 0
Text3.Text = b
Text5.Text = d

a = CInt(Text1.Text)
b = CInt(Text2.Text)
Text4.Text = a * b / Text3.Text
End If
Else
MsgBox "请输入有效的数字", vbOKOnly, "提示"
End If
End Sub

