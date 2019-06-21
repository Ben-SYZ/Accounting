VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   4335
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "cross"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "dot"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      Height          =   615
      Left            =   1920
      TabIndex        =   11
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label3 
      Height          =   615
      Left            =   1920
      TabIndex        =   10
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label2 
      Height          =   615
      Left            =   1920
      TabIndex        =   9
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   1560
      TabIndex        =   8
      Top             =   2160
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 6) As Single
Dim DA As Single
Dim CA(1 To 3)  As Single


Private Sub Command1_Click()
 DA = 0
 a(1) = Val(Text1.Text)
 a(2) = Val(Text2.Text)
 a(3) = Val(Text3.Text)
 a(4) = Val(Text4.Text)
 a(5) = Val(Text5.Text)
 a(6) = Val(Text6.Text)
 For i = 1 To 3
 DA = DA + a(i) * a(i + 3)
 Next i
 Label1.Caption = Str(DA)
End Sub


Private Sub Command2_Click()
 For i = 1 To 3
 CA(i) = 0
 Next i
 a(1) = Val(Text1.Text)
 a(2) = Val(Text2.Text)
 a(3) = Val(Text3.Text)
 a(4) = Val(Text4.Text)
 a(5) = Val(Text5.Text)
 a(6) = Val(Text6.Text)
 CA(1) = a(2) * a(6) - a(5) * a(3)
 CA(2) = a(3) * a(4) - a(6) * a(1)
 CA(3) = a(1) * a(5) - a(4) * a(2)
 Label2.Caption = Str(CA(1))
 Label3.Caption = Str(CA(2))
 Label4.Caption = Str(CA(3))
End Sub

