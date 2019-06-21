VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   12615
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      Caption         =   "显示已有账单"
      Height          =   495
      Left            =   9960
      TabIndex        =   51
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "显示修改已有科目"
      Height          =   495
      Left            =   9960
      TabIndex        =   50
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "print tblnm"
      Height          =   375
      Left            =   1680
      TabIndex        =   49
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Main.frx":0000
      Left            =   240
      List            =   "Main.frx":0002
      TabIndex        =   48
      Text            =   "(请选择单位)"
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打印凭证"
      Height          =   495
      Left            =   9960
      TabIndex        =   47
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "list additem"
      Height          =   975
      Left            =   7800
      TabIndex        =   46
      Top             =   5520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   600
      Left            =   4440
      TabIndex        =   45
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "DEL TABL"
      Height          =   375
      Left            =   10680
      TabIndex        =   44
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "添加凭证"
      Height          =   1095
      Left            =   9960
      TabIndex        =   43
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text33 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   42
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text32 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   41
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text31 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   40
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Text30 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   39
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox Text29 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text28 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   37
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text27 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   36
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text26 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   35
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text25 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   34
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox Text24 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text23 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   32
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text22 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   31
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   30
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox Text20 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   29
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   27
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   26
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   25
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   24
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   22
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   21
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   20
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "增加科目"
      Height          =   495
      Left            =   9960
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "贷方"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8280
      TabIndex        =   12
      Top             =   1680
      Width           =   945
   End
   Begin VB.Label Label8 
      Caption         =   "借方"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6600
      TabIndex        =   11
      Top             =   1680
      Width           =   1140
   End
   Begin VB.Label Label7 
      Caption         =   "二级或明细科目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3960
      TabIndex        =   10
      Top             =   1680
      Width           =   2340
   End
   Begin VB.Label Label6 
      Caption         =   "总账科目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      TabIndex        =   9
      Top             =   1680
      Width           =   1740
   End
   Begin VB.Label Label5 
      Caption         =   "摘要"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   1260
   End
   Begin VB.Label Label4 
      Caption         =   "日"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6720
      TabIndex        =   3
      Top             =   960
      Width           =   420
   End
   Begin VB.Label Label3 
      Caption         =   "月"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5520
      TabIndex        =   2
      Top             =   960
      Width           =   420
   End
   Begin VB.Label Label2 
      Caption         =   "年"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4440
      TabIndex        =   1
      Top             =   960
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "记  账  凭  证"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Dim read(1000) As String
    Dim tblnm As String, DT As Date
    Private Type Transaction
    Date As String
    Abstract As String
    Ledger As String
    Detail As String
    Debit As Double
    credit As Double
    CD As String
    num As Integer
    Print As Boolean
    End Type
'文本框调动程序
Public Function chge(piny As String, i As Integer) As String 'i 哪个文本框



    Dim connsear As New ADODB.Connection
    Dim rssear As New ADODB.Recordset
    
    Dim newitem As String
    connsear.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"

    'connsear.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
    connsear.Open
    'rs2 查询表是否存在
    rssear.CursorType = adOpenStatic
    rssear.Open "select Item from Items where Pinyin ='" & piny & "'", connsear 'https://bbs.csdn.net/topics/39386
    If rssear.RecordCount = 0 Then
        confirm = MsgBox("未找到此科目，是否需要新建", vbYesNo, "确认新建")
        If confirm = vbYes Then
            newitem = InputBox("请输入拼音" & piny & "对应的科目", "添加科目")
            connsear.Execute "insert into Items (Item,Pinyin) VALUES ('" & newitem & "','" & piny & "')"
            chge = newitem
        Else
            'Me.Controls("Text" & i).SelStart = 0'return = 0
            Me.Controls("Text" & i).SetFocus
        End If
    Else
        chge = rssear.Fields("Item").Value
    End If
    
End Function
   

Private Sub Combo1_lostfocus()
    tblnm = LTrim(Combo1.Text) + Text1.Text & Label2.Caption & Text2.Text & Label3.Caption & Text3.Text & Label4.Caption

    'Text2.SetFocus
    'Me.Controls("Text" & ).SetFocus
End Sub

Private Sub Command1_Click() '批量
    Form2.Show
End Sub


Private Sub Command2_Click() 'print
    Form3.Show
End Sub

Private Sub Command4_Click() 'ado
    
    'tblnm = LTrim(Combo1.Text) + LTrim(Str(Year(Now - Day(Now)) & Month(Now - Day(Now))))
    If tblnm = "" Then
        MsgBox ("请选择单位！")
    Else
        
        Dim sum_debit As Single, sum_credit As Single
        Dim k As Integer
        sum_debit = 0
        sum_credit = 0
        For k = 7 To 32 Step 5 'debit
            sum_debit = sum_debit + Val(Me.Controls("Text" & (k)).Text)
        Next k
        
        For k = 8 To 33 Step 5 'credit
            sum_credit = sum_credit + Val(Me.Controls("Text" & (k)).Text)
        Next k
        
        If sum_credit <> sum_debit Then
            MsgBox "借贷不等,请检查！"
        
        Else
        
            Field = " (ID int IDENTITY(1,1) ,DT datetime,Abstract VarChar(255), Ledger VarChar(255), Detail VarChar(255), Debit decimal(9, 2), Credit decimal(9, 2), CD varChar(255),Num decimal(9,0))" ',PrintOrNot Bit   https://blog.csdn.net/hfly2005/article/details/388809
            'Field = " (ID int IDENTITY(1,1) ,日期 datetime,摘要 VarChar(255), 总账科目 VarChar(255), 二级或明细科目 VarChar(255), 贷方 decimal(9, 2), 借方 decimal(9, 2),总号 decimal(9,0))" ',PrintOrNot Bit   https://blog.csdn.net/hfly2005/article/details/388809
            Dim conn As New ADODB.Connection
            Dim rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
            Dim num As Integer '总号
            conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
            'conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
            conn.Open
            'rs2 查询表是否存在
            rs2.CursorType = adOpenStatic
            rs2.Open "select * from list where Tablelist='" & tblnm & "'", conn 'https://bbs.csdn.net/topics/39386
            If Not rs2.RecordCount = 1 Then
                conn.Execute "create table " & tblnm & Field
                conn.Execute "INSERT INTO list (Tablelist) VALUES ('" & tblnm & "')"
                'conn.Execute "INSERT INTO " & tblnm & "(Abstract,Credit) VALUES ('Gates','55555.12')"
                'MsgBox "创建表成功"
            End If
            'rs 确定总号
            rs.CursorType = adOpenStatic
            rs.Open "select * from " & tblnm & " order by Num asc", conn
            'rs.Open "select * from " & tblnm & " order by 总号 asc", conn
            If rs.RecordCount = 0 Then
                num = 1
            Else
                rs.MoveLast
                'Print (rs.Fields("Num").Value)
                num = rs.Fields("Num").Value + 1
            End If
            rs.Close
            
            Print num
            '
            Dim writein(6) As Transaction
            
            Dim i As Integer '(texti)
            Dim j As Integer, jmax As Integer 'writein(j)
            j = 0
            For i = 7 To 32 Step 5
            
                If (Me.Controls("Text" & i).Text) <> "" Then
                    j = j + 1
                    writein(j).Date = DT
                    writein(j).Abstract = Me.Controls("Text" & (i - 3)).Text 'http://tieba.baidu.com/p/868677926?traceid=
                    writein(j).Ledger = Me.Controls("Text" & (i - 2)).Text
                    writein(j).Detail = Me.Controls("Text" & (i - 1)).Text
                    writein(j).Debit = Val(Me.Controls("Text" & (i)).Text)
                    writein(j).CD = "Debit"
                    writein(j).num = num
                End If
            Next i
            jmax = j
            Print (Date)
            For j = 1 To jmax
            conn.Execute "insert into " & tblnm & " (DT,Abstract,Ledger,Detail,Debit,Credit,CD,Num) VALUES ('" & writein(j).Date & "','" & writein(j).Abstract & "','" & writein(j).Ledger & "','" & writein(j).Detail & "','" & writein(j).Debit & "','" & writein(j).credit & "','" & writein(j).CD & "','" & writein(j).num & "')"
            'conn.Execute "insert into " & tblnm & " (日期,摘要,总账科目,二级或明细科目，借方,贷方,总号) VALUES ('" & writein(j).Date & "','" & writein(j).Abstract & "','" & writein(j).Ledger & "','" & writein(j).Detail & "','" & writein(j).Debit & "','" & writein(j).credit & "','" & writein(j).num & "')"
            Next j
            Erase writein
            j = 0
            For i = 8 To 33 Step 5
            
                If (Me.Controls("Text" & i).Text) <> "" Then
                    j = j + 1
                    writein(j).Date = DT
                    writein(j).Abstract = Me.Controls("Text" & (i - 4)).Text
                    writein(j).Ledger = Me.Controls("Text" & (i - 3)).Text
                    writein(j).Detail = Me.Controls("Text" & (i - 2)).Text
                    writein(j).credit = Val(Me.Controls("Text" & (i)).Text)
                    writein(j).CD = "Credit"
                    writein(j).num = num
                End If
            Next i
            jmax = j
            For j = 1 To jmax
            conn.Execute "insert into " & tblnm & " (DT,Abstract,Ledger,Detail,Debit,Credit,CD,Num) VALUES ('" & writein(j).Date & "','" & writein(j).Abstract & "','" & writein(j).Ledger & "','" & writein(j).Detail & "','" & writein(j).Debit & "','" & writein(j).credit & "','" & writein(j).CD & "','" & writein(j).num & "')"
            'conn.Execute "insert into " & tblnm & " (日期,摘要,总账科目,二级或明细科目，借方,贷方,总号) VALUES ('" & writein(j).Date & "','" & writein(j).Abstract & "','" & writein(j).Ledger & "','" & writein(j).Detail & "','" & writein(j).Debit & "','" & writein(j).credit & "','" & writein(j).num & "')"
            Next j
            Erase writein
            For k = 4 To 33
                Me.Controls("Text" & k).Text = ""
            Next k
            rs2.Close
            conn.Close
        End If
    End If
End Sub



Private Sub Command5_Click()
Print tblnm
End Sub

Private Sub Command6_Click() 'drop table
    
    Dim conn As New ADODB.Connection
    Dim rs2 As New ADODB.Recordset
    
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
    'conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
    conn.Open
    rs2.CursorType = adOpenStatic
    rs2.Open "select * from list where Tablelist='" & tblnm & "'", conn 'https://bbs.csdn.net/topics/39386
    If rs2.RecordCount = 1 Then
        conn.Execute "drop table " & tblnm
        conn.Execute "delete from list where Tablelist='" & tblnm & "'"
        'MsgBox "删除表成功"
    End If
    rs2.Close
    conn.Close
End Sub


Private Sub Command3_Click()
    
    For i = 0 To i_max
       Print (read(i))
       List1.AddItem (read(i))
    Next i
End Sub



Private Sub Command7_Click()
    Form4.Show
End Sub

Private Sub Command8_Click()
    Form5.Show
End Sub

Private Sub Form_Load()
    Text1.Text = Year(Now - Day(Now))
    Text2.Text = Month(Now - Day(Now))
    Text3.Text = Day(Now - Day(Now))
    Combo1.AddItem ("亚华")
    Combo1.AddItem ("广屹")
    Combo1.AddItem ("硬质合金")
    Combo1.AddItem ("杭屹")
    'tblnm = Combo1.Text + Str(Year(Now - Day(Now)) & Month(Now - Day(Now)))
    DT = Text1.Text & "/" & Text2.Text & "/" & Text3.Text
    'i = 0
    'Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
    'Adodc1.RecordSource = "select * from Items"
    'Adodc1.Refresh
    'Adodc1.Recordset.MoveFirst
    'While Not Adodc1.Recordset.EOF
    '    read(i) = Adodc1.Recordset.Fields("Item").Value
    '    'Combo1.AddItem (Adodc1.Recordset.Fields("Item").Value)
    '    i = i + 1
    '    Adodc1.Recordset.MoveNext
    'Wend
    'i_max = i - 1'在cmd3中会用
    
  

End Sub

Private Sub Text5_lostfocus()
    If Not Text5.Text = "" Then Text5.Text = chge(Text5.Text, 5)
End Sub
Private Sub Text10_lostfocus()
    If Not Text10.Text = "" Then Text10.Text = chge(Text10.Text, 10)
End Sub
Private Sub Text15_lostfocus()
    If Not Text15.Text = "" Then Text15.Text = chge(Text15.Text, 15)
End Sub
Private Sub Text20_lostfocus()
    If Not Text20.Text = "" Then Text20.Text = chge(Text20.Text, 20)
End Sub
Private Sub Text25_lostfocus()
    If Not Text25.Text = "" Then Text25.Text = chge(Text25.Text, 25)
End Sub
Private Sub Text30_lostfocus()
    If Not Text30.Text = "" Then Text30.Text = chge(Text30.Text, 30)
End Sub

