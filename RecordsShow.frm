VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15090
   LinkTopic       =   "Form5"
   ScaleHeight     =   5955
   ScaleWidth      =   15090
   StartUpPosition =   3  '窗口缺省
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   5880
      TabIndex        =   3
      Top             =   1920
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   6588
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除记录"
      Height          =   735
      Left            =   11160
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      TabIndex        =   1
      Text            =   "(请选择要显示的单位)"
      Top             =   600
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   7920
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9360
      Top             =   360
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ben\Documents\VB Files\input\input\Counting.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ben\Documents\VB Files\input\input\Counting.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from empty where ID='Ben'"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tblnm As String
Private Sub Combo1_lostfocus()
    tblnm = LTrim(Combo1.Text)
End Sub

Private Sub Command1_Click()
    
    If tblnm = "" Then
        MsgBox ("请选择单位！")
    Else
        'Print ("select * from " & tblnm)
        Adodc1.RecordSource = "select Abstract,Ledger,Detail,Debit,Credit,Num,DT from " & tblnm & " order by Num asc"
        Adodc1.Refresh
        DataGrid1.Font.Size = 13
        DataGrid1.Columns(1).Width = 15
        Set DataGrid1.DataSource = Adodc1
    End If
End Sub

Private Sub Command2_Click()
    Dim delitem As Integer
    delitem = Val(InputBox("请输入要删除的凭证总号", "删除"))
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
    conn.Open
    
    conn.Execute "delete from " & tblnm & " where Num=" & delitem
    
    rs.Open "select * from " & tblnm & " order by Num asc", conn, adOpenStatic
    rs.MoveFirst
        
    
    While Not rs.EOF
        If rs.Fields("Num").Value > delitem Then
        rs.Fields("Num").Value = rs.Fields("Num").Value - 1
        End If
        rs.MoveNext
    Wend
    'Print ("minus" & rsminus.Bookmark & "--" & rsminus.Fields("Num").Value)
    rs.Close
    conn.Close
End Sub

Private Sub DataGrid1_LostFocus()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    Dim concomb As New ADODB.Connection 'connection for combo1
    Dim rscomb As New ADODB.Recordset
    
    
    concomb.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
    'concomb.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
    concomb.Open
    
    rscomb.CursorType = adOpenStatic
    rscomb.Open "select * from list", concomb
    If rscomb.RecordCount = 0 Then
        MsgBox ("没有记录显示") '有无数据表
    Else
        rscomb.MoveFirst
        While Not rscomb.EOF
           Combo1.AddItem (rscomb.Fields("Tablelist").Value)
           rscomb.MoveNext
        Wend
    End If
    
    rscomb.Close
    concomb.Close
    
End Sub

