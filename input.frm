VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   8955
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7560
      Top             =   1320
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "input.frx":0000
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "进数据库"
      Height          =   735
      Left            =   5760
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "添加科目"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "拆解"
      Enabled         =   0   'False
      Height          =   855
      Left            =   1320
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "input.frx":0026
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "拼音"
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
      Left            =   5280
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "科目"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Line Line11 
      X1              =   4215
      X2              =   4695
      Y1              =   3450
      Y2              =   3450
   End
   Begin VB.Line Line10 
      X1              =   4215
      X2              =   4695
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Line Line9 
      X1              =   4200
      X2              =   4680
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line Line8 
      X1              =   4215
      X2              =   4695
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Line Line7 
      X1              =   4200
      X2              =   4680
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Line Line6 
      X1              =   4185
      X2              =   4665
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line Line5 
      X1              =   4200
      X2              =   4680
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line Line4 
      X1              =   4200
      X2              =   4680
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line Line3 
      X1              =   4185
      X2              =   4665
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Line Line2 
      X1              =   4200
      X2              =   4680
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Line Line1 
      X1              =   4185
      X2              =   4665
      Y1              =   1005
      Y2              =   1005
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Item
    Name As String
    piny As String
End Type

'Dim items(1000) As Item
Dim i As Integer, j As Integer, k As Integer
Private Function selwrd(lon As String) As String()

    Dim tem(1000) As String
    For i = 0 To 1000
    tem(i) = ""
    Next i
    j = 0
    
    Dim flag As Boolean
    
    
    For i = 1 To Len(lon)
        flag = False
        k = i
        
        Do While (Mid(lon, k, 1) <> Chr(10)) And (Mid(lon, k, 1) <> Chr(13)) And (k <= Len(lon))
            k = k + 1
            'Print (k)
            flag = True
        Loop 'k points to \n
        
        If flag = True Then
        tem(j) = Mid(lon, i, k - i)
        'Print (tem(j))
        j = j + 1
        End If
        
        i = k + 1 'because of next i
    
    Next i
    j_max = j - 1
    
    
    'print array tem()
    For j = 0 To j_max
        Print ("::" & tem(j))
    Next j

    selwrd = tem

End Function

Private Function Inpt(items() As String, py() As String, num As Integer)


confirm = MsgBox("确定要增加这" & num + 1 & "条记录吗？", vbYesNo, "确认增加")
If confirm = vbYes Then
    
    Dim adoin As New ADODB.Connection
    adoin.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
    'adoin.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
    adoin.Open
    For i = 0 To num
        adoin.Execute "insert into Items (Item,Pinyin) VALUES ('" & items(i) & "','" & py(i) & "')"
    Next i
    adoin.Close
    'Adodc1.Refresh
    'For j = 0 To num '输入进数据库
    '    Print (items(j))
    '    Adodc1.Recordset.AddNew
    '    Adodc1.Recordset.Fields("Item") = items(j)
    '    Adodc1.Recordset.Update
    'Next j
    
    MsgBox "记录增加成功！", , "Message"
Else
    MsgBox "记录未增加！", , "Message"
End If


End Function
Private Sub Command1_Click() 'one word



j = 0
For i = 1 To Len(Text1.Text)

'Print ("::::::" & Mid(Text1.Text, i, 1))
'If (Mid(Text1.Text, i, 1) <> Chr(10)) And (Mid(Text1.Text, i, 1) <> Chr(13)) Then
items(j) = Mid(Text1.Text, i, 1)
Print (":::::" & items(j))

'j = j + 1
'End If


Next i
End Sub

Private Sub command2_click() 'add names

Dim itemsnm() As String
Dim itemspy() As String
'Text2.Text = Split(Text1.Text)

'itemsnm = CopyMemory(selwrd(Text1.Text))
itemsnm() = Split(Text1.Text, vbCrLf)

itemspy() = Split(Text2.Text, vbCrLf)
Dim lennm As Integer, lenpy As Integer
lennm = UBound(itemsnm)
lenpy = UBound(itemspy)

'For i = 0 To lennm 'print for comfirming
'    Me.CurrentY = 300 * i
'    Print (itemsnm(i))
'Next i

'For i = 0 To lenpy 'print for comfirming
'    Me.CurrentX = 800
'    Me.CurrentY = 300 * i
'    print (itemspy(i))
'Next i
'selwrd (Text2.Text)
'Print selwrd
If lennm <> lenpy Then
    MsgBox "科目数和拼音数不符（最后不要加回车），请检查！", , "Message"
Else
    
    For i = 0 To lennm 'print for comfirming
        Me.CurrentY = 300 * i
        Print (itemsnm(i))
        Me.CurrentX = 800
        Me.CurrentY = 300 * i
        Print (itemspy(i))
    Next i
    Call Inpt(itemsnm(), itemspy(), lennm)
End If

'For i = 0 To jmax
'Print (itemsnm(i))

'Next i
End Sub

Private Sub Command3_Click()
Adodc1.Refresh
Dim j As Integer

For j = 0 To j_max
Print (items(j))
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("Item") = items(j)
Adodc1.Recordset.Update
Next j

End Sub

Private Sub Command4_Click()
'Form2.Hide
Unload (Form2)
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
Adodc1.RecordSource = "select * from Items"
End Sub
