VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   8850
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "print.frx":0000
      Left            =   2040
      List            =   "print.frx":0002
      TabIndex        =   1
      Text            =   "(ѡ��Ҫ��ӡ�ĵ�λ���·�)"
      Top             =   480
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ӡ"
      Height          =   1455
      Left            =   2160
      TabIndex        =   0
      Top             =   2760
      Width           =   3255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Transaction
    Date As String
    Abstract As String
    Ledger As String
    Detail As String
    Debit As Double
    credit As Double
    CD As String
    num As Integer
End Type

Dim tblnm As String

Dim data(1000) As Transaction '������


Dim i_prt As Integer 'ÿһ�м�¼��i
Dim num As Integer '��ǰ��ӡ�е��ܺ�

Dim nextpage As Single '����һ��

Dim a(30) As Single, b(30) As Single
Dim c(30) As Single, d(30) As Single
Dim e(5) As Single, f(5) As Single
Dim g(5) As Single, h(5) As Single


Private Function location(y)
    Erase a, b, c, d, e, f, g, h
    
    Dim i As Single
      
    
    a(0) = 2.4 '2.5
    a(1) = a(0) + 3.8 '6.3
    a(2) = a(1) + 2.9 '9.2
    a(3) = a(2) + 3 '12.2
    For i = 4 To 12
        a(i) = a(i - 1) + 0.322222
    Next i
    'Print ("a12=" & a(12)) '15.1
    a(13) = a(12) + 0.68 '15.78
    
    For i = 14 To 22
        a(i) = a(i - 1) + 0.322222
    Next i
    'Print ("a22=" & a(22)) '18.6799
    a(23) = a(22) + 0.68 '��a13
    
    
    b(0) = y '3.4
    b(1) = b(0) + 0.5  '3.9
    b(2) = b(1) + 0.5 '4.4
    For i = 3 To 9
        b(i) = b(i - 1) + 0.65714
    Next i
    
    'Print ("b9=" & b(9)) '8.999979
    b(10) = b(9) + 4.200021 '13.2
    b(11) = b(10) + 0.5 '13.7
    b(12) = b(11) + 0.5 '14.2
    For i = 13 To 19
        b(i) = b(i - 1) + 0.65714
    Next i
    
    'Print ("b19=" & a(19)) '17.71333
    b(20) = b(19) + 5.28667 '23
    b(21) = b(20) + 0.5 '23.5
    b(22) = b(21) + 0.5 '24
    
    
    For i = 0 To 30
        c(i) = a(i) * 567
        d(i) = b(i) * 567
    Next i
    
        
    
    e(0) = a(1)
    e(1) = a(14)
    e(2) = a(18)
    e(3) = a(22)
    f(2) = b(0) - 0.5
    f(1) = f(2) - 0.7
    f(0) = f(1) - 0.7
    
    For i = 0 To 4
        g(i) = e(i) * 567
        h(i) = f(i) * 567
    Next i

End Function


Private Function PRTtable(y As Single)

    While data(i_prt).num <> 0
        
        location (y)
    '------------------------------------------------------------start to print ----------------------------------------------------------
        Printer.Line (c(0), d(0))-Step(c(23) - c(0), 0)
        Printer.Line (c(3), d(1))-Step(c(12) - c(3), 0) '���º�
        For i = 2 To 9
            Printer.Line (c(0), d(i))-Step(c(23) - c(0), 0)
        Next i
        Printer.Line (c(13), d(1))-Step(c(22) - c(13), 0) '���º�
        
        '��
        '��Ŀ��
        For i = 0 To 3
            Printer.Line (c(i), d(0))-Step(0, d(9) - d(0))
        Next i
        '������
        For i = 4 To 11
            Printer.Line (c(i), d(1))-Step(0, d(9) - d(1))
        Next i
        
        Printer.Line (c(12), d(0))-Step(0, d(9) - d(0))
        Printer.Line (c(13), d(0))-Step(0, d(9) - d(0))
        
        
        '������
        
        For i = 14 To 21
            Printer.Line (c(i), d(1))-Step(0, d(9) - d(1))
        Next i
        Printer.Line (c(22), d(0))-Step(0, d(9) - d(0))
        Printer.Line (c(23), d(0))-Step(0, d(9) - d(0))
        
        
        
        Printer.CurrentX = ((c(1) + c(0)) / 2) - 400
        Printer.CurrentY = ((d(2) + d(0)) / 2) - 100
        Printer.FontSize = 10
        Printer.Print "ժ    Ҫ"
        
        
        Printer.CurrentX = ((c(2) + c(1)) / 2) - 400
        Printer.CurrentY = ((d(2) + d(0)) / 2) - 100
        Printer.FontSize = 10
        Printer.Print "���˿�Ŀ"
        
        
        Printer.CurrentX = ((c(3) + c(2)) / 2) - 700
        Printer.CurrentY = ((d(2) + d(0)) / 2) - 100
        Printer.FontSize = 10
        Printer.Print "��������ϸ��Ŀ"
        
        Printer.CurrentX = ((c(12) + c(3)) / 2) - 200
        Printer.CurrentY = ((d(1) + d(0)) / 2) - 90
        Printer.FontSize = 10
        Printer.Print "��  ��"
        
        
        Printer.CurrentX = ((c(22) + c(13)) / 2) - 200
        Printer.CurrentY = ((d(1) + d(0)) / 2) - 90
        Printer.FontSize = 10
        Printer.Print "��  ��"
        
        Dim units As String
        units = "��ʮ��ǧ��ʮԪ�Ƿ�"
        For i = 1 To 9
            Printer.CurrentX = ((c(3 + i) + c(2 + i)) / 2) - 80
            Printer.CurrentY = ((d(2) + d(1)) / 2) - 80
            Printer.FontSize = 8
            Printer.Print (Mid(units, i, 1))
        Next i
        Printer.CurrentX = ((c(13) + c(12)) / 2) - 150
        Printer.CurrentY = ((d(2) + d(0)) / 2) - 80
        Printer.FontSize = 15
        Printer.Print "��"
        
        For i = 1 To 9
            Printer.CurrentX = ((c(12 + i) + c(13 + i)) / 2) - 80
            Printer.CurrentY = ((d(2) + d(1)) / 2) - 80
            Printer.FontSize = 8
            Printer.Print (Mid(units, i, 1))
        Next i
        Printer.CurrentX = ((c(23) + c(22)) / 2) - 150
        Printer.CurrentY = ((d(2) + d(0)) / 2) - 80
        Printer.FontSize = 15
        Printer.Print "��"
        
        Dim rght As String
        rght = "�����ݼ�����"
        For i = 1 To 6
            Printer.CurrentX = c(23) + 100
            Printer.CurrentY = d(i - 1)
            Printer.FontSize = 10
            Printer.Print (Mid(rght, i, 1))
        Next i
        
        Printer.CurrentX = c(23) + 100
        Printer.CurrentY = d(7)
        Printer.FontSize = 10
        Printer.Print ("��")
        
        
        Printer.CurrentX = ((c(1) + c(0)) / 2) - 400
        Printer.CurrentY = ((d(8) + d(9)) / 2) - 100
        Printer.FontSize = 10
        Printer.Print "��    ��"
          
        Printer.CurrentX = ((c(1) + c(0)) / 2) - 400
        Printer.CurrentY = d(9) + 70
        Printer.FontSize = 8
        Printer.Print "���"
        
        Printer.CurrentX = c(1) + 70
        Printer.CurrentY = d(9) + 70
        Printer.FontSize = 8
        Printer.Print "����"
        
        Printer.CurrentX = ((c(2) + c(3)) / 2) - 400
        Printer.CurrentY = d(9) + 70
        Printer.FontSize = 8
        Printer.Print "����"
        
        
        Printer.CurrentX = c(7)
        Printer.CurrentY = d(9) + 70
        Printer.FontSize = 8
        Printer.Print "�Ʊ�"
    '------------------------------print head------------------------------
        
        '�ܺ�
        Printer.Line (g(1), h(0))-Step(g(3) - g(1), 0)
        Printer.Line (g(1), h(1))-Step(g(3) - g(1), 0)
        Printer.Line (g(1), h(2))-Step(g(3) - g(1), 0)
        
        Printer.Line (g(1), h(0))-Step(0, h(2) - h(0))
        Printer.Line (g(2), h(0))-Step(0, h(2) - h(0))
        Printer.Line (g(3), h(0))-Step(0, h(2) - h(0))
        
        
        Printer.CurrentX = ((g(1) + g(2)) / 2) - 200
        Printer.CurrentY = ((h(0) + h(1)) / 2) - 100
        Printer.FontSize = 12
        Printer.Print "�ܺ�"
        
        Printer.CurrentX = ((g(1) + g(2)) / 2) - 200
        Printer.CurrentY = ((h(1) + h(2)) / 2) - 100
        Printer.FontSize = 12
        Printer.Print "�ֺ�"
        
        Printer.CurrentX = g(0) + 200
        Printer.CurrentY = h(0) + 100
        Printer.FontSize = 15
        Printer.Print "��   ��   ƾ   ֤"
        
        Printer.CurrentX = g(0) + 500
        Printer.CurrentY = h(1) + 100
        Printer.FontSize = 10
        Printer.Print Year(data(0).Date) & " �� " & Month(data(0).Date) & " �� " & Day(data(0).Date) & " ��"
        
        Printer.CurrentX = ((g(2) + g(3)) / 2) - 200
        Printer.CurrentY = ((h(0) + h(1)) / 2) - 100
        Printer.FontSize = 12
        Printer.Print data(i_prt).num
    
        
        
        
        
        'Printer.CurrentX = g(0) + 500
        'Printer.CurrentY = h(1) + 100
        'Printer.FontSize = 10
        'Printer.Print "2019 �� 2 �� 25 ��"
        
        'i_prt_words = 2
        PRTwords (2) 'change i_prt-->data(i_prt).num
        
        num = num + 1
        If nextpage <> 3 Then
            'Print ("abc" & nextpage)
            nextpage = nextpage + 1
            PRTtable (y + 9.9) 'next table on same page
        Else
            Printer.NewPage
            'Print ("def" & nextpage)
            nextpage = 1
            
            PRTtable (3.4)
        End If
    Wend
End Function
    
    
Private Function PRTwords(i_prt_words As Single) 'i ����һ����һ���ֵ�����
    While data(i_prt).num = num
        Printer.CurrentX = c(0) + 100
        Printer.CurrentY = ((d(i_prt_words + 1) + d(i_prt_words)) / 2) - 100
        Printer.FontSize = 8
        Printer.Print data(i_prt).Abstract
        
        
        Printer.CurrentX = c(1) + 100
        Printer.CurrentY = ((d(i_prt_words + 1) + d(i_prt_words)) / 2) - 100
        Printer.FontSize = 8
        Printer.Print data(i_prt).Ledger
        
        
        Printer.CurrentX = c(2) + 100
        Printer.CurrentY = ((d(i_prt_words + 1) + d(i_prt_words)) / 2) - 100
        Printer.FontSize = 8
        Printer.Print data(i_prt).Detail
        ''''''
        Dim number_prt() As Integer
        number_prt() = seperatenumber((data(i_prt).Debit))
        Dim flag As Boolean '�ж�0��ǰ�滹�Ǻ���
        flag = False
        For i = 0 To 8
            If number_prt(i) <> 0 Or flag Then
                Printer.CurrentX = c(3 + i) - 30
                Printer.CurrentY = ((d(i_prt_words + 1) + d(i_prt_words)) / 2) - 90
                Printer.FontSize = 8
                Printer.Print number_prt(i)
                flag = True
            End If
        Next i
        Erase number_prt
        
        number_prt() = seperatenumber((data(i_prt).credit))
        
        flag = False
        For i = 0 To 8
            If number_prt(i) <> 0 Or flag Then
                Printer.CurrentX = c(13 + i) - 30
                Printer.CurrentY = ((d(i_prt_words + 1) + d(i_prt_words)) / 2) - 90
                Printer.FontSize = 8
                Printer.Print number_prt(i)
                flag = True
            End If
        Next i
        Erase number_prt
    
        i_prt = i_prt + 1
        PRTwords (i_prt_words + 1)
    
    Wend
    
    
'----------------------------------------------------------------------------ȷ���м��--------------------------------------
    'Print ("g(0)=" & g(0))
    'Print ("d(9)=" & d(9))
    'Print ("g0-d9" & g(0) - d(9))
    
    'Dim paper_length As Single, blank As Single, sheetWithBlank As Single, sheet As Single
    
    
    'Dim secondSheet As Single
    'sheet = d(9) - g(0)
    'paper_length = 29.7 * 567
    'blank = paper_length / 3 - (sheet)
    
    'sheetWithBlank = blank + sheet
    
    'secondSheet = 3.4 + sheetWithBlank / 567
    'Print ("blankWithSheet = " & sheetWithBlank)
    'Print ("secondSheet=" & secondSheet)
    
    ''        Printer.EndDoc
'------------------------------------------------------------------------------------------------------------------

    



End Function

Private Sub Combo1_lostfocus()
    tblnm = LTrim(Combo1.Text)
    
    
    Dim i As Integer
    Dim conprt As New ADODB.Connection
    Dim rsprt As New ADODB.Recordset
    
    
    
    conprt.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
    'conprt.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
    conprt.Open
    
    rsprt.CursorType = adOpenStatic
    'Print ("select * from " & tblnm)
    rsprt.Open "select * from " & tblnm, conprt
    'Print rsprt.RecordCount
    i = 0
    If rsprt.RecordCount = 0 Then
        MsgBox "û�е��¼�¼"
    Else
        rsprt.MoveFirst
        While Not rsprt.EOF
            data(i).Date = rsprt.Fields("DT").Value
            data(i).Abstract = rsprt.Fields("Abstract").Value
            data(i).Ledger = rsprt.Fields("Ledger").Value
            data(i).Detail = rsprt.Fields("Detail").Value
            data(i).Debit = rsprt.Fields("Debit").Value
            data(i).credit = rsprt.Fields("Credit").Value
            data(i).num = rsprt.Fields("num").Value
            rsprt.MoveNext
            i = i + 1
        Wend
    End If
    num = 1
    
    rsprt.Close
    conprt.Close


End Sub



Private Sub Command1_Click()

    If Combo1.Text = "" Then
        MsgBox ("��ѡ��λ��")
    Else
        Call PRTtable(3.4)
        Printer.EndDoc
    End If

End Sub



Private Function seperatenumber(number As Single) As Integer()
    Dim seperate(0 To 8) As Integer

    seperate(0) = 100 * number \ 10 ^ (8)
    For i = 1 To 8
        seperate(i) = (100 * number Mod (10 ^ (9 - i))) \ (10 ^ (8 - i))
    Next i
    seperatenumber = seperate
End Function




Private Sub Form_Load()
    Dim concomb As New ADODB.Connection 'connection for combo1
    Dim rscomb As New ADODB.Recordset
    
    
    concomb.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
    'concomb.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\Counting.mdb;Persist Security Info=False"
    concomb.Open
    
    rscomb.CursorType = adOpenStatic
    'Print ("select * from " & tblnm)
    rscomb.Open "select * from list", concomb
    If rscomb.RecordCount = 0 Then
        MsgBox ("û�м�¼���Դ�ӡ") '�������ݱ�
    Else
        rscomb.MoveFirst
        While Not rscomb.EOF
           Combo1.AddItem (rscomb.Fields("Tablelist").Value)
           rscomb.MoveNext
        Wend
    End If
    
    rscomb.Close
    concomb.Close
    
    Erase data
    i_prt = 0
    nextpage = 1
'------------------------------get data--------------------
'At combo1_lostfocus

End Sub
