VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   Caption         =   "获取网页源文件"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   FillColor       =   &H80000005&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   10980
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCommand2 
      Caption         =   "Command2"
      Height          =   360
      Left            =   7965
      TabIndex        =   9
      Top             =   90
      Width           =   990
   End
   Begin VB.TextBox txtHttpWww 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   1350
      TabIndex        =   0
      Text            =   "http://www.vodtw.com/html/book/28/28902/"
      Top             =   90
      Width           =   6225
   End
   Begin VB.CheckBox chk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "自动下载下一章"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8910
      TabIndex        =   4
      Top             =   450
      Value           =   1  'Checked
      Width           =   1950
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2385
      Top             =   1260
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000005&
      Height          =   2505
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   4635
      Width           =   12735
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   675
      Left            =   45
      TabIndex        =   6
      Top             =   810
      Width           =   10875
      ExtentX         =   19182
      ExtentY         =   1191
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "获取"
      Default         =   -1  'True
      Height          =   375
      Left            =   7875
      TabIndex        =   3
      Top             =   405
      Width           =   855
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H80000007&
      Height          =   1425
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1530
      Width           =   10890
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   1350
      TabIndex        =   2
      Text            =   "http://www.vodtw.com/html/book/28/28902/21716386.html"
      Top             =   450
      Width           =   6225
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "书记目录网址："
      Height          =   180
      Left            =   45
      TabIndex        =   8
      Top             =   135
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "当前阅读网址："
      Height          =   255
      Left            =   45
      TabIndex        =   5
      Top             =   495
      Width           =   1650
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim filename


Private Sub cmdCommand2_Click()
    Text2 = GetHtmlStr(Text1)
End Sub

Private Function GetHtmlStr(strUrl As String) As String
    Dim xml As Object
    Set xml = CreateObject("msxml2.serverxmlhttp")
    xml.Open "GET", strUrl, False
    xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xml.setRequestHeader "Accept-Language", "zh-cn"
'    xml.setRequestHeader "Accept-Encoding", "gzip, deflate"
    xml.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; SV1; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.04506.648; .NET CLR 3.5.21022; .NET CLR 3.0.4506.2152; .NET CLR 3.5.30729)"
    xml.send
    Do While xml.ReadyState <> 4
        DoEvents
    Loop
    GetHtmlStr = StrConv(xml.ResponseBody, vbUnicode)
    Set xml = Nothing
End Function


 

Private Sub Command1_Click()
    
    If Text1.Text = "" Then
        'MsgBox "请输入正确的网址", , "错误！"
        Text1.SetFocus
    Else
        WebBrowser1.Silent = True
        WebBrowser1.Navigate Text1.Text
        
    End If

End Sub

Private Sub Form_Load()
    filename = ""
End Sub

Private Sub Timer1_Timer()
    Call Command1_Click
    Timer1.Enabled = False
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)

'End Sub
'
''Downloads By http://www.veryhuo.com
'Private Sub WebBrowser1_DownloadComplete()
    On Error Resume Next
  
    Dim Cont, Rcont, NextPage
    Dim Pos, posL, posR
    Dim DocTitle
    
    If URL = vbNullString Then
        Exit Sub
    End If
    
    If WebBrowser1.ReadyState <> READYSTATE_COMPLETE Then
        Exit Sub
    End If
    
    'cont = WebBrowser1.Document.documentElement.outertext
    Cont = WebBrowser1.Document.documentElement.outerHTML

    If Err Then
        Err.Clear
        Exit Sub
    End If


    Cont = LCase(Cont)
    Cont = Replace$(Cont, " ", "")
    
    '寻找标题
    Pos = InStr(1, Cont, "htmltimu")
    
    If Pos = 0 Then

        Exit Sub

    End If

    Cont = Mid$(Cont, Pos)
    posL = InStr(1, Cont, ">") + 1
    posR = InStr(1, Cont, "<")
    Rcont = Mid$(Cont, posL, posR - posL)
    If filename = "" Then
        filename = Rcont
    End If
    Rcont = Rcont & vbCrLf
    Text2 = Rcont & "  " & Text1.Text & vbCrLf & Text2
   
    '寻找内容
    Pos = InStr(1, Cont, "trail")

    If Pos = 0 Then

        Exit Sub

    End If

    Cont = Mid$(Cont, Pos)
    '继续寻找
  
'    Text2 = cont
   
    posL = InStr(1, (Cont), "<p>")
    posR = InStr(posL, (Cont), "</div>")

    Rcont = Rcont & Mid$(Cont, posL, posR - posL)
    
    Cont = Mid$(Cont, posR)
    
    
    Rcont = Replace$(Rcont, " ", "")
    Rcont = Replace$(Rcont, "　", "")
    Rcont = Replace(Rcont, "&nbsp;", "")
    
    Dim a
    a = InStr(1, Rcont, "<p>本书来自品书网")
    If a Then
        Rcont = Left$(Rcont, a - 1)
    End If
    
    Dim b
    Dim strCut
    a = InStr(1, Rcont, "请大家搜索")
    If a Then
        b = InStr(a, Rcont, "更新最快的小说")
        If b Then '
            strCut = Mid$(Rcont, a, b + 7 - a)
            Rcont = Replace$(Rcont, strCut, "")
        End If
    End If
    
    
    Rcont = Replace$(Rcont, "<p>", "")
    Rcont = Replace$(Rcont, "</p>", "")
    Rcont = Replace$(Rcont, "品书网", "")
    Rcont = Replace$(Rcont, "www.vodtw.com", "")
    Rcont = Replace$(Rcont, "复制网址访问", "")
    Rcont = Replace$(Rcont, "http://%77%77%77%2e%76%6f%64%74%77%2e%63%6f%6d", "")
    Rcont = Replace$(Rcont, "（）", "")
    Rcont = Rcont & vbCrLf & vbCrLf
    
    '
    Dim i As Integer
    i = FreeFile
    Open App.Path & "\" & filename & ".txt" For Append As #i
    Print #i, Rcont
    Close #i
    '
    
     
'    Text3 = zzzzz
'    Text2 = cont
    
    posL = InStr(1, Cont, "下一页")
    posR = InStrRev(Cont, """", posL)
    posL = InStrRev(Cont, """", posR - 1) + 1
    NextPage = Mid$(Cont, posL, posR - posL)
    
    
    If Right$(txtHttpWww, 1) <> "/" Then
        txtHttpWww = txtHttpWww & "/"
    End If
    
    Text1.Text = txtHttpWww & NextPage

    If chk.Value = vbChecked Then
        Timer1.Enabled = True
        Text3 = Rcont
    End If

End Sub
