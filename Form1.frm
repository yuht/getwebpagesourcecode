VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   Caption         =   "获取网页源文件"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12030
   FillColor       =   &H80000005&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   12030
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "一直自动获取"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   10215
      TabIndex        =   6
      Top             =   270
      Width           =   1725
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
      TabIndex        =   5
      Top             =   4635
      Width           =   12735
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   900
      Left            =   90
      TabIndex        =   4
      Top             =   675
      Width           =   11865
      ExtentX         =   20929
      ExtentY         =   1587
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
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "获取"
      Default         =   -1  'True
      Height          =   375
      Left            =   9180
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   225
      Width           =   855
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H80000007&
      Height          =   2505
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1665
      Width           =   11925
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   945
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "http://www.vodtw.com/html/book/28/28902/21716386.html"
      Top             =   225
      Width           =   8070
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "输入网址："
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   270
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim filename

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
    Dim Pos, posL, posR, posP, posDiv
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
   
    posP = InStr(1, (Cont), "<p>")
    posDiv = InStr(posP, (Cont), "</div>")

    Rcont = Rcont & Mid$(Cont, posP, posDiv - posP)
    
    Cont = Mid$(Cont, posDiv)
    
    
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
    
    Text1.Text = "http://www.vodtw.com/html/book/28/28902/" & NextPage

    If chk.Value = vbChecked Then
        Timer1.Enabled = True
        Text3 = Rcont
    End If

End Sub
