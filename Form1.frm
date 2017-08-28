VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   Caption         =   "www.vodtw.com在线小说下载器"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   FillColor       =   &H80000005&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   10995
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtText3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8100
      TabIndex        =   9
      Text            =   "2000"
      Top             =   90
      Width           =   780
   End
   Begin VB.CommandButton cmdgetmenu 
      Caption         =   "下载"
      Height          =   300
      Left            =   9945
      TabIndex        =   7
      Top             =   90
      Width           =   990
   End
   Begin VB.CommandButton cmdCommand2 
      Caption         =   "下载"
      Height          =   315
      Left            =   7785
      TabIndex        =   6
      Top             =   630
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
      TabIndex        =   3
      Top             =   630
      Value           =   1  'Checked
      Width           =   1950
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2385
      Top             =   1260
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      ForeColor       =   &H80000007&
      Height          =   2415
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1125
      Width           =   10890
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   1350
      TabIndex        =   2
      Text            =   "http://www.vodtw.com/html/book/28/28902/21716386.html"
      Top             =   630
      Width           =   6225
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "从第          章开始"
      Height          =   180
      Index           =   1
      Left            =   7695
      TabIndex        =   8
      Top             =   135
      Width           =   1800
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "书记目录网址："
      Height          =   180
      Index           =   0
      Left            =   45
      TabIndex        =   5
      Top             =   135
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "当前阅读网址："
      Height          =   255
      Left            =   45
      TabIndex        =   4
      Top             =   675
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
Dim BookTitle
    
    
Private Type UrlandTitle
    Url As String
    Title As String
End Type

Private Declare Function timeGetTime Lib "winmm.dll" () As Long '该声明得到系统开机到现在的时间(单位：毫秒)

Public Function DelayMs(T As Long)
    Dim Savetime As Long
    Savetime = timeGetTime '记下开始时的时间
    While timeGetTime < Savetime + T '循环等待
        DoEvents '转让控制权
    Wend
End Function



Private Sub cmdCommand2_Click()
    Dim strCont     As String
    Dim tmpstrCont  As String
    Dim tmpstrTitle As String
    Dim strHTML     As String
    
    '=====================================获取html内容
    strHTML = GetHtmlStr(Text1)
    strHTML = LCase$(strHTML)
    '====================================获取章节标题
    tmpstrTitle = GetTitle(strHTML)
    
    If Len(tmpstrTitle) = 0 Then
        Exit Sub
    End If
    
    tmpstrTitle = Trim$(tmpstrTitle)
    '将章节标题作为文件名
    If filename = "" Then
        filename = tmpstrTitle
    End If
    
    '====================================获取章节内容
    tmpstrCont = GetContent(strHTML)

    '    Text3 = tmpstrCont
    If Len(tmpstrCont) = 0 Then
        Exit Sub
    End If
    
    tmpstrCont = ContentUnescape(tmpstrCont)
    tmpstrCont = ContentFilter(tmpstrCont)
    
    strCont = tmpstrTitle & vbCrLf & tmpstrCont & vbCrLf & vbCrLf

    Call fileWrite(App.Path & "\" & filename & ".txt", strCont)
    '==================================== 获取下一章对应的rul
    tmpstrCont = GetNextUrl(strHTML)
 
    If Len(tmpstrCont) = 0 Then
        Exit Sub
    End If
    
    '==================================== 对书籍首页网址进行格式化
    If Right$(txtHttpWww, 1) <> "/" Then
        txtHttpWww = txtHttpWww & "/"
    End If
    
    '==================================== 下一章的网址，填入文本框
    Text1.Text = txtHttpWww & tmpstrCont

    Text2 = tmpstrTitle & vbTab & vbTab & Text1 & vbCrLf & Text2

    If chk.Value = vbChecked Then
        Timer1.Enabled = True
    End If
    
End Sub

Private Function GetHtmlStr(strUrl As String) As String
    Dim xml As Object
    Set xml = CreateObject("msxml2.serverxmlhttp")
    xml.Open "GET", strUrl, False
    '    xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xml.setRequestHeader "Accept-Language", "zh-cn"
    '    xml.setRequestHeader "Accept-Encoding", "gzip, deflate"
    xml.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; SV1; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.04506.648; .NET CLR 3.5.21022; .NET CLR 3.0.4506.2152; .NET CLR 3.5.30729)"
    xml.send

    Do While xml.ReadyState <> 4
        DelayMs (30)
    Loop

    GetHtmlStr = StrConv(xml.ResponseBody, vbUnicode)
    'GetHtmlStr = xml.responsetext
    Set xml = Nothing
End Function

Private Function GetTitle(strHTML As String) As String

    Dim Pos, Cont, posL, posR
    '寻找标题
    Pos = InStr(1, strHTML, "htmltimu")
    
    If Pos = 0 Then

        Exit Function

    End If

    Cont = Mid$(strHTML, Pos)
    posL = InStr(1, Cont, ">") + 1
    posR = InStr(posL, Cont, "<")
    GetTitle = Mid$(Cont, posL, posR - posL)
    '寻找标题结束

End Function
 
Private Function GetContent(strHTML As String) As String
    Dim Pos, Cont, posL, posR
    '寻找内容
    Pos = InStr(1, strHTML, "trail")

    If Pos = 0 Then

        Exit Function

    End If

    Cont = Mid$(strHTML, Pos)
    '继续寻找
  
    '    Text2 = cont
   
    posL = InStr(1, Cont, "<p>")
    posR = InStr(posL, Cont, "</div>")

    GetContent = Mid$(Cont, posL, posR - posL)
    
    'Cont = Mid$(Cont, posR)
    '寻找内容结束

End Function
 
Private Function ContentUnescape(strContent As String) As String
    Dim i As Double, j As Double, k As Integer
    Dim tmpStr
    Dim tmpNum
    j = 1

    Do
        DoEvents
        j = InStr(j, strContent, "&#")

        If j > 0 Then
            k = InStr(j, strContent, ";")

            If k > 0 Then
                tmpStr = Mid$(strContent, j, k - j + 1)
                tmpNum = Hex(Val(Mid$(tmpStr, 3, Len(tmpStr) - 3)))
                strContent = Replace(strContent, tmpStr, ChrW(CLng("&h" & tmpNum)))
            End If
        End If

    Loop While j
    
    ContentUnescape = strContent
    
End Function
 
Private Function ContentFilter(strContent As String) As String
    '内容过滤
    Dim strCut As String
    strContent = Replace$(strContent, " ", "")
    strContent = Replace$(strContent, "　", "")
    strContent = Replace(strContent, "&nbsp;", "")
    
    Dim a
    a = InStr(1, strContent, "<p>本书来自品书网")

    If a Then
        strContent = Left$(strContent, a - 1)
    End If
    
    Dim b
    a = InStr(1, strContent, "请大家搜索")

    If a Then
        b = InStr(a, strContent, "更新最快的小说")

        If b Then '
            strCut = Mid$(strContent, a, b + 7 - a)
            strContent = Replace$(strContent, strCut, "")
        End If
    End If
    
    strContent = Replace$(strContent, "<p>", "")
    'strContent = Replace$(strContent, "</p>", "")
    strContent = Replace$(strContent, "</p>", vbCrLf)
     
    
    strContent = Replace$(strContent, "如您已阅读到此章节，请移步到:新匕匕奇中文小說xinыqi.com阅读最新章节", "")
    strContent = Replace$(strContent, "http://%77%77%77%2e%76%6f%64%74%77%2e%63%6f%6d", "")
    strContent = Replace$(strContent, "敬请记住我们的网址:匕匕奇小說xinыqi.com。", "")
    strContent = Replace$(strContent, "新·匕匕·奇·中·文·网·首·发xin", "")
    strContent = Replace$(strContent, "[就上+新^^匕匕^^奇^^中^^文^^网+", "")
    strContent = Replace$(strContent, "新匕匕·奇·中·文·蛧·首·发", "")
    strContent = Replace$(strContent, "（閱讀最新章節首发.com）", "")
    strContent = Replace$(strContent, "更多精彩小说请访问.com", "")
    strContent = Replace$(strContent, "新匕匕奇新地址：www.m", "")
    strContent = Replace$(strContent, "｛新匕匕奇中文小說m｝", "")
    strContent = Replace$(strContent, "www.xinbiqi.com", "")
    strContent = Replace$(strContent, "http：／／xin／", "")
    strContent = Replace$(strContent, "www.vodtw.com", "")
    strContent = Replace$(strContent, "www.vodtw.net", "")
    strContent = Replace$(strContent, "www.xinbiqi.", "")
    strContent = Replace$(strContent, "复制网址访问", "")
    strContent = Replace$(strContent, "新比奇中文网", "")
    strContent = Replace$(strContent, ".xinыqi.com", "")
    strContent = Replace$(strContent, "品书网", "")
    strContent = Replace$(strContent, "（）", "")
    strContent = Replace$(strContent, "()", "")
    
     

    
    ContentFilter = strContent
End Function
 
Private Function GetNextUrl(strHTML As String) As String
    Dim posL, posR
    posL = InStr(1, strHTML, "下一页")
    posR = InStrRev(strHTML, """", posL)
    posL = InStrRev(strHTML, """", posR - 1) + 1
    GetNextUrl = Mid$(strHTML, posL, posR - posL)
End Function

'App.Path & "\" & filename & ".txt"
Function fileWrite(strFilename As String, strContent As String)
    '写入文件
    Dim i As Integer
    i = FreeFile
    Open strFilename For Append As #i
    Print #i, strContent
    Close #i
End Function

Private Sub cmdgetmenu_Click()

    
    
    Dim strMenulist
    Dim posL, posR
        
    '==================================== 对书籍首页网址进行格式化
    If Right$(txtHttpWww, 1) <> "/" Then
        txtHttpWww = txtHttpWww & "/"
    End If
    
    
    
    strMenulist = GetHtmlStr(txtHttpWww)
    strMenulist = LCase$(strMenulist)
    
    strMenulist = Replace$(strMenulist, "  ", "")
    strMenulist = Replace$(strMenulist, vbTab, "")
    
    'Call fileWrite(App.Path & "/debug.log", "去掉""  ""两个空格之后的原始信息:" & vbCrLf & strMenulist)
    
    posL = InStr(1, strMenulist, "bookname")
    If posL = 0 Then
        Exit Sub
    End If
    
    posL = InStr(posL, strMenulist, "<h1>")
    If posL = 0 Then
        Exit Sub
    End If
    
    posL = InStr(posL, strMenulist, ">")
    If posL = 0 Then
        Exit Sub
    End If
    
    posR = InStr(posL, strMenulist, "<")
    If posL = 0 Then
        Exit Sub
    End If
    
    BookTitle = Mid(strMenulist, posL + 1, posR - posL - 1)
    
    Text2 = "书名:" & BookTitle
    
    
    'Call fileWrite(App.Path & "/debug.log", "Book Title:" & BookTitle)
    '将章节标题作为文件名
    If filename = "" Then
        filename = BookTitle 'tmpstrTitle
    End If
    
    
    posL = InStr(1, strMenulist, "<dd>")
    If posL = 0 Then
        Exit Sub
    End If
    
    strMenulist = Mid$(strMenulist, posL)
    'Call fileWrite(App.Path & "/debug.log", "截取dd后的信息:" & vbCrLf & strMenulist)
    
    posL = InStr(1, strMenulist, "<ul>")
    If posL = 0 Then
        Exit Sub
    End If
    
    strMenulist = Mid$(strMenulist, posL)
    'Call fileWrite(App.Path & "/debug.log", "截取<ul>后的信息:" & vbCrLf & strMenulist)
    
    
    posR = InStr(1, strMenulist, "</ul>")
    If posR = 0 Then
        Exit Sub
    End If
    
    strMenulist = Left(strMenulist, posR - 1)
    'Call fileWrite(App.Path & "/debug.log", "截取<ul></ul>之间的信息:" & vbCrLf & strMenulist)
    
    strMenulist = Replace$(strMenulist, "<ul>", "")
    strMenulist = Replace$(strMenulist, "</ul>", "")
    strMenulist = Replace$(strMenulist, "<li>", "")
    strMenulist = Replace$(strMenulist, "</li>", "")
    strMenulist = Replace$(strMenulist, vbCr, "")
    strMenulist = Replace$(strMenulist, vbLf, "")
    
    'Call fileWrite(App.Path & "/debug.log", "去掉<ul></ul><li></li>标记的信息:" & vbCrLf & strMenulist)
    
    Dim k
    Dim i, j
    Dim s As String
    k = Split(strMenulist, "</a>")
    i = UBound(k) - 1
    ReDim UT(i) As UrlandTitle
    
    Text2 = "一共找到 " & i & " 章" & vbCrLf & Text2
    
    j = Val(txtText3) - 1
    If j < 0 Then j = 0
     
    txtText3 = j + 1
    
    For i = j To UBound(k) - 1
        k(i) = Trim$(k(i))
        posL = InStr(1, k(i), """") + 1
        If posL <> 0 Then
            posR = InStr(posL, k(i), """") - 1
             If posR <> 0 Then
                UT(i).Url = txtHttpWww & Mid$(k(i), posL, posR - posL + 1)
            End If
        End If
        If Len(UT(i).Url) Then
            posL = InStrRev(k(i), " ") + 1
            If posL <> 0 Then
                UT(i).Title = Mid$(k(i), posL)
                's = s & i + 1 & "-" & UT(i).Url & "-" & UT(i).Title & vbCrLf
            End If
        End If
        
        DoEvents
    Next
    'Call fileWrite(App.Path & "/debuglist.log", s)
    
    'Exit Sub
    
    For i = j To UBound(UT)
        DoEvents
'        If i < 1763 Then
'            UT(i).Title = "第" & i + 1 & "章 " & UT(i).Title
'        ElseIf i = 1763 Then
'            UT(i).Title = "关于更新想说的话"
'        Else
            UT(i).Title = "第" & i & "章 " & UT(i).Title
'        End If
        
        Call SaveContent(UT(i))
        
    Next
    Text2 = "下载完毕!" & vbCrLf & Text2
End Sub

Private Function SaveContent(UT As UrlandTitle)
    Dim strCont     As String
    Dim tmpstrCont  As String
    Dim tmpstrTitle As String
    Dim strHTML     As String

    
    '=====================================获取html内容
    strHTML = GetHtmlStr(UT.Url)
    strHTML = LCase$(strHTML)
    
     'Call fileWrite(App.Path & "/debugcont.log", UT.Url & vbCrLf & UT.Title & strHTML)
    
    '====================================获取章节标题
        
    tmpstrTitle = UT.Title
    

    
    '====================================获取章节内容
    tmpstrCont = GetContent(strHTML)

    
    If Len(tmpstrCont) = 0 Then
        Exit Function
    End If
    
    tmpstrCont = ContentUnescape(tmpstrCont)
    tmpstrCont = ContentFilter(tmpstrCont)
    
    strCont = tmpstrTitle & vbCrLf & tmpstrCont & vbCrLf & vbCrLf

    Call fileWrite(App.Path & "\" & filename & ".txt", strCont)
 
    Text2 = UT.Url & vbTab & tmpstrTitle & vbCrLf & Text2
    
End Function


Private Sub Form_Load()
    filename = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call cmdCommand2_Click
End Sub

