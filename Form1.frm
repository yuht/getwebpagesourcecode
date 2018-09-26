VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "信息查询"
   ClientHeight    =   8325
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   14850
   FillColor       =   &H80000005&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   14850
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      ForeColor       =   &H80000007&
      Height          =   2490
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5640
      Width           =   14610
   End
   Begin VB.CommandButton Command1 
      Caption         =   "搞一下上面的数据"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   5040
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      ForeColor       =   &H80000007&
      Height          =   2970
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":0000
      Top             =   2040
      Width           =   14610
   End
   Begin VB.CommandButton cmdgetmenu 
      Caption         =   "点击查询"
      Height          =   300
      Left            =   7920
      TabIndex        =   3
      Top             =   315
      Width           =   990
   End
   Begin VB.TextBox txtHttpWww 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   1470
      TabIndex        =   0
      Text            =   "略阳县东信矿业有限责任公司"
      Top             =   360
      Width           =   6225
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      ForeColor       =   &H80000007&
      Height          =   930
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   8850
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "https://www.qichacha.com/search?key="
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "再下面粘贴从 http://www.gpsspg.com/latitude-and-longitude.htm 导出的txt文本"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   7695
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查询企业名称"
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   2
      Top             =   375
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit



Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001
 
Dim FileName
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
'    strHTML = GetHtmlStr(Text1)
    strHTML = LCase$(strHTML)
    '====================================获取章节标题
    tmpstrTitle = GetTitle(strHTML)
    
    If Len(tmpstrTitle) = 0 Then
        Exit Sub
    End If
    
    tmpstrTitle = Trim$(tmpstrTitle)
    '将章节标题作为文件名
    If FileName = "" Then
        FileName = tmpstrTitle
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

    Call fileWrite(App.Path & "\" & FileName & ".txt", strCont)
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
    'Text1.Text = txtHttpWww & tmpstrCont

'    Text2 = tmpstrTitle & vbTab & vbTab & Text1 & vbCrLf & Text2

'    If chk.Value = vbChecked Then
'        Timer1.Enabled = True
'    End If
    
End Sub

Private Function GetHtmlStr(strUrl As String) As String
    Dim Vera
    Dim k() As Byte
    Dim xml As Object
    Set xml = CreateObject("msxml2.serverxmlhttp")
    xml.Open "GET", strUrl, False
'        xml.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    'xml.setRequestHeader "Accept-Language", "zh-cn"
'        xml.setRequestHeader "Accept-Encoding", "gzip, deflate"
    xml.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; SV1; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.04506.648; .NET CLR 3.5.21022; .NET CLR 3.0.4506.2152; .NET CLR 3.5.30729)"
    xml.send

    Do While xml.ReadyState <> 4
        DelayMs (100)
    Loop
     
'    GetHtmlStr = StrConv(xml.ResponseBody, vbUnicode)
'    Call fileWrite(App.Path & "/1-1loadsourcecode.log", "去掉""  ""两个空格之后的原始信息:" & vbCrLf & GetHtmlStr)
    
    GetHtmlStr = UTF8ToGB2312(xml.ResponseBody)
    'Call fileWrite(App.Path & "/1-2loadsourcecode.log", "去掉""  ""两个空格之后的原始信息:" & vbCrLf & GetHtmlStr)
 
    Set xml = Nothing
End Function

Public Function UTF8ToGB2312(ByVal varIn As Variant) As String
    Dim bytesData() As Byte
    Dim adoStream As Object

    bytesData = varIn
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = "utf-8"
    adoStream.Type = 1 'adTypeBinary
    adoStream.Open
    adoStream.Write bytesData
    adoStream.Position = 0
    adoStream.Type = 2 'adTypeText
    UTF8ToGB2312 = adoStream.ReadText()
    adoStream.Close
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
 
Private Function GetTitles(strHTML As String) As String
    Dim Pos, Cont, posL, posR
    '寻找内容
    Pos = InStr(1, strHTML, "class=""bookname"">")

    If Pos = 0 Then
        Exit Function
    End If
    
    Pos = InStr(Pos, strHTML, "<h1>")

    If Pos = 0 Then
        Exit Function
    End If
    Cont = Mid$(strHTML, Pos + 4)
    '继续寻找
  
    '    Text2 = cont
   
    'posL = InStr(1, Cont, "content"">") + 9
    posR = InStr(1, Cont, "</h1>")
    
    If posR = 0 Then
        Exit Function
    End If

'    GetContent = Mid$(Cont, posL, posR - posL)
    GetTitles = Trim(Left(Cont, posR - 1))
    
End Function

 
 
Private Function GetContent(strHTML As String) As String
    Dim Pos, Cont, posL, posR
    '寻找内容
    Pos = InStr(1, strHTML, "id=""content"">") + 13

    If Pos = 0 Then
        Exit Function
    End If

    Cont = Mid$(strHTML, Pos)
    '继续寻找
  
    '    Text2 = cont
   
    'posL = InStr(1, Cont, "content"">") + 9
    posR = InStr(1, Cont, "</div>") - 1
    
    If posR = -1 Then
        Exit Function
    End If

'    GetContent = Mid$(Cont, posL, posR - posL)
    GetContent = Left(Cont, posR)
    'Call fileWrite(App.Path & "/debug_getcontent.log", CStr(GetContent))
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
    strContent = Replace(strContent, " ", "")
    strContent = Replace(strContent, "　", "")
    strContent = Replace(strContent, "‘", "")
    strContent = Replace(strContent, "’", "")
    strContent = Replace(strContent, "&nbsp;", "")
    'strContent = Replace(strContent, "?", "")
    strContent = Replace(strContent, "<br/>", vbCrLf)
    'strContent = Replace(strContent, vbCr, vbCrLf)
    'strContent = Replace(strContent, vbLf, vbCrLf)
    strContent = Replace(strContent, vbCrLf & vbCrLf, vbCrLf)
    strContent = Replace(strContent, vbCrLf & vbCrLf, vbCrLf)
    strContent = Replace(strContent, vbCrLf & vbCrLf, vbCrLf)
    
    strContent = Replace(strContent, "<script>readx();</script>", "")

    
    
    Dim tmpAD
    Dim a
    a = InStr(1, strContent, "<ahref")

    If a Then
        tmpAD = Mid(strContent, a - 1)

        Dim b
        b = InStr(1, tmpAD, "</a>")
        If b Then '
            tmpAD = Left$(tmpAD, b + 4)
            strContent = Replace$(strContent, tmpAD, "")
        End If
    End If
    
'    strContent = Replace$(strContent, "<p>", "")
'    'strContent = Replace$(strContent, "</p>", "")
'    strContent = Replace$(strContent, "</p>", vbCrLf)
     
    
'    strContent = Replace$(strContent, "壹看书ｗｗ看ｗ·１ｋａｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "壹看书ｗｗｗ看·１ｋａｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "要看书ｗ书ｗｗ·１ｋａｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "一看书ｗｗｗ·１ｋａｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "一看书ｗｗｗ·１ｋａ要ｎ书ｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "要看书ｗｗｗ·１ｋａ书ｎｓｈｕ·ｃｃ ", "")
'    strContent = Replace$(strContent, "要看书ｗｗｗ·１ｋａｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "壹看书ｗｗｗ·１ｋａｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "要看书ｗｗｗ·１书ｋａｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "壹看书ｗｗｗ·１ｋａ看ｎｓｈｕ看·ｃｃ", "")
'    strContent = Replace$(strContent, "一看书ｗｗｗ要·１要ｋａｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "壹看书ｗｗｗ·１ｋ要ａｎｓ看ｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "一看书ｗｗｗ·１ｋａｎｓ书ｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "壹看书ｗｗｗ书·１ｋａｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "要看书ｗｗ要ｗ·１ｋａ书ｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "壹看书ｗ?ｗｗ?·１?ｋａｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "一看书?ｗ?ｗｗ·１ｋａｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "要看书ｗ?ｗ?ｗ·１ｋａｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "一看书ｗ?ｗｗ·１ｋａｎｓｈｕ·ｃｃ", "")
'    strContent = Replace$(strContent, "一看书ｗ?ｗｗ·１ｋａｎｓｈｕ·ｃｃ", "")

'
'    strContent = Replace$(strContent, "938小说网www.938xs.com", "")
'    strContent = Replace$(strContent, "http://www.938xs.com", "")
'    strContent = Replace$(strContent, "938小说网", "")
'    strContent = Replace$(strContent, "本书来自", "")
    
     

    
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



 Public Function UTF8Encode(ByVal szInput As String) As String
    Dim wch  As String
    Dim uch As String
    Dim szRet As String
    Dim x As Long
    Dim inputLen As Long
    Dim nAsc  As Long
    Dim nAsc2 As Long
    Dim nAsc3 As Long
     
    If szInput = "" Then
        UTF8Encode = szInput
        Exit Function
    End If
    inputLen = Len(szInput)
    For x = 1 To inputLen
    '得到每个字符
        wch = Mid(szInput, x, 1)
        '得到相应的UNICODE编码
        nAsc = AscW(wch)
    '对于<0的编码　其需要加上65536
        If nAsc < 0 Then nAsc = nAsc + 65536
    '对于<128位的ASCII的编码则无需更改
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
            '真正的第二层编码范围为000080 - 0007FF
            'Unicode在范围D800-DFFF中不存在任何字符，基本多文种平面中约定了这个范围用于UTF-16扩展标识辅助平面（两个UTF-16表示一个辅助平面字符）.
            '当然，任何编码都是可以被转换到这个范围，但在unicode中他们并不代表任何合法的值。
     
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
                 
            Else
            '第三层编码00000800 – 0000FFFF
            '首先取其前四位与11100000进行或去处得到UTF-8编码的前8位
            '其次取其前10位与111111进行并运算，这样就能得到其前10中最后6位的真正的编码　再与10000000进行或运算来得到UTF-8编码中间的8位
            '最后将其与111111进行并运算，这样就能得到其最后6位的真正的编码　再与10000000进行或运算来得到UTF-8编码最后8位编码
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
     
    UTF8Encode = szRet
End Function



Private Sub cmdgetmenu_Click()

    
    
    Dim strMenulist
    Dim posL, posR
        
    '==================================== 对书籍首页网址进行格式化
'    If Right$(txtHttpWww, 1) <> "/" Then
'        txtHttpWww = txtHttpWww & "/"
'    End If
    strMenulist = GetHtmlStr("https://www.qichacha.com/search?key=" + UTF8Encode(txtHttpWww))
    
    'strMenulist = Text4
'    Call fileWrite(App.Path & "/1loadsourcecode.log", "去掉""  ""两个空格之后的原始信息:" & vbCrLf & strMenulist)
    strMenulist = LCase$(strMenulist)
    strMenulist = Replace$(strMenulist, "  ", "")
    strMenulist = Replace$(strMenulist, "&nbsp;", "")
    strMenulist = Replace$(strMenulist, vbTab, "")
    strMenulist = Replace$(strMenulist, "<br/>", vbCrLf)
    
    strMenulist = Replace$(strMenulist, vbCrLf & vbCrLf, vbCrLf)
    strMenulist = Replace$(strMenulist, vbCrLf & " ", vbCrLf)
    
    'Call fileWrite(App.Path & "/debug_strmenu.log", "strMenulist" & vbCrLf & vbCrLf & "-------------------------------------" & vbCrLf & strMenulist)
    'Call fileWrite(App.Path & "/debug_replaceSPACE.log", "去掉""  ""两个空格之后的原始信息:" & vbCrLf & strMenulist)
    
    posL = InStr(1, strMenulist, "小查为您找到")
    If posL = 0 Then
        Exit Sub
    End If
    
    
    
    posR = InStr(posL, strMenulist, "点击搜索更多")
    If posL = 0 Then
        Exit Sub
    End If
    
    Dim i
    Dim j
    Dim k
    
      
    '截取企业信息
    BookTitle = Mid(strMenulist, posL + 1, posR - posL - 1)
    
    '过滤全部标签
    Do
        i = InStr(1, BookTitle, "<")
        If i Then
            j = InStr(i, BookTitle, ">")
            If j Then
                BookTitle = Replace(BookTitle, Mid(BookTitle, i, j - i + 1), "")
                 
            End If
             
        End If
    Loop While (j <> 0 And i <> 0)
    
    '清理无用字符
    BookTitle = Replace(BookTitle, vbCrLf, "")
    BookTitle = Replace(BookTitle, vbCr, "")
    BookTitle = Replace(BookTitle, vbLf, "")
    BookTitle = Replace(BookTitle, "  ", " ")
    
    
    '显示有效数据
    Text2 = BookTitle
    
    posL = InStr(1, BookTitle, "地址：")
    
    If (posL) Then
        posR = InStr(posL, BookTitle, " ")
        If posR Then
            BookTitle = Mid(BookTitle, posL, posR - posL + 1)
        End If
    End If
    
    BookTitle = Replace(BookTitle, " ", "")
    Text2 = BookTitle
    
    Exit Sub
'
'    'Call fileWrite(App.Path & "/debug_Booktitle.log", "Book Title:" & BookTitle)
'    '将章节标题作为文件名
'    If FileName = "" Then
'        FileName = BookTitle 'tmpstrTitle
'    End If
'
'
'
'    Exit Sub
'
'    '截取章节信息
'
'
'    posL = InStr(1, strMenulist, "</dt>")
'    If posL = 0 Then
'        Exit Sub
'    End If
'
'    posL = InStr(posL + 1, strMenulist, "</dt>")
'    If posL = 0 Then
'        Exit Sub
'    End If
'
'    posL = InStr(posL + 1, strMenulist, "<dd><a")
'    If posL = 0 Then
'        Exit Sub
'    End If
'
'
'
'    strMenulist = Mid$(strMenulist, posL)
'    'Call fileWrite(App.Path & "/debug_cap2.log", "截取<dd><a后的信息:" & vbCrLf & strMenulist)
'
'
'    posR = InStr(1, strMenulist, "</dl>")
'    If posR = 0 Then
'        Exit Sub
'    End If
'
'
'
'    strMenulist = Left(strMenulist, posR - 1)
'    'Call fileWrite(App.Path & "/debug_ul_eul.log", "截取<ul></ul>之间的信息:" & vbCrLf & strMenulist)
'
''    strMenulist = Replace$(strMenulist, "<ul>", "")
''    strMenulist = Replace$(strMenulist, "</ul>", "")
'    strMenulist = Replace$(strMenulist, "<dd>", "")
'    strMenulist = Replace$(strMenulist, "</dd>", "")
''    strMenulist = Replace$(strMenulist, "<span></span>", "")
''    strMenulist = Replace$(strMenulist, "-" & BookTitle, "")
''    strMenulist = Replace$(strMenulist, vbCr, "")
''    strMenulist = Replace$(strMenulist, vbLf, "")
'
'    'Call fileWrite(App.Path & "/debug_dd.log", "去掉<ul></ul><li></li>标记的信息:" & vbCrLf & strMenulist)
'
'
'
'    Dim hostUrl
'    Dim TotalCap
'    Dim s As String
'    k = Split(strMenulist, "</a>")
'
''    For i = 0 To TotalCap - 1
''        For j = i + 1 To TotalCap
''            If k(i) = k(j) Then
''                k(j) = ""
''            End If
''            DoEvents
''        Next
''        DoEvents
''    Next
'
'
'    TotalCap = UBound(k) - 2
'
'    ReDim UT(TotalCap) As UrlandTitle
'
'    Text2 = "一共找到 " & TotalCap & " 章" & vbCrLf & Text2
'
'    j = Val(txtText3) - 1
'    If j < 0 Then j = 0
'
'    txtText3 = j + 1
'    hostUrl = Left(txtHttpWww, InStr(10, txtHttpWww, "/") - 1)
'
'    For i = j To TotalCap
'        If Len(k(i)) Then
'            k(i) = Trim$(k(i))
'            posL = InStr(1, k(i), """") + 1
'            If posL <> 0 Then
'                posR = InStr(posL, k(i), """") - 1
'                 If posR <> 0 Then
'                    UT(i).Url = hostUrl & Mid$(k(i), posL, posR - posL + 1)
'
'                End If
'            End If
'            If Len(UT(i).Url) Then
'                posL = InStrRev(k(i), ">") + 1
'                If posL <> 0 Then
'                    UT(i).Title = Mid$(k(i), posL)
'                End If
'            End If
'            'Call fileWrite(App.Path & "/debug_Totalcaps.log", "第 " & i & vbTab & " 章:" & UT(i).Title & " - " & UT(i).Url & vbCrLf)
'
'        End If
'        DoEvents
'    Next
'
'
'
'
'
'
'    For i = j To UBound(UT)
'        DoEvents
''        If i < 1763 Then
''            UT(i).Title = "第" & i + 1 & "章 " & UT(i).Title
''        ElseIf i = 1763 Then
''            UT(i).Title = "关于更新想说的话"
''        Else
''            UT(i).Title = "第" & i & "章 " & UT(i).Title
'            UT(i).Title = UT(i).Title
''        End If
'
'        Me.Caption = BookTitle & "   " & i & "/" & UBound(k) - 1
'        Call SaveContent(UT(i))
'
'    Next
'    Text2 = "下载完毕!" & vbCrLf & Text2
End Sub

Private Function SaveContent(UT As UrlandTitle)
    Dim strCont     As String
    Dim tmpstrCont  As String
    Dim tmpstrTitle As String
    Dim strHTML     As String

    
    '=====================================获取html内容
    strHTML = GetHtmlStr(UT.Url)
    strHTML = LCase$(strHTML)
    
    'Call fileWrite(App.Path & "/debug_cont.log", UT.Url & vbCrLf & UT.Title & strHTML)
    
    '====================================获取章节标题
        
    tmpstrTitle = GetTitles(strHTML) 'UT.Title
    

    
    '====================================获取章节内容
    tmpstrCont = GetContent(strHTML)

    
    If Len(tmpstrCont) = 0 Then
        Exit Function
    End If
    'Call fileWrite(App.Path & "/debug_cont.log", UT.Url & vbCrLf & UT.Title & tmpstrCont)
    tmpstrCont = ContentUnescape(tmpstrCont)
    tmpstrCont = ContentFilter(tmpstrCont)
    
    strCont = tmpstrTitle & vbCrLf & tmpstrCont & vbCrLf & vbCrLf

    Call fileWrite(App.Path & "\" & FileName & ".txt", strCont)
 
    Text2 = UT.Url & vbTab & tmpstrTitle & vbCrLf & Text2
    
End Function


Private Sub Command1_Click()
Dim k
Dim l
Dim i
k = Split(Text1, vbCrLf)
For i = 0 To UBound(k)
    Debug.Print "k(" & i & "):" & k(i)
    l = Split(k(i), ",")
    If UBound(l) > 5 Then
        Text3 = Text3 & l(5) & "," & l(4) & vbCrLf
    End If
Next
End Sub

Private Sub Form_Load()
    FileName = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

'Private Sub Timer1_Timer()
'    Timer1.Enabled = False
'    Call cmdCommand2_Click
'End Sub


