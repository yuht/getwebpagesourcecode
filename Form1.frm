VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ϣ��ѯ"
   ClientHeight    =   8325
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   14850
   FillColor       =   &H80000005&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   14850
   StartUpPosition =   1  '����������
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
      Caption         =   "��һ�����������"
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
      Caption         =   "�����ѯ"
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
      Text            =   "�����ض��ſ�ҵ�������ι�˾"
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
      Caption         =   "������ճ���� http://www.gpsspg.com/latitude-and-longitude.htm ������txt�ı�"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   7695
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѯ��ҵ����"
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

Private Declare Function timeGetTime Lib "winmm.dll" () As Long '�������õ�ϵͳ���������ڵ�ʱ��(��λ������)

Public Function DelayMs(T As Long)
    Dim Savetime As Long
    Savetime = timeGetTime '���¿�ʼʱ��ʱ��
    While timeGetTime < Savetime + T 'ѭ���ȴ�
        DoEvents 'ת�ÿ���Ȩ
    Wend
End Function



Private Sub cmdCommand2_Click()
    Dim strCont     As String
    Dim tmpstrCont  As String
    Dim tmpstrTitle As String
    Dim strHTML     As String
    
    '=====================================��ȡhtml����
'    strHTML = GetHtmlStr(Text1)
    strHTML = LCase$(strHTML)
    '====================================��ȡ�½ڱ���
    tmpstrTitle = GetTitle(strHTML)
    
    If Len(tmpstrTitle) = 0 Then
        Exit Sub
    End If
    
    tmpstrTitle = Trim$(tmpstrTitle)
    '���½ڱ�����Ϊ�ļ���
    If FileName = "" Then
        FileName = tmpstrTitle
    End If
    
    '====================================��ȡ�½�����
    tmpstrCont = GetContent(strHTML)

    '    Text3 = tmpstrCont
    If Len(tmpstrCont) = 0 Then
        Exit Sub
    End If
    
    tmpstrCont = ContentUnescape(tmpstrCont)
    tmpstrCont = ContentFilter(tmpstrCont)
    
    strCont = tmpstrTitle & vbCrLf & tmpstrCont & vbCrLf & vbCrLf

    Call fileWrite(App.Path & "\" & FileName & ".txt", strCont)
    '==================================== ��ȡ��һ�¶�Ӧ��rul
    tmpstrCont = GetNextUrl(strHTML)
 
    If Len(tmpstrCont) = 0 Then
        Exit Sub
    End If
    
    '==================================== ���鼮��ҳ��ַ���и�ʽ��
    If Right$(txtHttpWww, 1) <> "/" Then
        txtHttpWww = txtHttpWww & "/"
    End If
    
    '==================================== ��һ�µ���ַ�������ı���
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
'    Call fileWrite(App.Path & "/1-1loadsourcecode.log", "ȥ��""  ""�����ո�֮���ԭʼ��Ϣ:" & vbCrLf & GetHtmlStr)
    
    GetHtmlStr = UTF8ToGB2312(xml.ResponseBody)
    'Call fileWrite(App.Path & "/1-2loadsourcecode.log", "ȥ��""  ""�����ո�֮���ԭʼ��Ϣ:" & vbCrLf & GetHtmlStr)
 
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
    'Ѱ�ұ���
    Pos = InStr(1, strHTML, "htmltimu")
    
    If Pos = 0 Then

        Exit Function

    End If

    Cont = Mid$(strHTML, Pos)
    posL = InStr(1, Cont, ">") + 1
    posR = InStr(posL, Cont, "<")
    GetTitle = Mid$(Cont, posL, posR - posL)
    'Ѱ�ұ������

End Function
 
Private Function GetTitles(strHTML As String) As String
    Dim Pos, Cont, posL, posR
    'Ѱ������
    Pos = InStr(1, strHTML, "class=""bookname"">")

    If Pos = 0 Then
        Exit Function
    End If
    
    Pos = InStr(Pos, strHTML, "<h1>")

    If Pos = 0 Then
        Exit Function
    End If
    Cont = Mid$(strHTML, Pos + 4)
    '����Ѱ��
  
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
    'Ѱ������
    Pos = InStr(1, strHTML, "id=""content"">") + 13

    If Pos = 0 Then
        Exit Function
    End If

    Cont = Mid$(strHTML, Pos)
    '����Ѱ��
  
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
    'Ѱ�����ݽ���

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
    '���ݹ���
    Dim strCut As String
    strContent = Replace(strContent, " ", "")
    strContent = Replace(strContent, "��", "")
    strContent = Replace(strContent, "��", "")
    strContent = Replace(strContent, "��", "")
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
     
    
'    strContent = Replace$(strContent, "Ҽ���������������������������", "")
'    strContent = Replace$(strContent, "Ҽ���������������������������", "")
'    strContent = Replace$(strContent, "Ҫ��������������������������", "")
'    strContent = Replace$(strContent, "һ�������������������������", "")
'    strContent = Replace$(strContent, "һ����������������Ҫ������������", "")
'    strContent = Replace$(strContent, "Ҫ��������������������������� ", "")
'    strContent = Replace$(strContent, "Ҫ�������������������������", "")
'    strContent = Replace$(strContent, "Ҽ�������������������������", "")
'    strContent = Replace$(strContent, "Ҫ��������������������������", "")
'    strContent = Replace$(strContent, "Ҽ���������������ῴ������������", "")
'    strContent = Replace$(strContent, "һ���������Ҫ����Ҫ������������", "")
'    strContent = Replace$(strContent, "Ҽ���������������Ҫ���󿴣�������", "")
'    strContent = Replace$(strContent, "һ���������������������������", "")
'    strContent = Replace$(strContent, "Ҽ����������顤��������������", "")
'    strContent = Replace$(strContent, "Ҫ�������Ҫ��������������������", "")
'    strContent = Replace$(strContent, "Ҽ�����?����?����?������������", "")
'    strContent = Replace$(strContent, "һ����?��?��������������������", "")
'    strContent = Replace$(strContent, "Ҫ�����?��?������������������", "")
'    strContent = Replace$(strContent, "һ�����?��������������������", "")
'    strContent = Replace$(strContent, "һ�����?��������������������", "")

'
'    strContent = Replace$(strContent, "938С˵��www.938xs.com", "")
'    strContent = Replace$(strContent, "http://www.938xs.com", "")
'    strContent = Replace$(strContent, "938С˵��", "")
'    strContent = Replace$(strContent, "��������", "")
    
     

    
    ContentFilter = strContent
End Function
 
Private Function GetNextUrl(strHTML As String) As String
    Dim posL, posR
    posL = InStr(1, strHTML, "��һҳ")
    posR = InStrRev(strHTML, """", posL)
    posL = InStrRev(strHTML, """", posR - 1) + 1
    GetNextUrl = Mid$(strHTML, posL, posR - posL)
End Function

'App.Path & "\" & filename & ".txt"
Function fileWrite(strFilename As String, strContent As String)
    'д���ļ�
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
    '�õ�ÿ���ַ�
        wch = Mid(szInput, x, 1)
        '�õ���Ӧ��UNICODE����
        nAsc = AscW(wch)
    '����<0�ı��롡����Ҫ����65536
        If nAsc < 0 Then nAsc = nAsc + 65536
    '����<128λ��ASCII�ı������������
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
            '�����ĵڶ�����뷶ΧΪ000080 - 0007FF
            'Unicode�ڷ�ΧD800-DFFF�в������κ��ַ�������������ƽ����Լ���������Χ����UTF-16��չ��ʶ����ƽ�棨����UTF-16��ʾһ������ƽ���ַ���.
            '��Ȼ���κα��붼�ǿ��Ա�ת���������Χ������unicode�����ǲ��������κκϷ���ֵ��
     
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
                 
            Else
            '���������00000800 �C 0000FFFF
            '����ȡ��ǰ��λ��11100000���л�ȥ���õ�UTF-8�����ǰ8λ
            '���ȡ��ǰ10λ��111111���в����㣬�������ܵõ���ǰ10�����6λ�������ı��롡����10000000���л��������õ�UTF-8�����м��8λ
            '�������111111���в����㣬�������ܵõ������6λ�������ı��롡����10000000���л��������õ�UTF-8�������8λ����
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
        
    '==================================== ���鼮��ҳ��ַ���и�ʽ��
'    If Right$(txtHttpWww, 1) <> "/" Then
'        txtHttpWww = txtHttpWww & "/"
'    End If
    strMenulist = GetHtmlStr("https://www.qichacha.com/search?key=" + UTF8Encode(txtHttpWww))
    
    'strMenulist = Text4
'    Call fileWrite(App.Path & "/1loadsourcecode.log", "ȥ��""  ""�����ո�֮���ԭʼ��Ϣ:" & vbCrLf & strMenulist)
    strMenulist = LCase$(strMenulist)
    strMenulist = Replace$(strMenulist, "  ", "")
    strMenulist = Replace$(strMenulist, "&nbsp;", "")
    strMenulist = Replace$(strMenulist, vbTab, "")
    strMenulist = Replace$(strMenulist, "<br/>", vbCrLf)
    
    strMenulist = Replace$(strMenulist, vbCrLf & vbCrLf, vbCrLf)
    strMenulist = Replace$(strMenulist, vbCrLf & " ", vbCrLf)
    
    'Call fileWrite(App.Path & "/debug_strmenu.log", "strMenulist" & vbCrLf & vbCrLf & "-------------------------------------" & vbCrLf & strMenulist)
    'Call fileWrite(App.Path & "/debug_replaceSPACE.log", "ȥ��""  ""�����ո�֮���ԭʼ��Ϣ:" & vbCrLf & strMenulist)
    
    posL = InStr(1, strMenulist, "С��Ϊ���ҵ�")
    If posL = 0 Then
        Exit Sub
    End If
    
    
    
    posR = InStr(posL, strMenulist, "�����������")
    If posL = 0 Then
        Exit Sub
    End If
    
    Dim i
    Dim j
    Dim k
    
      
    '��ȡ��ҵ��Ϣ
    BookTitle = Mid(strMenulist, posL + 1, posR - posL - 1)
    
    '����ȫ����ǩ
    Do
        i = InStr(1, BookTitle, "<")
        If i Then
            j = InStr(i, BookTitle, ">")
            If j Then
                BookTitle = Replace(BookTitle, Mid(BookTitle, i, j - i + 1), "")
                 
            End If
             
        End If
    Loop While (j <> 0 And i <> 0)
    
    '���������ַ�
    BookTitle = Replace(BookTitle, vbCrLf, "")
    BookTitle = Replace(BookTitle, vbCr, "")
    BookTitle = Replace(BookTitle, vbLf, "")
    BookTitle = Replace(BookTitle, "  ", " ")
    
    
    '��ʾ��Ч����
    Text2 = BookTitle
    
    posL = InStr(1, BookTitle, "��ַ��")
    
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
'    '���½ڱ�����Ϊ�ļ���
'    If FileName = "" Then
'        FileName = BookTitle 'tmpstrTitle
'    End If
'
'
'
'    Exit Sub
'
'    '��ȡ�½���Ϣ
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
'    'Call fileWrite(App.Path & "/debug_cap2.log", "��ȡ<dd><a�����Ϣ:" & vbCrLf & strMenulist)
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
'    'Call fileWrite(App.Path & "/debug_ul_eul.log", "��ȡ<ul></ul>֮�����Ϣ:" & vbCrLf & strMenulist)
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
'    'Call fileWrite(App.Path & "/debug_dd.log", "ȥ��<ul></ul><li></li>��ǵ���Ϣ:" & vbCrLf & strMenulist)
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
'    Text2 = "һ���ҵ� " & TotalCap & " ��" & vbCrLf & Text2
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
'            'Call fileWrite(App.Path & "/debug_Totalcaps.log", "�� " & i & vbTab & " ��:" & UT(i).Title & " - " & UT(i).Url & vbCrLf)
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
''            UT(i).Title = "��" & i + 1 & "�� " & UT(i).Title
''        ElseIf i = 1763 Then
''            UT(i).Title = "���ڸ�����˵�Ļ�"
''        Else
''            UT(i).Title = "��" & i & "�� " & UT(i).Title
'            UT(i).Title = UT(i).Title
''        End If
'
'        Me.Caption = BookTitle & "   " & i & "/" & UBound(k) - 1
'        Call SaveContent(UT(i))
'
'    Next
'    Text2 = "�������!" & vbCrLf & Text2
End Sub

Private Function SaveContent(UT As UrlandTitle)
    Dim strCont     As String
    Dim tmpstrCont  As String
    Dim tmpstrTitle As String
    Dim strHTML     As String

    
    '=====================================��ȡhtml����
    strHTML = GetHtmlStr(UT.Url)
    strHTML = LCase$(strHTML)
    
    'Call fileWrite(App.Path & "/debug_cont.log", UT.Url & vbCrLf & UT.Title & strHTML)
    
    '====================================��ȡ�½ڱ���
        
    tmpstrTitle = GetTitles(strHTML) 'UT.Title
    

    
    '====================================��ȡ�½�����
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


