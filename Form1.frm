VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����С˵������"
   ClientHeight    =   3600
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10995
   FillColor       =   &H80000005&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   10995
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtText3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8100
      TabIndex        =   9
      Text            =   "1"
      Top             =   90
      Width           =   780
   End
   Begin VB.CommandButton cmdgetmenu 
      Caption         =   "����"
      Height          =   300
      Left            =   9945
      TabIndex        =   7
      Top             =   90
      Width           =   990
   End
   Begin VB.TextBox txtHttpWww 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   1350
      TabIndex        =   0
      Text            =   "http://www.ykanshu.com/files/article/html/129/129598/"
      Top             =   90
      Width           =   6225
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
      Height          =   3090
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   450
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
   Begin VB.CheckBox chk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "�Զ�������һ��"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8910
      TabIndex        =   3
      Top             =   630
      Value           =   1  'Checked
      Width           =   1950
   End
   Begin VB.CommandButton cmdCommand2 
      Caption         =   "����"
      Height          =   315
      Left            =   7785
      TabIndex        =   6
      Top             =   630
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ӵ�          �¿�ʼ"
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
      Caption         =   "���Ŀ¼��ַ��"
      Height          =   180
      Index           =   0
      Left            =   45
      TabIndex        =   5
      Top             =   135
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ�Ķ���ַ��"
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
    strHTML = GetHtmlStr(Text1)
    strHTML = LCase$(strHTML)
    '====================================��ȡ�½ڱ���
    tmpstrTitle = GetTitle(strHTML)
    
    If Len(tmpstrTitle) = 0 Then
        Exit Sub
    End If
    
    tmpstrTitle = Trim$(tmpstrTitle)
    '���½ڱ�����Ϊ�ļ���
    If filename = "" Then
        filename = tmpstrTitle
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

    Call fileWrite(App.Path & "\" & filename & ".txt", strCont)
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
        DelayMs (100)
    Loop

    GetHtmlStr = StrConv(xml.ResponseBody, vbUnicode)
    'GetHtmlStr = xml.responsetext
    Set xml = Nothing
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
    strContent = Replace$(strContent, " ", "")
    strContent = Replace$(strContent, "��", "")
    strContent = Replace(strContent, "&nbsp;", "")
    strContent = Replace(strContent, "?", "")
    strContent = Replace$(strContent, "<br/>", vbCrLf)
    strContent = Replace$(strContent, vbCrLf & vbCrLf, vbCrLf)
    

    
    
'    Dim a
'    a = InStr(1, strContent, "<p>��������Ʒ����")
'
'    If a Then
'        strContent = Left$(strContent, a - 1)
'    End If
'
'    Dim b
'    a = InStr(1, strContent, "��������")
'
'    If a Then
'        b = InStr(a, strContent, "��������С˵")
'
'        If b Then '
'            strCut = Mid$(strContent, a, b + 7 - a)
'            strContent = Replace$(strContent, strCut, "")
'        End If
'    End If
    
'    strContent = Replace$(strContent, "<p>", "")
'    'strContent = Replace$(strContent, "</p>", "")
'    strContent = Replace$(strContent, "</p>", vbCrLf)
     
    
    strContent = Replace$(strContent, "Ҽ���������������������������", "")
    strContent = Replace$(strContent, "Ҽ���������������������������", "")
    strContent = Replace$(strContent, "Ҫ��������������������������", "")
    strContent = Replace$(strContent, "һ�������������������������", "")
    strContent = Replace$(strContent, "һ����������������Ҫ������������", "")
    strContent = Replace$(strContent, "Ҫ��������������������������� ", "")
    strContent = Replace$(strContent, "Ҫ�������������������������", "")
    strContent = Replace$(strContent, "Ҽ�������������������������", "")
    strContent = Replace$(strContent, "Ҫ��������������������������", "")
    strContent = Replace$(strContent, "Ҽ���������������ῴ������������", "")
    strContent = Replace$(strContent, "һ���������Ҫ����Ҫ������������", "")
    strContent = Replace$(strContent, "Ҽ���������������Ҫ���󿴣�������", "")
    strContent = Replace$(strContent, "һ���������������������������", "")
    strContent = Replace$(strContent, "Ҽ����������顤��������������", "")
    strContent = Replace$(strContent, "Ҫ�������Ҫ��������������������", "")
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

Private Sub cmdgetmenu_Click()

    
    
    Dim strMenulist
    Dim posL, posR
        
    '==================================== ���鼮��ҳ��ַ���и�ʽ��
    If Right$(txtHttpWww, 1) <> "/" Then
        txtHttpWww = txtHttpWww & "/"
    End If
    
    
    
    strMenulist = GetHtmlStr(txtHttpWww)
    strMenulist = LCase$(strMenulist)
    'Call fileWrite(App.Path & "/debug.log", "strMenulist" & vbCrLf & vbCrLf & "-------------------------------------" & vbCrLf & strMenulist)
    strMenulist = Replace$(strMenulist, "  ", "")
    strMenulist = Replace$(strMenulist, vbTab, "")
    strMenulist = Replace$(strMenulist, "<br/>", vbCrLf)
    
    strMenulist = Replace$(strMenulist, vbCrLf & vbCrLf, vbCrLf)
    strMenulist = Replace$(strMenulist, vbCrLf & " ", vbCrLf)
    
    'Call fileWrite(App.Path & "/debug_strmenu.log", "strMenulist" & vbCrLf & vbCrLf & "-------------------------------------" & vbCrLf & strMenulist)
'    Call fileWrite(App.Path & "/debug_replaceSPACE.log", "ȥ��""  ""�����ո�֮���ԭʼ��Ϣ:" & vbCrLf & strMenulist)
    
    posL = InStr(1, strMenulist, "<div id=""info"">")
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
    
    Text2 = "����:" & BookTitle
    
    
'    Call fileWrite(App.Path & "/debug_Booktitle.log", "Book Title:" & BookTitle)
    '���½ڱ�����Ϊ�ļ���
    If filename = "" Then
        filename = BookTitle 'tmpstrTitle
    End If
    
    
    '��ȡ�½���Ϣ
'    posL = InStr(1, strMenulist, "<dd>")
'    If posL = 0 Then
'        Exit Sub
'    End If
'    strMenulist = Mid$(strMenulist, posL)
    
    
    'Call fileWrite(App.Path & "/debug_cap.log", "��ȡdd�����Ϣ:" & vbCrLf & strMenulist)
    
    posL = InStr(1, strMenulist, "<dd><a")
    If posL = 0 Then
        Exit Sub
    End If
    
'    posL = InStr(posL, strMenulist, "<a")
'    If posL = 0 Then
'        Exit Sub
'    End If
    
    
    
    strMenulist = Mid$(strMenulist, posL)
'    Call fileWrite(App.Path & "/debug_cap2.log", "��ȡ<ul>�����Ϣ:" & vbCrLf & strMenulist)
    
    
    posR = InStr(1, strMenulist, "</dl>")
    If posR = 0 Then
        Exit Sub
    End If
    
    
    
    strMenulist = Left(strMenulist, posR - 1)
'    Call fileWrite(App.Path & "/debug_ul_eul.log", "��ȡ<ul></ul>֮�����Ϣ:" & vbCrLf & strMenulist)
    
'    strMenulist = Replace$(strMenulist, "<ul>", "")
'    strMenulist = Replace$(strMenulist, "</ul>", "")
    strMenulist = Replace$(strMenulist, "<dd>", "")
    strMenulist = Replace$(strMenulist, "</dd>", "")
'    strMenulist = Replace$(strMenulist, "<span></span>", "")
'    strMenulist = Replace$(strMenulist, "-" & BookTitle, "")
'    strMenulist = Replace$(strMenulist, vbCr, "")
'    strMenulist = Replace$(strMenulist, vbLf, "")
    
'    Call fileWrite(App.Path & "/debug_dd.log", "ȥ��<ul></ul><li></li>��ǵ���Ϣ:" & vbCrLf & strMenulist)
    
    Dim k
    Dim i, j
    Dim hostUrl
    
    Dim s As String
    k = Split(strMenulist, "</a>")
    i = UBound(k) - 1
    ReDim UT(i) As UrlandTitle
    
    Text2 = "һ���ҵ� " & i & " ��" & vbCrLf & Text2
    
    j = Val(txtText3) - 1
    If j < 0 Then j = 0
     
    txtText3 = j + 1
    hostUrl = Left(txtHttpWww, InStr(10, txtHttpWww, "/") - 1)
    
    For i = j To UBound(k) - 1
        k(i) = Trim$(k(i))
        posL = InStr(1, k(i), """") + 1
        If posL <> 0 Then
            posR = InStr(posL, k(i), """") - 1
             If posR <> 0 Then
                UT(i).Url = hostUrl & Mid$(k(i), posL, posR - posL + 1)
            End If
        End If
        If Len(UT(i).Url) Then
            posL = InStrRev(k(i), ">") + 1
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
'            UT(i).Title = "��" & i + 1 & "�� " & UT(i).Title
'        ElseIf i = 1763 Then
'            UT(i).Title = "���ڸ�����˵�Ļ�"
'        Else
'            UT(i).Title = "��" & i & "�� " & UT(i).Title
            UT(i).Title = UT(i).Title
'        End If
        
        Me.Caption = BookTitle & "   " & i & "/" & UBound(k) - 1
        Call SaveContent(UT(i))
        
    Next
    Text2 = "�������!" & vbCrLf & Text2
End Sub

Private Function SaveContent(UT As UrlandTitle)
    Dim strCont     As String
    Dim tmpstrCont  As String
    Dim tmpstrTitle As String
    Dim strHTML     As String

    
    '=====================================��ȡhtml����
    strHTML = GetHtmlStr(UT.Url)
    strHTML = LCase$(strHTML)
    
'    Call fileWrite(App.Path & "/debug_cont.log", UT.Url & vbCrLf & UT.Title & strHTML)
    
    '====================================��ȡ�½ڱ���
        
    tmpstrTitle = UT.Title
    

    
    '====================================��ȡ�½�����
    tmpstrCont = GetContent(strHTML)

    
    If Len(tmpstrCont) = 0 Then
        Exit Function
    End If
    'Call fileWrite(App.Path & "/debug_cont.log", UT.Url & vbCrLf & UT.Title & tmpstrCont)
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

