VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NOI_CPP Helper by TNTCompany"
   ClientHeight    =   11565
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   18960
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11565
   ScaleWidth      =   18960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   3508
      Left            =   10680
      Top             =   9000
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3506
      Left            =   10680
      Top             =   8280
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   15840
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "打开输出目录 (&O)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   8160
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "等线"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Form1.frx":3482
      Top             =   9480
      Width           =   18540
   End
   Begin VB.CommandButton Command4 
      Caption         =   "find return"
      Height          =   375
      Left            =   13440
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "恢复为初始代码 (&F)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "    保存以上代码       到本目录 (&S)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   8160
      Width           =   2655
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   102
      Left            =   14160
      Top             =   8280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1501
      Left            =   14640
      Top             =   8280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "复制以上代码 (&C)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   8160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1200
      Width           =   18495
   End
   Begin VB.Label Label6 
      Caption         =   "打开 cpp 文件：Ctrl + O"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   7200
      TabIndex        =   12
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "编译：F9"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   10920
      TabIndex        =   11
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "编译运行：F11"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   13320
      TabIndex        =   10
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "请在下面的文本框内书写代码，或者直接将代码粘贴在文本框内。"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "提示：恢复为初始代码时，当前代码会被保存到输出目录。"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This software can save your time when you're coding."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   3720
      TabIndex        =   2
      Top             =   0
      Width           =   11295
   End
   Begin VB.Menu Fil 
      Caption         =   "文件 (&F)"
      Begin VB.Menu DK 
         Caption         =   "打开... (&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu Ext 
         Caption         =   "退出 (&X)"
      End
   End
   Begin VB.Menu Edt 
      Caption         =   "初始代码 (&S)"
      Begin VB.Menu He 
         Caption         =   "头文件 (&H)"
         Begin VB.Menu WanCan 
            Caption         =   "使用万能头文件"
            Checked         =   -1  'True
         End
         Begin VB.Menu Si 
            Caption         =   "使用常用单独头文件"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu Rd 
         Caption         =   "快读 (&R)"
         Begin VB.Menu Fr 
            Caption         =   "启用"
            Checked         =   -1  'True
         End
         Begin VB.Menu NFr 
            Caption         =   "禁用"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu Ru 
      Caption         =   "运行 (&R)"
      Begin VB.Menu CR 
         Caption         =   "编译并运行 (&Y)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu CO 
         Caption         =   "编译并打开输出文件夹 (&N)"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu Hp 
      Caption         =   "帮助 (&H)"
      Begin VB.Menu Abt 
         Caption         =   "关于 (&A)"
      End
      Begin VB.Menu WTF 
         Caption         =   "这是什么? (&W)"
      End
      Begin VB.Menu Web 
         Caption         =   "TNT 公司官网 (&T)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Abt_Click()

Abtf.Show (1)

End Sub

Private Sub CO_Click()
On Error Resume Next
Call Command2_Click

Kill (App.Path & "\src\usr.exe")

Dim yin As String
yin = """"

'Shell (App.Path & "\MinGW64\bin\g++.exe " & yin & App.Path & "\src\usr.cpp" & yin & " -o " & yin & App.Path & "\src\usr.exe"), 0
'xshell(g++.exe "D:\Debug\未命名1.cpp" -o "D:\Debug\未命名1.exe"  -I"D:\Program Files (x86)\Dev-Cpp\MinGW64\include" -I"D:\Program Files (x86)\Dev-Cpp\MinGW64\x86_64-w64-mingw32\include" -I"D:\Program Files (x86)\Dev-Cpp\MinGW64\lib\gcc\x86_64-w64-mingw32\4.9.2\include" -I"D:\Program Files (x86)\Dev-Cpp\MinGW64\lib\gcc\x86_64-w64-mingw32\4.9.2\include\c++" -L"D:\Program Files (x86)\Dev-Cpp\MinGW64\lib" -L"D:\Program Files (x86)\Dev-Cpp\MinGW64\x86_64-w64-mingw32\lib" -static-libgcc
Text2.Text = Text2.Text + vbCrLf + "命令: g++.exe " & yin & App.Path & "\src\usr.cpp" & yin & " -o " & yin & App.Path & "\src\usr.exe" & yin & "  -I" & yin & App.Path & "\MinGW64\include" & yin & " -I" & yin & App.Path & "\MinGW64\x86_64-w64-mingw32\include" & yin & " -I" & yin & App.Path & "\MinGW64\lib\gcc\x86_64-w64-mingw32\4.9.2\include" & yin & " -I" & yin & App.Path & "\MinGW64\lib\gcc\x86_64-w64-mingw32\4.9.2\include\c++" & yin & " -L" & yin & App.Path & "\MinGW64\lib" & yin & " -L" & yin & App.Path & "MinGW64\x86_64-w64-mingw32\lib " & yin & "-g -static-libgcc" + vbCrLf

Text2.SelStart = Len(Text2.Text)
Shell (App.Path & "\MinGW64\bin\g++.exe " & yin & App.Path & "\src\usr.cpp" & yin & " -o " & yin & App.Path & "\src\usr.exe" & yin & "  -I" & yin & App.Path & "\MinGW64\include" & yin & " -I" & yin & App.Path & "\MinGW64\x86_64-w64-mingw32\include" & yin & " -I" & yin & App.Path & "\MinGW64\lib\gcc\x86_64-w64-mingw32\4.9.2\include" & yin & " -I" & yin & App.Path & "\MinGW64\lib\gcc\x86_64-w64-mingw32\4.9.2\include\c++" & yin & " -L" & yin & App.Path & "\MinGW64\lib" & yin & " -L" & yin & App.Path & "MinGW64\x86_64-w64-mingw32\lib " & yin & "-g -static-libgcc"), 0
On Error Resume Next
Open App.Path & "\MinGW64\run.bat" For Output As #1
    Print #1, "@echo off"
    Print #1, "set time1=%time:~0,2%%time:~3,2%%time:~6,2%"
    Print #1, """" & App.Path & "\src\usr.exe"""
    Print #1, "@echo\"
    Print #1, "echo ================================"
    Print #1, "set time2=%time:~0,2%%time:~3,2%%time:~6,2%"
    Print #1, "set /a time3=%time2%-%time1%"
    Print #1, "echo Process exited after %time3% second(s) with return value %errorlevel%"
    Print #1, "echo TNT-IDE By: TNTCompany WYL"
    Print #1, "echo Version: V1.0"
    Print #1, "pause"
    Close #1
    

'Call Command5_Click
Timer4.Enabled = True
End Sub

Private Sub Command1_Click()
Clipboard.Clear
'Clipboard.SetText ("Loser")
Clipboard.SetText (Text1.Text)
Command1.Caption = "已复制"
Timer1.Enabled = True
Text1.SetFocus
End Sub

Private Sub Command2_Click()
On Error Resume Next
Open App.Path & "\src\usr.cpp" For Output As #1
Print #1, Text1.Text
Close #1

Text2.Text = Text2.Text + vbCrLf + "已保存为 """ & App.Path & "\src\usr.cpp"""
Text2.SelStart = Len(Text2.Text)
Text1.SetFocus

End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("是否确认将现在的代码替换为初始代码？" + Chr(13) + "您目前的代码会被清除！", 48 + vbDefaultButton2 + 4096 + vbYesNo, "") = vbYes Then
Text1.Text = ""
    If WanCan.Checked And Fr.Checked Then
        Open App.Path & "\def\def.tnt" For Input As #1
            Do While Not EOF(1)
                Line Input #1, textline
                Text1 = Text1 + textline
                Text1 = Text1 + vbCrLf
            Loop
        Close #1
        Text2.Text = Text2.Text + vbCrLf + "已经根据您对初始代码的选择，恢复了初始代码。" + vbCrLf
        Text2.SelStart = Len(Text2.Text)
        Text1.SetFocus
    End If
    
    If (WanCan.Checked And Not (Fr.Checked)) Then
        
    End If
    
    If ((Not WanCan.Checked) And (Not Fr.Checked)) Then
         Open App.Path & "\def\dq.tnt" For Input As #1
            Do While Not EOF(1)
                Line Input #1, textline
                Text1 = Text1 + textline
                Text1 = Text1 + vbCrLf
            Loop
        Close #1
        Text2.Text = Text2.Text + vbCrLf + "已经根据您对初始代码的选择，恢复了初始代码。" + vbCrLf
        Text2.SelStart = Len(Text2.Text)
        Text1.SetFocus
    End If
    
    If (WanCan.Checked And Not (Fr.Checked)) Then
        
    End If
    
    
End If
End Sub

Private Sub Command4_Click() 'mei yong
    Dim lStartTime As Long
  
    '比较两个方式的运行速度
    lStartTime = GetTickCount
    MsgBox FindText(App.Path & "\src\usr.cpp", "return 0;") '此返回值为字符位置
    'MsgBox GetTickCount - lStartTime
End Sub

'使用普通方式查找文件中包含的字符串（返回字符位置）
Private Function FindText(ByVal strFileName As String, ByVal strText As String) As Long
    Dim fn As Integer
    Dim strFileText As String
      
    Dim MyString, MyNumber
    Dim S As String
      
    fn = FreeFile()
    Open strFileName For Binary As #fn   ' 打开输入文件。
    strFileText = Input(LOF(fn), fn)
    Close #fn
    FindText = InStr(strFileText, strText)
End Function



Private Sub Command5_Click()
Shell ("explorer.exe " & App.Path & "\src"), 1


'Shell ("explorer.exe" & App.Path & "\src\"), 1
End Sub

Private Sub CR_Click()
On Error Resume Next
Call Command2_Click



Dim yin As String
yin = """"

'Shell (App.Path & "\MinGW64\bin\g++.exe " & yin & App.Path & "\src\usr.cpp" & yin & " -o " & yin & App.Path & "\src\usr.exe"), 0
'xshell(g++.exe "D:\Debug\未命名1.cpp" -o "D:\Debug\未命名1.exe"  -I"D:\Program Files (x86)\Dev-Cpp\MinGW64\include" -I"D:\Program Files (x86)\Dev-Cpp\MinGW64\x86_64-w64-mingw32\include" -I"D:\Program Files (x86)\Dev-Cpp\MinGW64\lib\gcc\x86_64-w64-mingw32\4.9.2\include" -I"D:\Program Files (x86)\Dev-Cpp\MinGW64\lib\gcc\x86_64-w64-mingw32\4.9.2\include\c++" -L"D:\Program Files (x86)\Dev-Cpp\MinGW64\lib" -L"D:\Program Files (x86)\Dev-Cpp\MinGW64\x86_64-w64-mingw32\lib" -static-libgcc
Text2.Text = Text2.Text + vbCrLf + "命令: g++.exe " & yin & App.Path & "\src\usr.cpp" & yin & " -o " & yin & App.Path & "\src\usr.exe" & yin & "  -I" & yin & App.Path & "\MinGW64\include" & yin & " -I" & yin & App.Path & "\MinGW64\x86_64-w64-mingw32\include" & yin & " -I" & yin & App.Path & "\MinGW64\lib\gcc\x86_64-w64-mingw32\4.9.2\include" & yin & " -I" & yin & App.Path & "\MinGW64\lib\gcc\x86_64-w64-mingw32\4.9.2\include\c++" & yin & " -L" & yin & App.Path & "\MinGW64\lib" & yin & " -L" & yin & App.Path & "MinGW64\x86_64-w64-mingw32\lib " & yin & "-g -static-libgcc" + vbCrLf

Text2.SelStart = Len(Text2.Text)

'Open App.Path & "\MinGW64\inf.bat" For Output As #1
'Print #1, "cd " & App.Path & "\MinGW64\bin"
'Print #1, "g++.exe 2>""" & App.Path & "\MinGW64\inf.txt"""

'Close #1


Shell (App.Path & "\MinGW64\bin\g++.exe " & yin & App.Path & "\src\usr.cpp" & yin & " -o " & yin & App.Path & "\src\usr.exe" & yin & "  -I" & yin & App.Path & "\MinGW64\include" & yin & " -I" & yin & App.Path & "\MinGW64\x86_64-w64-mingw32\include" & yin & " -I" & yin & App.Path & "\MinGW64\lib\gcc\x86_64-w64-mingw32\4.9.2\include" & yin & " -I" & yin & App.Path & "\MinGW64\lib\gcc\x86_64-w64-mingw32\4.9.2\include\c++" & yin & " -L" & yin & App.Path & "\MinGW64\lib" & yin & " -L" & yin & App.Path & "MinGW64\x86_64-w64-mingw32\lib " & yin & "-g -static-libgcc"), 0
On Error Resume Next
Open App.Path & "\MinGW64\run.bat" For Output As #1
    Print #1, "@echo off"
    Print #1, "set time1=%time:~0,2%%time:~3,2%%time:~6,2%"
    Print #1, """" & App.Path & "\src\usr.exe"""
    Print #1, "@echo\"
    Print #1, "echo ================================"
    Print #1, "set time2=%time:~0,2%%time:~3,2%%time:~6,2%"
    Print #1, "set /a time3=%time2%-%time1%"
    Print #1, "echo Process exited after %time3% second(s) with return value %errorlevel%"
    Print #1, "echo TNT-IDE By: TNTCompany WYL"
    Print #1, "echo Version: V1.0"
    Print #1, "pause"
    Close #1
'Shell ("cmd /c attrib +h """ & App.Path & "\src\run.bat"""), 0

Timer2.Enabled = True

End Sub

Private Sub DK_Click()
CommonDialog1.Filter = "C++ Source File (*.cpp)|*.cpp|C (*.c)|*.c"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.DialogTitle = "Select cpp File"
CommonDialog1.ShowOpen

'The FileName property gives you the variable you need to use
On Error Resume Next
If CommonDialog1.FileName <> "" Then
Text1.Text = ""
 Open CommonDialog1.FileName For Input As #1
        Do While Not EOF(1)
            Line Input #1, textline
            Text1 = Text1 + textline
            Text1 = Text1 + vbCrLf
        Loop
    Close #1

Text2.Text = Text2.Text + vbCrLf + "打开：" & CommonDialog1.FileName + vbCrLf
Text2.SelStart = Len(Text2.Text)
End If
End Sub

Private Sub Ext_Click()
End
End Sub

Private Sub Form_Load()
Text1.Text = ""
On Error Resume Next
 Open App.Path & "\def\def.tnt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, textline
            Text1 = Text1 + textline
            Text1 = Text1 + vbCrLf
        Loop
    Close #1
    
    Si.Checked = False
    WanCan.Checked = True
    Fr.Checked = True
    NFr.Checked = False
    
    
    'Timer1.Enabled = True
End Sub


Private Sub Fr_Click()
If Not (Fr.Checked) Then
Fr.Checked = True
NFr.Checked = False
Call Cg
End If
End Sub

Private Sub NFr_Click()
If Not (NFr.Checked) Then
NFr.Checked = True
Fr.Checked = False
Call Cg
End If
End Sub

Private Sub Si_Click()
If Not (Si.Checked) Then
Si.Checked = True
WanCan.Checked = False
Call Cg
End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 1 Then
  
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
    End If

End Sub


Private Sub Timer1_Timer()
'Text1.SetFocus
Command1.Caption = "复制 (&C)"
'Text1.SetFocus
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Kill (App.Path & "\src\usr.exe")


'If Not fs.fileexists(App.Path & "\src\usr.exe") Then


Timer3.Enabled = True

'Shell (App.Path & "\MinGW64\run.bat"), 1
Timer2.Enabled = False
'End If
End Sub


Private Sub Cg()
On Error Resume Next
Open App.Path & "\src\usr.cpp" For Output As #1
Print #1, Text1.Text
Close #1

Text2.Text = Text2.Text + vbCrLf + "刚才的代码已保存为 """ & App.Path & "\src\usr.cpp"""
Text2.SelStart = Len(Text2.Text)
Text1.SetFocus


If WanCan.Checked And Fr.Checked Then
    Text1.Text = ""
    Open App.Path & "\def\def.tnt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, textline
            Text1 = Text1 + textline
            Text1 = Text1 + vbCrLf
        Loop
    Close #1
End If

If WanCan.Checked And Not (Fr.Checked) Then
    Text1.Text = ""
    Open App.Path & "\def\wn.tnt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, textline
            Text1 = Text1 + textline
            Text1 = Text1 + vbCrLf
        Loop
    Close #1
End If

If Not (WanCan.Checked) And Fr.Checked Then
    Text1.Text = ""
    Open App.Path & "\def\fu.tnt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, textline
            Text1 = Text1 + textline
            Text1 = Text1 + vbCrLf
        Loop
    Close #1
End If

If Not (WanCan.Checked) And Not (Fr.Checked) Then
    Text1.Text = ""
    Open App.Path & "\def\dq.tnt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, textline
            Text1 = Text1 + textline
            Text1 = Text1 + vbCrLf
        Loop
    Close #1
End If
End Sub




Private Sub Timer3_Timer()
On Error Resume Next
    Dim fs As New FileSystemObject
    
    If fs.FileExists(App.Path & "\src\usr.exe") Then
    Text2.Text = Text2.Text + vbCrLf + "编译成功！" + vbCrLf
    Text2.SelStart = Len(Text2.Text)
        Shell (App.Path & "\MinGW64\run.bat"), 1
        'cnt = 0
        Timer3.Enabled = False
   Else
        Text2.Text = Text2.Text + vbCrLf + "编译失败！请检查是否有语法错误！" + vbCrLf
    Text2.SelStart = Len(Text2.Text)
    Timer3.Enabled = False
    End If
End Sub

Private Sub Timer4_Timer()
Dim fs As New FileSystemObject
    
    If fs.FileExists(App.Path & "\src\usr.exe") Then
    Text2.Text = Text2.Text + vbCrLf + "编译成功！" + vbCrLf
    Text2.SelStart = Len(Text2.Text)
        'Shell (App.Path & "\MinGW64\run.bat"), 1
        Call Command5_Click
        'cnt = 0
        Timer4.Enabled = False
   Else
        Text2.Text = Text2.Text + vbCrLf + "编译失败！请检查是否有语法错误！" + vbCrLf
    Text2.SelStart = Len(Text2.Text)
    Timer4.Enabled = False
    End If
End Sub

Private Sub WanCan_Click()
If Not (WanCan.Checked) Then
WanCan.Checked = True
Si.Checked = False
    Call Cg
End If
End Sub

Private Sub Web_Click()
Shell "explorer.exe http://www.tntco.icoc.me", 1
End Sub

Private Sub WTF_Click()
What.Show (1)
End Sub
