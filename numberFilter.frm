VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form numberFilter 
   Caption         =   "号码筛选软件"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10005
   Icon            =   "numberFilter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   10005
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton exportNotMatch 
      Caption         =   "导  出"
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton exportMatch 
      Caption         =   "导  出"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   3480
      Width           =   3015
   End
   Begin VB.ListBox notMatchedList 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   6840
      TabIndex        =   5
      Top             =   720
      Width           =   3015
   End
   Begin VB.ListBox matchedList 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   3015
   End
   Begin VB.ListBox sourceList 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog fileBrowse 
      Left            =   2040
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton browse 
      Caption         =   "浏览"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "规则"
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "筛选"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label fileName 
      Caption         =   "请选择号码文件进行分析"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "numberFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fileSelected As Boolean
Dim numberList() As String
Dim ruleReg As RegExp
Dim colMatches   As MatchCollection
Dim m As Match

Private Sub browse_Click()
 fileBrowse.ShowOpen
 fileName.Caption = fileBrowse.fileName
 fileSelected = True
 '开始处理数据
 Open fileBrowse.fileName For Input As #1
 i = 0
 ReDim numberList(1000)
 sourceList.Clear
 Do While Not EOF(1)
   Line Input #1, a '读整行的数据
   If Trim(a) <> "" Then
    sourceList.AddItem a
    If i >= UBound(numberList) Then
      ReDim Preserve numberList(UBound(numberList) + 1000)
    End If
    numberList(i) = a
    i = i + 1
   End If
 Loop
 Close #1
End Sub

Private Sub Command1_Click()
  
  If fileSelected = False Then
     MsgBox "请选择文件"
  Else
  
   Dim rulesArray(100) As String
   Dim ruleNumber As String
   Dim ruleName As String
         '读取规则
          
          If Dir(Environ("APPDATA") & "\numberFilter.txt") <> "" Then
          Open Environ("APPDATA") & "\numberFilter.txt" For Input As #1
          i = 0
          Do While Not EOF(1)
            Line Input #1, a '读整行的数据
            If Trim(a) <> "" Then
             rulesArray(i) = a
             i = i + 1
             End If
          Loop
          Close #1
          
       '解析规则
       For rulePosition = 0 To 100
        If rulesArray(rulePosition) <> "" Then
        
            Set colMatches = ruleReg.Execute(rulesArray(rulePosition))
            For Each m In colMatches
             ruleNumber = m.SubMatches(0)
             ruleName = m.SubMatches(1)
            Next
            
            For position = 0 To UBound(numberList)
                 If Left(numberList(position), 7) = ruleNumber Then
                  matchedList.AddItem (numberList(position) & "(" & ruleName & ")")
                   numberList(position) = ""
               
                 End If
            Next
        End If
       Next
       
       For position = 0 To UBound(numberList)
         If numberList(position) <> "" Then
           notMatchedList.AddItem (numberList(position))
         End If
       Next
       
       End If
  End If
End Sub

Private Sub Command2_Click()
  ruleWindow.Show (1)
End Sub


Private Sub exportMatch_Click()
 fileBrowse.Filter = "号码导出文件"
 fileBrowse.ShowSave
 
  Open addSufix(fileBrowse.fileName) For Output As #1  ' 打开输出文件。
   For i = 0 To matchedList.ListCount
        If matchedList.List(i) <> "" Then
             Print #1, matchedList.List(i)
       End If
    Next i
  Close #1
  
End Sub

Private Sub exportNotMatch_Click()
   fileBrowse.Filter = "号码导出文件"
 fileBrowse.ShowSave
 
  Open addSufix(fileBrowse.fileName) For Output As #1  ' 打开输出文件。
   For i = 0 To notMatchedList.ListCount
        If notMatchedList.List(i) <> "" Then
             Print #1, notMatchedList.List(i)
       End If
    Next i
  Close #1
End Sub

Private Sub Form_Load()
fileSelected = False
Set ruleReg = New RegExp
ruleReg.Pattern = "^(\d{7})\((.*)\)$"
End Sub

Function addSufix(fileName) As String
  If Right(fileName, 4) <> ".txt" Then
    fileName = fileName & ".txt"
  End If
  addSufix = fileName
End Function
