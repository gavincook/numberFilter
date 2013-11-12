VERSION 5.00
Begin VB.Form ruleWindow 
   Caption         =   "规则管理"
   ClientHeight    =   3885
   ClientLeft      =   8430
   ClientTop       =   5070
   ClientWidth     =   6900
   Icon            =   "ruleWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   6900
   Begin VB.CommandButton Command3 
      Caption         =   "删除规则"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "更新规则"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox location 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加规则"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox ruleText 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox ruleList 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      ItemData        =   "ruleWindow.frx":76115
      Left            =   120
      List            =   "ruleWindow.frx":76117
      TabIndex        =   0
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Label Label2 
      Caption         =   "归属地："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "号码："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "ruleWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rulesArray(100) As String
 Dim i As Integer
 Dim ruleReg As RegExp

Private Sub Command1_Click()

  If Not ruleReg.Test(ruleText.Text) Then
    MsgBox "请输入手机号的前7位"
    Exit Sub
  End If
  
  For i = 0 To UBound(rulesArray)
       If ruleText.Text = Mid(rulesArray(i), 1, 7) Then
         MsgBox "该号码已经存在"
          Exit Sub
         End If
  Next i
  
  Open Environ("APPDATA") & "\numberFilter.txt" For Append As #1  ' 打开输出文件。
  Print #1, ruleText.Text & "(" & location.Text & ")" ' 将文本数据写入文件。
  Close #1
 
  ruleList.AddItem ruleText.Text & "(" & location.Text & ")"
  ruleText.Text = ""
End Sub

Private Sub Command2_Click()


    If Not ruleReg.Test(ruleText.Text) Then
      MsgBox "请输入手机号的前7位"
      Exit Sub
    End If
  
    If ruleList.ListIndex > -1 Then
       ruleList.List(ruleList.ListIndex) = ruleText.Text & "(" & location.Text & ")"
       Else
       MsgBox "请选择要更新的规则"
    End If
    
    Open Environ("APPDATA") & "\numberFilter.txt" For Output As #1  ' 打开输出文件。
    For i = 0 To ruleList.ListCount
        If ruleList.List(i) <> "" Then
             Print #1, ruleList.List(i)
       End If
    Next i
    Close #1
End Sub

Private Sub Command3_Click()
    If ruleList.ListIndex > -1 Then
       ruleList.RemoveItem (ruleList.ListIndex)
       Else
       MsgBox "请选择要删除的规则"
    End If
    
    Open Environ("APPDATA") & "\numberFilter.txt" For Output As #1  ' 打开输出文件。
    For i = 0 To ruleList.ListCount
        If ruleList.List(i) <> "" Then
             Print #1, ruleList.List(i)
       End If
    Next i
    Close #1
End Sub

Private Sub Form_Load()
  Set ruleReg = New RegExp
  ruleReg.Pattern = "^\d{7}$"

  If Dir(Environ("APPDATA") & "\numberFilter.txt") <> "" Then
  Open Environ("APPDATA") & "\numberFilter.txt" For Input As #1
  i = 0
  Do While Not EOF(1)
    Line Input #1, a '读整行的数据
    If Trim(a) <> "" Then
     ruleList.AddItem a
     rulesArray(i) = a
     i = i + 1
     End If
  Loop
  Close #1
End If

End Sub

Private Sub ruleList_Click()
  Dim reg As RegExp
  Dim colMatches   As MatchCollection
  Dim m As Match
 Set reg = New RegExp
reg.Pattern = "^(\d{7})\((.*)\)$"
Set colMatches = reg.Execute(ruleList.Text)

For Each m In colMatches
 ruleText.Text = m.SubMatches(0)
 location.Text = m.SubMatches(1)
Next
' MsgBox ruleList.Text
  'ruleList.Selected) = "12312313213"
    
End Sub



Private Sub ruleText_Change()
 'For i = 0 To UBound(rulesArray)
  '     If ruleText.Text.StartsWith("132") Then
 '        ruleList.Selected(i) = True
 '        MsgBox i
 '        Exit For
 '        End If
 'Next i
  
End Sub
