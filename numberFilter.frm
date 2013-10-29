VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form numberFilter 
   Caption         =   "号码筛选软件"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10155
   Icon            =   "numberFilter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   10155
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton browse 
      Caption         =   "浏览"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "规则"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "筛选"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox ruleCombo 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   1  'ON
      ItemData        =   "numberFilter.frx":76115
      Left            =   120
      List            =   "numberFilter.frx":76117
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "numberFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub browse_Click()
  Dim str
    str = GetFolder(Me.hWnd, "浏览文件夹")
End Sub

Private Sub Command2_Click()
  ruleWindow.Show (1)
End Sub


Private Sub Form_Load()
  If Dir(Environ("APPDATA") & "\numberFilter.txt") <> "" Then
    Open Environ("APPDATA") & "\numberFilter.txt" For Input As #1
    i = 0
    Do While Not EOF(1)
      Line Input #1, a '读整行的数据
      If Trim(a) <> "" Then
       ruleCombo.AddItem a
       i = i + 1
       End If
    Loop
    Close #1
  End If
End Sub
