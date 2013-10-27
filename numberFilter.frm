VERSION 5.00
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
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "numberFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
  ruleWindow.Show (1)
End Sub
