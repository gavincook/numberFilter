VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ruleWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������"
   ClientHeight    =   3750
   ClientLeft      =   8355
   ClientTop       =   5295
   ClientWidth     =   8100
   Icon            =   "ruleWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   8100
   Begin VB.CommandButton clearRule 
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton exportRule 
      Caption         =   "���򵼳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton importRule 
      Caption         =   "������"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog ruleFileDialog 
      Left            =   840
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ɾ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���¹���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox location 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ӹ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   130
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
         Name            =   "����"
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
      Width           =   2535
   End
   Begin VB.ListBox ruleList 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      ItemData        =   "ruleWindow.frx":76115
      Left            =   120
      List            =   "ruleWindow.frx":76117
      TabIndex        =   0
      Top             =   1080
      Width           =   7935
   End
   Begin VB.Label Label2 
      Caption         =   "�����أ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "���룺"
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Menu file 
      Caption         =   "�ļ�"
      Begin VB.Menu import 
         Caption         =   "������"
         Shortcut        =   {F2}
      End
      Begin VB.Menu export 
         Caption         =   "���򵼳�"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "ruleWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rulesArray() As String
 Dim i As Integer
 Dim ruleReg As RegExp


Private Sub clearRule_Click()
Dim answer As String
  answer = MsgBox("��պ󲻿ɻָ�! ȷ�����?", vbYesNo, "ȷ��")
  If answer = vbYes Then
    ruleList.Clear
    Open Environ("APPDATA") & "\numberFilter.txt" For Output As #1  ' ������ļ���
      Print #1, "" ' ���ı�����д���ļ���
    Close #1
    ReDim rulesArray(100)
  End If
  
End Sub

Private Sub Command1_Click()

  If Not ruleReg.Test(ruleText.Text) Then
    MsgBox "�������ֻ��ŵ�ǰ7λ"
    Exit Sub
  End If
  
  For i = 0 To UBound(rulesArray)
       If ruleText.Text = Mid(rulesArray(i), 1, 7) Then
         MsgBox "�ú����Ѿ�����"
          Exit Sub
         End If
  Next i
  
  Open Environ("APPDATA") & "\numberFilter.txt" For Append As #1  ' ������ļ���
  Print #1, ruleText.Text & "  " & location.Text  ' ���ı�����д���ļ���
  Close #1
 
  ruleList.AddItem ruleText.Text & "  " & location.Text
  ruleText.Text = ""
End Sub

Private Sub Command2_Click()


    If Not ruleReg.Test(ruleText.Text) Then
      MsgBox "�������ֻ��ŵ�ǰ7λ"
      Exit Sub
    End If
  
    If ruleList.ListIndex > -1 Then
       ruleList.List(ruleList.ListIndex) = ruleText.Text & "  " & location.Text
       Else
       MsgBox "��ѡ��Ҫ���µĹ���"
    End If
    
    Open Environ("APPDATA") & "\numberFilter.txt" For Output As #1  ' ������ļ���
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
       MsgBox "��ѡ��Ҫɾ���Ĺ���"
    End If
    
    Open Environ("APPDATA") & "\numberFilter.txt" For Output As #1  ' ������ļ���
    For i = 0 To ruleList.ListCount
        If ruleList.List(i) <> "" Then
             Print #1, ruleList.List(i)
       End If
    Next i
    Close #1
End Sub

Private Sub export_Click()
   ruleFileDialog.Filter = "*.txt"
   ruleFileDialog.ShowSave
   If ruleFileDialog.fileName <> "" Then
       Open numberFilter.addSufix(ruleFileDialog.fileName) For Output As #1
       Open Environ("APPDATA") & "\numberFilter.txt" For Input As #2   ' �������ļ���
       Do While Not EOF(2)
         Line Input #2, a '�����е�����
         If Trim(a) <> "" Then
          Print #1, a
         End If
       Loop
       Close #2
       Close #1
    MsgBox "�����ɹ����ļ�·��Ϊ " & numberFilter.addSufix(ruleFileDialog.fileName)
    End If
End Sub


Private Sub exportRule_Click()
 ruleFileDialog.Filter = "���뵼���ļ�|*.txt"
   ruleFileDialog.ShowSave
    If ruleFileDialog.fileName <> "" Then

       Open numberFilter.addSufix(ruleFileDialog.fileName) For Output As #1
       Open Environ("APPDATA") & "\numberFilter.txt" For Input As #2   ' �������ļ���
       Do While Not EOF(2)
         Line Input #2, a '�����е�����
         If Trim(a) <> "" Then
          Print #1, a
         End If
       Loop
       
       Close #2
       Close #1
    MsgBox "�����ɹ����ļ�·��Ϊ " & numberFilter.addSufix(ruleFileDialog.fileName)
    End If
End Sub

Private Sub Form_Load()
ReDim rulesArray(100)
 Dim fontSize As Double
  fontSize = 12
  clearRule.fontSize = fontSize
  importRule.fontSize = fontSize
  exportRule.fontSize = fontSize
  Command1.fontSize = fontSize
  Command2.fontSize = fontSize
  Command3.fontSize = fontSize
  

  
  Set ruleReg = New RegExp
  ruleReg.Pattern = "^\d{7}$"

  If Dir(Environ("APPDATA") & "\numberFilter.txt") <> "" Then
    Open Environ("APPDATA") & "\numberFilter.txt" For Input As #1
    i = 0
    Do While Not EOF(1)
      Line Input #1, a '�����е�����
      If Trim(a) <> "" Then
       ruleList.AddItem a
       rulesArray(i) = a
       i = i + 1
       End If
    Loop
    Close #1
   End If

End Sub


Private Sub import_Click()
    ruleFileDialog.ShowOpen
    If ruleFileDialog.fileName <> "" Then
       '��ʼ��������
       Open ruleFileDialog.fileName For Input As #1
       Open Environ("APPDATA") & "\numberFilter.txt" For Output As #2  ' ������ļ���
       Do While Not EOF(1)
         Line Input #1, a '�����е�����
         If Trim(a) <> "" Then
          Print #2, a
         End If
       Loop
       
       Close #2
       Close #1
       ruleList.Clear
         If Dir(Environ("APPDATA") & "\numberFilter.txt") <> "" Then
            Open Environ("APPDATA") & "\numberFilter.txt" For Input As #3
            i = 0
            Do While Not EOF(3)
              Line Input #3, a '�����е�����
              If Trim(a) <> "" Then
               ruleList.AddItem a
               rulesArray(i) = a
               i = i + 1
               End If
            Loop
            Close #3
        End If
        MsgBox "������ɹ�"
    End If
End Sub

Private Sub importRule_Click()
   ruleFileDialog.Filter = "���뵼���ļ�|*.txt"
   ruleFileDialog.ShowOpen
    If ruleFileDialog.fileName <> "" Then
       '��ʼ��������
       Open ruleFileDialog.fileName For Input As #1
       
       Open Environ("APPDATA") & "\numberFilter.txt" For Append As #2  ' ������ļ���
       Do While Not EOF(1)
       
         Line Input #1, a '�����е�����
         If contains(rulesArray, a) Then
           MsgBox ("����" & a & " �Ѿ�����")
         Else
            If Trim(a) <> "" Then
             Print #2, a
            End If
         End If
       Loop
       
       Close #2
       Close #1
       ruleList.Clear
         If Dir(Environ("APPDATA") & "\numberFilter.txt") <> "" Then
            Open Environ("APPDATA") & "\numberFilter.txt" For Input As #3
            i = 0
            Do While Not EOF(3)
              Line Input #3, a '�����е�����
              If Trim(a) <> "" Then
               ruleList.AddItem a
               rulesArray(i) = a
               i = i + 1
               End If
            Loop
            Close #3
        End If
        MsgBox "������ɹ�"
    End If
End Sub

Function contains(ruleArray, element) As Boolean
contains = False
   For Each rule In ruleArray
     If rule = Trim(element) Then
      contains = True
      Exit For
     End If
   Next
   
End Function


Private Sub ruleList_Click()
  Dim reg As RegExp
  Dim colMatches   As MatchCollection
  Dim m As Match
 Set reg = New RegExp
reg.Pattern = "^(\d{7})\s*(.*)$"
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
