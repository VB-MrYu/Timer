VERSION 5.00
Begin VB.Form Frm 
   Caption         =   "��ʱ��"
   ClientHeight    =   3135
   ClientLeft      =   8460
   ClientTop       =   5220
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   Begin VB.CommandButton Command3 
      Caption         =   "ȡ��"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ͣ"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   2640
   End
   Begin VB.Label Label 
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Timer1.Enabled = True Then
Call tcqr
Else
End
End If
Cancel = True
End Sub

Private Sub tcqr()
If MsgBox("��ʱ�������ڼ����������Ҫ�˳���", vbYesNo) = vbYes Then End
End Sub

Private Sub Command1_Click()
Timer1.Enabled = True
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
Label = 0
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Timer1_Timer()
Label = Label + 1
End Sub

