VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Art effect"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog com 
      Left            =   5880
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Save picture"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Load picture"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clear effect"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton cmdfilleffect 
      Caption         =   "&Fill effect"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6960
      Top             =   9000
   End
   Begin VB.HScrollBar scrred 
      Height          =   255
      Left            =   840
      Max             =   255
      Min             =   1
      TabIndex        =   3
      Top             =   8880
      Value           =   1
      Width           =   1335
   End
   Begin VB.HScrollBar scrgreen 
      Height          =   255
      Left            =   840
      Max             =   255
      Min             =   1
      TabIndex        =   2
      Top             =   9240
      Value           =   1
      Width           =   1335
   End
   Begin VB.HScrollBar scrblue 
      Height          =   255
      Left            =   840
      Max             =   255
      Min             =   1
      TabIndex        =   1
      Top             =   9600
      Value           =   1
      Width           =   1335
   End
   Begin VB.PictureBox p1 
      AutoRedraw      =   -1  'True
      Height          =   8535
      Left            =   120
      ScaleHeight     =   8475
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   120
      Width           =   11295
   End
   Begin VB.Label Label3 
      Caption         =   "Green"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Blue"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   8880
      Width           =   735
   End
   Begin VB.Label lblcolour 
      Height          =   975
      Left            =   2280
      TabIndex        =   4
      Top             =   8880
      Width           =   615
   End
End
Attribute VB_Name = "frmart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mousepress As Boolean

Private Sub cmdfilleffect_Click()
p1.Line (0, 0)-(9400, 6700), lblcolour.BackColor, BF
End Sub

Private Sub Command1_Click()
p1.Cls
End Sub

Private Sub Command2_Click()
com.ShowOpen
p1.Picture = LoadPicture(com.FileName)
End Sub

Private Sub Command3_Click()
com.ShowSave
SavePicture p1.Image, com.FileName
End Sub

Private Sub Form_Load()
p1.BackColor = vbWhite
p1.DrawWidth = 60
p1.DrawMode = 15
End Sub

Private Sub scrblue_Change()
p1.Cls
p1.Line (0, 0)-(9400, 6700), lblcolour.BackColor, BF
End Sub

Private Sub scrgreen_Change()
p1.Cls
p1.Line (0, 0)-(9400, 6700), lblcolour.BackColor, BF
End Sub

Private Sub scrred_Change()
p1.Cls
p1.Line (0, 0)-(9400, 6700), lblcolour.BackColor, BF
End Sub

Private Sub Timer1_Timer()
lblcolour.BackColor = RGB(scrred.Value, scrgreen.Value, scrblue.Value)

End Sub
Private Sub p1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
mousepress = True
p1.PSet (X, Y), lblcolour.BackColor
End If
End Sub

Private Sub p1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mousepress = True Then
p1.Line -(X, Y), lblcolour.BackColor
End If
End Sub

Private Sub p1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
mousepress = False
End If
End Sub
