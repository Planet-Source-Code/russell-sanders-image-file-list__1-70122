VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   345
      Left            =   9090
      TabIndex        =   14
      Top             =   1410
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Maintain Ratio"
      Height          =   255
      Left            =   6840
      TabIndex        =   13
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Auto"
      Height          =   255
      Index           =   4
      Left            =   9360
      TabIndex        =   12
      Top             =   1920
      Width           =   945
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8340
      TabIndex        =   11
      Top             =   1920
      Width           =   765
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Custom:"
      Height          =   255
      Index           =   3
      Left            =   7350
      TabIndex        =   10
      Top             =   1950
      Width           =   945
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Large"
      Height          =   255
      Index           =   2
      Left            =   6450
      TabIndex        =   9
      Top             =   1950
      Width           =   825
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Medium"
      Height          =   255
      Index           =   1
      Left            =   5400
      TabIndex        =   8
      Top             =   1950
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Small"
      Height          =   255
      Index           =   0
      Left            =   4530
      TabIndex        =   7
      Top             =   1950
      Width           =   795
   End
   Begin Thumbs.ucThumbs ctlThumbs 
      Height          =   7965
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   14049
      MaintainRatio   =   -1  'True
      AutoRedraw      =   -1  'True
      Path            =   "C:\Documents and Settings\Mary\Desktop\Russell\ThumbNailViewer\Images"
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      BorderStyle     =   1
      ForeColor       =   0
      LastPath        =   "C:\Documents and Settings\Mary\Desktop\Russell\ThumbNailViewer"
   End
   Begin VB.PictureBox Picture1 
      Height          =   5685
      Left            =   4440
      ScaleHeight     =   5625
      ScaleWidth      =   6255
      TabIndex        =   4
      Top             =   2280
      Width           =   6315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "List count"
      Height          =   435
      Index           =   2
      Left            =   4410
      TabIndex        =   3
      Top             =   1380
      Width           =   2205
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Selected"
      Height          =   465
      Index           =   1
      Left            =   4410
      TabIndex        =   2
      Top             =   900
      Width           =   2205
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Path"
      Height          =   435
      Index           =   0
      Left            =   4410
      TabIndex        =   1
      Top             =   450
      Width           =   2205
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   435
      Left            =   4410
      TabIndex        =   0
      Top             =   0
      Width           =   2205
   End
   Begin VB.Label lbl 
      Caption         =   $"Form1.frx":0000
      Height          =   1125
      Left            =   6750
      TabIndex        =   5
      Top             =   30
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    ctlThumbs.MaintainRatio = (Check1.value = vbChecked)
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Prompt As String
        Select Case Index
            Case 0: Prompt = ctlThumbs.path
            Case 1: Prompt = IIf(ctlThumbs.ListIndex > -1, ctlThumbs.list(ctlThumbs.ListIndex), "Nothing Selected!")
            Case 2: Prompt = ctlThumbs.ListCount
        End Select
    MsgBox Prompt
End Sub

Private Sub Command2_Click()
    ctlThumbs.BrowseFolder
End Sub

Private Sub ctlThumbs_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single, Index As Integer)
On Error Resume Next 'index will return -1 if you click on a blank area
    If Button = 1 And Index > -1 Then Picture1.Picture = LoadPicture(ctlThumbs.path & "\" & ctlThumbs.list(Index))
End Sub

Private Sub ctlThumbs_SelChanged()
    Debug.Print ctlThumbs.FileName
End Sub

Private Sub Form_Load()
    Option1(ctlThumbs.tmbSize).value = True
    Check1.value = (ctlThumbs.MaintainRatio And vbChecked)
End Sub

Private Sub Option1_Click(Index As Integer)
    ctlThumbs.tmbSize = Index
    Text1.Text = ctlThumbs.curDrawSize
End Sub

Private Sub Text1_Change()
    If Val(Text1.Text) > 0 And ctlThumbs.tmbSize = 3 Then ctlThumbs.curDrawSize = Val(Text1.Text)
End Sub

