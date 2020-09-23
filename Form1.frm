VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "No OCX statusbar example made by JOHANNES BOHMAN!"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   371
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Text            =   "%"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Text            =   "0"
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change color"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   4815
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   20
      Left            =   1080
      Max             =   100
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   480
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   0
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "STATUS BAR WITHOUT OCX!!!! Try changing the max value on the scroll bar! Or try resizing the statusbar! IT WORKS ANYWAY!!!"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   360
      Top             =   2520
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim JB As Integer
Dim temp, temp2, temp3 As Integer

Private Sub Command1_Click()
Picture1.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

Private Sub Form_Load()
JB = Picture1.Width
Picture1.Width = 0
temp3 = HScroll1.Value / 100
End Sub

Private Sub HScroll1_Change()
On Error GoTo kalle
temp = JB / HScroll1.Max
temp2 = temp * HScroll1.Value
Picture1.Width = temp2

Label2.Caption = HScroll1.Value

Text1.Text = Picture1.Width / JB * 100
Exit Sub
kalle:
Picture1.Width = 0

Exit Sub
End Sub

Private Sub HScroll1_Scroll()
HScroll1_Change
End Sub


