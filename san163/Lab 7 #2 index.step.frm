VERSION 5.00
Begin VB.Form frm6_3_2 
   Caption         =   "For index = 0 To n Step s"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      Caption         =   "exit"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.PictureBox picValues 
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   2040
      Width           =   3735
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Values of index"
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox txtStep 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtEnd 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblS 
      Caption         =   "Addition step value (from O): DO NOT INPUT ""0""!"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblN 
      Caption         =   "Number index can't exceed:"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frm6_3_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Mark Holton
Rem San Lab 7#2
Option Explicit

Private Sub cmdDisplay_Click()
    Dim n As Single
    Dim s As Single
    Dim index As Single
    ' Display values of index ranging from 0 to n step s
    picValues.Cls
    Let n = Trim(Val(txtEnd.Text))
    Let s = Trim(Val(txtStep.Text))
    For index = 0 To n Step s
        picValues.Print index;
    Next index
End Sub

Private Sub cmdEnd_Click()
End

End Sub
