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
   Begin VB.PictureBox picValues 
      Height          =   975
      Left            =   480
      ScaleHeight     =   915
      ScaleWidth      =   3675
      TabIndex        =   5
      Top             =   2160
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
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtEnd 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblS 
      Caption         =   "s:"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblN 
      Caption         =   "n:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frm6_3_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDisplay_Click()
    Dim n As Single
    Dim s As Single
    Dim index As Single
    ' Display values of index ranging from 0 to n step s
    picValues.Cls
    Let n = Val(txtEnd.Text)
    Let s = Val(txtStep.Text)
    For index = 0 To n Step s
        picValues.Print index;
    Next index
End Sub
