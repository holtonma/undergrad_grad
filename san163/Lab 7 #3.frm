VERSION 5.00
Begin VB.Form frm6_3_3 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTranspose 
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   1920
      Width           =   3975
   End
   Begin VB.CommandButton cmdReverse 
      Caption         =   "Reverse Letters"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin VB.TextBox txtWord 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label lblWord 
      Caption         =   "Enter Word"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frm6_3_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReverse_Click()
    picTranspose.Cls
    picTranspose.Print Reverse(txtWord.Text)
End Sub

Private Function Reverse(info As String) As String
    Dim m As Integer
    Dim j As Integer
    Dim temp As String
    
    Let m = Len(info)
    Let temp = ""
    For j = m To 1 Step -1
        Let temp = temp & Mid(info, j, 1)
    Next j
    Reverse = temp
End Function

