VERSION 5.00
Begin VB.Form frmMultiply 
   Caption         =   "Multiplication Table"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTable 
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Table"
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmMultiply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDisplay_Click()
    Dim j As Integer
    Dim k As Integer
    
    picTable.Cls
    For j = 1 To 4
        For k = 1 To 4
            picTable.Print j; "x"; k; "="; j * k
        Next k
        picTable.Print
    Next j
End Sub
