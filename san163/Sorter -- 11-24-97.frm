VERSION 5.00
Begin VB.Form frmSorter 
   Caption         =   "Sorter"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort 
      Caption         =   "&Sort"
      Height          =   495
      Left            =   2700
      TabIndex        =   6
      Top             =   1620
      Width           =   1215
   End
   Begin VB.TextBox txtSort 
      Height          =   375
      Index           =   5
      Left            =   420
      TabIndex        =   5
      Top             =   2820
      Width           =   1395
   End
   Begin VB.TextBox txtSort 
      Height          =   375
      Index           =   4
      Left            =   420
      TabIndex        =   4
      Top             =   2340
      Width           =   1395
   End
   Begin VB.TextBox txtSort 
      Height          =   375
      Index           =   3
      Left            =   420
      TabIndex        =   3
      Top             =   1860
      Width           =   1395
   End
   Begin VB.TextBox txtSort 
      Height          =   375
      Index           =   2
      Left            =   420
      TabIndex        =   2
      Top             =   1380
      Width           =   1395
   End
   Begin VB.TextBox txtSort 
      Height          =   375
      Index           =   1
      Left            =   420
      TabIndex        =   1
      Text            =   "Bean"
      Top             =   900
      Width           =   1395
   End
   Begin VB.TextBox txtSort 
      Height          =   375
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Text            =   "Potato"
      Top             =   420
      Width           =   1395
   End
End
Attribute VB_Name = "frmSorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Swap(Val1 As TextBox, Val2 As TextBox)
Dim Temp As String
    
    Temp = Val1
    Val1 = Val2
    Val2 = Temp

End Sub

Private Sub cmdSort_Click()
Dim Temp As String
Dim Passes As Integer
Dim Compares As Integer
Dim SwapFlag As Boolean

    For Passes = 1 To 5
        SwapFlag = False
        For Compares = 1 To 6 - Passes
            If txtSort(Compares) < txtSort(Compares - 1) Then
                Call Swap(txtSort(Compares), txtSort(Compares - 1))
                SwapFlag = True
            End If
        Next Compares
        If Not SwapFlag Then Exit For
    Next Passes

End Sub
