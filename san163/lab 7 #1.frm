VERSION 5.00
Begin VB.Form FRM6_3_1 
   Caption         =   "POPULATION GROWTH"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTable 
      Height          =   1935
      Left            =   960
      ScaleHeight     =   1875
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Population"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FRM6_3_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDisplay_Click()
    Dim pop As Single
    Dim yr As Integer
    ' Display population from 1990 to 1995
    
    picTable.Cls
    Let pop = 300000
    For yr = 1990 To 1995
        picTable.Print yr, Format(pop, "#,#")
        Let pop = pop + 0.03 * pop
    Next yr
End Sub
