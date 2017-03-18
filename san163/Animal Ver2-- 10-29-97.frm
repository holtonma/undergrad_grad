VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "He Likes:"
      Height          =   495
      Left            =   780
      TabIndex        =   6
      Top             =   1500
      Width           =   1215
   End
   Begin VB.TextBox txtFood5 
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtFood4 
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtFood3 
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtFood2 
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   780
      Width           =   1215
   End
   Begin VB.TextBox txtFood1 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   300
      Width           =   1215
   End
   Begin VB.TextBox txtAnimal 
      Height          =   315
      Left            =   780
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()
Dim Animal As String
Dim F(10) As String

    Animal = txtAnimal
    Call Likes(Animal, F())
    
    txtFood1 = F(1)
    txtFood2 = F(2)
    txtFood3 = F(3)
    txtFood4 = F(4)
    txtFood5 = F(5)

End Sub

Private Sub Likes(Animal As String, Food() As String)

    Select Case Animal
        Case "Dog"
            Food(1) = "Meat"
            Food(2) = "Dog Fod"
            Food(3) = "Bones"
            Food(4) = "Pizza"
            
        Case "Zebra"
            Food(1) = "Grass"
            Food(2) = "Carrots"
            Food(3) = "Small Plants"
    
    End Select



End Sub

