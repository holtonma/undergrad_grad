VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtPercent 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Text            =   "95"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtLetter 
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Round(Number As Single)

    Number = Int(Number + 0.5)
    
End Sub

Private Function LetterGrade(Percent As Single) As String
Dim Letter As String

    Select Case Percent
        Case Is >= 90
            Letter = "A"
            
        Case 80 To 90
            Letter = "B"
            
        Case 70 To 80
            Letter = "C"
            
        Case 60 To 70
            Letter = "D"
            
        Case Else
            Letter = "F"
    End Select
                    
    Percent = Percent Mod 10
                        
    
    If Not (Letter = "F" Or Letter = "A") Then
        Select Case Percent
            Case 0 To 3
                Letter = Letter & "-"
            
            Case 7, 8, 9
                Letter = Letter & "+"
        End Select
    End If

    LetterGrade = Letter
    
End Function


Private Sub cmdConvert_Click()
Dim Percent As Single
                    
    
    Percent = Val(txtPercent)

    Call Round(Percent)
    
    txtLetter = LetterGrade(Percent)
               
End Sub
