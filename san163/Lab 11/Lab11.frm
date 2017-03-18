VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Binary Search"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find It"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtFind 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Text            =   "14"
      Top             =   420
      Width           =   1515
   End
   Begin VB.ListBox lstNums 
      Height          =   2595
      Left            =   540
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFind_Click()
Dim n As Integer
n = 50
Dim Mid As Integer
Dim TrueorFalse As Integer
Dim ThisArray(n) As Single
Dim ThisQuery As Integer
Dim ThisLast As Integer
Dim ThisFirst As Integer
ThisLast = 50
ThisFirst = 0
Mid = (ThisLast + ThisFirst) \ 2



TrueorFalse = BinarySearch(ThisArray(), ThisQuery, ThisLast)
If TrueorFalse = -1 Then
    MsgBox ThisQuery & " Is not in the list!"
Else
    MsgBox "Found " & ThisQuery & " at Index: " & Mid
End If

End Sub

Private Sub Form_Load()
Dim i As Integer

    For i = 0 To 50
        lstNums.AddItem i & " -- " & 2 * i
    Next i

End Sub


Private Function BinarySearch(FindArray() As Single, Query As Integer, Last As Integer) As Integer


Dim i As Integer
Dim n As Integer
Dim FindArray(n) As Integer
Dim FoundFlag As Boolean

Dim First As Integer
Dim Middle As Integer
Dim Last As Integer

Dim Query As Integer

    For i = 0 To n
        FindArray(i) = 2 * i
    Next i

    Query = Val(txtFind)
    
    First = 0
    Last = n
    FoundFlag = False
    
    Do While Last >= First And FoundFlag = False

        Middle = (Last + First) \ 2
        
        MsgBox "First = " & First & ", Last = " & Last & ",  Middle = " & Middle
        
        Select Case Query
        
            Case FindArray(Middle)
                FoundFlag = True
        
            Case Is < FindArray(Middle)
                Last = Middle - 1
            
            Case Else
                First = Middle + 1
        
        End Select
        
    Loop
    

    If FoundFlag Then
        BinarySearch = Query
    Else
        BinarySearch = -1
    End If
    
End Function

