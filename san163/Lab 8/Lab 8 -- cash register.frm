VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Visual Basic Calculator"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1020
      TabIndex        =   3
      Top             =   2580
      Width           =   1275
   End
   Begin VB.PictureBox picTape 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3780
      ScaleHeight     =   3315
      ScaleWidth      =   1455
      TabIndex        =   1
      Top             =   300
      Width           =   1515
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -120
         TabIndex        =   2
         Top             =   3060
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   3060
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Visual Basic Calculator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   360
      TabIndex        =   4
      Top             =   300
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
Dim i As Integer    ' Counter Variable
Dim Sum As Single   ' Total of all numbers entered so far
Dim StartVal As Integer  ' First number of For-Next loop
Dim StopVal As Integer   ' Last  number of For-Next loop

' Note the these variables are created with Static instead of Dim.
' Static tells Visual Basic to remember the value of the variable
' between calls to a sub or function.  In this case, Visual Basic
' Will preserve the values in the array as well as the variable
' that tracks the number of values that have been entered.
Static Numbers(100) As Single  ' Array that will hold the numbers entered so far
Static NumNumbers As Integer   ' Variable that keeps track of how
                               ' many numbers have been entered

    ' When NumNumbers is created it has a value or 0.
    ' Increment it here as one more numbers is being added
    'Your code here...
    NumNumbers = NumNumbers + 1
    
    ' NumNumbers can be used to determine which element of the
    ' array should get the next value from the text box.
    ' Write a line of code here that assigns the value of the
    ' text box to the next available arrray element.
    ' Use NumNumbers as the index of the array.
    'Your code here...
    Numbers(NumNumbers) = Val(txtInput.Text)
    
    
    ' It's time to update the picture box, first clear it here
    'Your code here...
    picTape.Cls
    
    
    
    ' Now you need to print all of the values that have been
    ' entered so far into the picture box.  You will do this
    ' using a For-Next loop.  What should the starting and
    ' stopping values be of the 'For' line?  Put these values
    ' (they may be variables) after the '=' sign on the next
    ' two lines
    StartVal = Numbers(0)
    StopVal = Numbers(100)
    
    
    ' I Started the loop for you.  Finish the loop, and put
    ' a line of code in it that will print out one of the
    ' array values.  Use i as the array index.
    For i = StartVal To StopVal
      picTape.Print Numbers(NumNumbers)
      Next i
      
    
    ' Now modify the For-Next loop, so that it keeps a running
    ' total of all of the values in the array. If the total is
    ' kept in Sun, then the next line will update the yellow
    ' label with the total
    lblTotal = "Total: " & Sum

End Sub

' This code get run every time the
' user presses a key in the text box
Private Sub txtInput_KeyPress(KeyAscii As Integer)
    
    ' Calls the sub for the command button if the
    ' User hits the enter key after typing a number
    ' It then blanks out the label so a new number
    ' can be typed quickly
    If KeyAscii = 13 Then
        cmdAdd_Click    ' Call the buttton
        txtInput = ""   ' Blank out the label
        KeyAscii = 0    ' This prevents a beep
    End If
End Sub
