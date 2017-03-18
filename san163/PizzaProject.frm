VERSION 5.00
Begin VB.Form Project2 
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   5280
      TabIndex        =   11
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdDrawpizza 
      Caption         =   "Draw the Pizza"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2535
      Begin VB.TextBox txtPepperonislices 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CheckBox chkVeggies 
         Caption         =   "Veggies                 $1.00"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CheckBox chkPineapple 
         Caption         =   "Pineapple              $0.75"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CheckBox chkExtracheese 
         Caption         =   "Extra Cheese         $1.50"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox chkOlives 
         Caption         =   "Olives                    $0.75"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblPeppslices 
         Caption         =   "# Pepperoni Slices"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
      End
   End
   Begin VB.PictureBox picPizza 
      Height          =   3135
      Left            =   3360
      ScaleHeight     =   3075
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label lblCompletecost 
      Caption         =   "Total Cost ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   12
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblBasecost 
      Caption         =   "Base Cost = $8.50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblTotalcost 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Project2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Rem Mark Holton
Rem 10/21/97
Rem
Rem SAN 163 - Program #2
Rem

Private Sub cmdDrawpizza_Click()
    'define variables
        Dim extracheese As Integer
        Dim olives As Integer
        Dim pineapple As Integer
        Dim veggies As Integer
        Dim totalcost As Single
        
    'define values from checkboxes
        extracheese = chkExtracheese.Value
        olives = chkOlives.Value
        pineapple = chkPineapple.Value
        veggies = chkVeggies.Value
        
    'Validates the input from the pepperoni text box
        If IsNumeric(txtPepperonislices) Then
            
            'this line makes the textbox whit
                txtPepperonislices.BackColor = vbWindowBackground
                
            'Calls total cost from function
                totalcost = pizzaprice(extracheese, olives, pineapple, veggies)
                
            'calls update function and updates the label
                Call update(totalcost)
                
        Else
            Beep
            
            'This line makes the textbox red
            txtPepperonislices.BackColor = vbRed
        
        End If
End Sub

Private Function pizzaprice(extracheese As Integer, olives As Integer, pineapple As Integer, veggies As Integer) As Single

    'define variable
        Dim total As Single
        
    'calculate total cost of pizza
        total = 8.5
        If extracheese = 1 Then total = total + 1.5
        If olives = 1 Then total = total + 0.75
        If pineapple = 1 Then total = total + 0.75
        If veggies = 1 Then total = total + 1#
        
        pizzaprice = total

End Function

Private Sub update(totalcost As Single)

    'Draws pizza on screen
        DrawPizza picPizza, Val(txtPepperonislices)
        
    'updates total cost on screen
        lblTotalcost = Format(totalcost, "Currency")
        
End Sub

        

        
    
        



Private Sub cmdQuit_Click()
End

End Sub

