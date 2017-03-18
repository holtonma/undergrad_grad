VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmProgram 
      BackColor       =   &H80000016&
      Caption         =   "ACME Namefinder"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.OptionButton optSectF 
         Caption         =   "Section &F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   8
         Top             =   2640
         Width           =   1935
      End
      Begin VB.OptionButton optSectI 
         Caption         =   "Section &I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   7
         Top             =   3120
         Width           =   1575
      End
      Begin VB.OptionButton optSectE 
         Caption         =   "Section &E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   6
         Top             =   2160
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cmdFindname 
         BackColor       =   &H80000016&
         Caption         =   "Find &Name"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   5
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox txtFirstname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblDirection 
         BackColor       =   &H80000016&
         Caption         =   "Type in a First Name, specify a section, and press ""Find Name"" button (or Enter key) to find your favorite Canfield student."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   5775
      End
      Begin VB.Label lblLastnamelblonly 
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblLastname 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   2
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblFirstname 
         Caption         =   "First &Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1440
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Rem Mark Holton
Rem
Rem SAN 163 Section e
Rem
Rem 11/3/97

Private Sub cmdFindname_Click()
    'Define Variables
        Dim section As String
        Dim firstname As String
        Dim Lastname As String
        Dim filename As String

'removes spaces from before and after word
'and makes all of the letters typed in text box capitalized

  firstname = UCase(Trim(txtFirstname))
    
    'what happens if nothing is typed into the first name box...
    'and the button (or enter key) is pressed
    If firstname = "" Then
        Beep
        MsgBox "Whoops, You need to enter a valid first name."
      
    '
    Else
        'determines which section to look in depending on which option...
        'button was pressed
        If optSectE.Value Then section = "Section E"
        If optSectF Then section = "Section F"
        If optSectI Then section = "Section I"
        
        'finds the actual section file through the use of the "selectsection" subfunction...
        'which is written below
        filename = selectsection(section)
        
        'finds the last name by calling the "getname" subfunction...
        'which is written below
        Lastname = getname(filename, firstname)
        
            'This is what occurs if no corresponding last name is found...
            'by the subfunctions below
            If Lastname = "" Then
                Beep
                MsgBox "Sorry, there is no person with that first name in the specified section."
                lblLastname = ""
                
            Else
                'When a correct last name is found, it is printed in the Last Name label
                lblLastname = Lastname
                
            End If
            
    End If
    
End Sub
 
 'This is the subfunction (returns a value!!) which tells the computer...
 'which section to look under and where depending upon which option button...
 'was selected.
 
Private Function selectsection(section As String) As String

    Select Case section
        Case "Section E"
                selectsection = "A:StudentE.Txt"
        Case "Section F"
                selectsection = "A:StudentF.Txt"
        Case "Section I"
                selectsection = "A:StudentI.Txt"
                
    End Select
 
  
End Function

'This is the subfunction which finds the last name of the specified person...
'in the specified section:
'This function begins looking at the first line (EOF (1)) of the section file...
'for example in "A:StudentI.Txt", and continues down the list until it finds...
'a first name in "A:StudentI.Txt" which matches the entered first name in...
'the First Name text box.  When it does it returns the last name of that person...
'If it doesn't find a match it returns this value: "" which then is handled...
'in the above code with a message box

Private Function getname(filename As String, firstname As String) As String

    'Define Variables
        Dim Secondname As String
        Dim Primaryname As String
        
    
        Open filename For Input As #1
            Do While Not EOF(1)
                Input #1, Secondname, Primaryname
                If UCase(Primaryname) = UCase(firstname) Then
                    getname = Secondname
                End If
            Loop
        Close #1
        
End Function

        
    
        
        
                
                

        

Private Sub Option3_Click()

End Sub
