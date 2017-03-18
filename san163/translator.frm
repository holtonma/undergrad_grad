VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      Caption         =   "END"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   5520
      Width           =   855
   End
   Begin VB.Frame frmChooseLang 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Choose Language:"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   2295
      Begin VB.OptionButton optSpanish 
         BackColor       =   &H0080C0FF&
         Caption         =   "Spanish"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optPigLatin 
         BackColor       =   &H000080FF&
         Caption         =   "PigLatin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdTranslate 
      BackColor       =   &H00C0E0FF&
      Caption         =   "TRANSLATE"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaskColor       =   &H00C0E0FF&
      TabIndex        =   6
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox lstExtractedwords 
      Height          =   1620
      Left            =   4440
      TabIndex        =   4
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox txtEnglishsentence 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   5295
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"translator.frx":0000
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   5535
   End
   Begin VB.Label lbldicttitle 
      Caption         =   "Extracted Words to be translated:"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label lblTranslatedsentence 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Label lbltranstitle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Translated Sentence"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Label lblEngtitle 
      BackColor       =   &H80000009&
      Caption         =   "Type in an English Sentence here:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00E0E0E0&
      Caption         =   "So you want to:                                               Eakspay Igpay Atinlay or Habla Espanol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Mark Holton
Rem San 163
Rem Section E
Rem 12/3/97

Option Explicit

Private Sub cmdEnd_Click()
End
End Sub

Private Sub cmdExtract_Click()
Dim i As Integer               'counter variable
Dim NumWords As Integer      'holds # words in Eng. Sent. user typed in
Dim EngString As String        'holds text in text box where user entered string
'The following array holds the individual words the user...
'typed into the text box.  Although the size of the array is...
'unspecified here, the parseline() function will resize the array...
'to the exact number of elements needed to hold the data.
Dim WordsinEngSent() As String

    'get the sentence typed into the text box
    EngString = UCase(Trim(txtEnglishsentence.Text))
    
    'separate the sentence into its individual words
    NumWords = ParseLine(EngString, WordsinEngSent(), " ")
    
    'Clear the "Extracted Words to be translated" list box...
    'and add the words from the English sentence into it.
    'Note that this step is not essential to translating the
    'sentence, rather it makes it easier for me (and the user)
    'to visualize what is going on
    lstExtractedwords.Clear
    For i = 0 To NumWords - 1
        lstExtractedwords.AddItem WordsinEngSent(i)
    Next i
    
End Sub


Private Sub cmdTranslate_Click()
    'Define variables
    Dim language As String          'variable that holds filename
    Dim FileName As String          'holds filepath
    Dim EngString As String         'holds user "inputted" sentence
    Dim i As Integer                'counter variable
    Dim Foreign(100) As String      'array that holds translated words in
                                    'sentence user typed
    Dim NumWords As Integer         'holds # words in EngString user typed
    Dim WordsinEngSent() As String  'holds the words in Eng. Sentence user typed
    Dim flagempty As Boolean
   
    lblTranslatedsentence = ""
    
    
    EngString = UCase(Trim(txtEnglishsentence.Text))
    NumWords = ParseLine(EngString, WordsinEngSent(), " ")
    'if the user does not type anything into the English sentence text...
    'box and the translate button is activated, a beep will sound...
    'and an appropriate Message Box will appear:
    'BE VERY AFRAID OF THE POWER OF THE BEEP!!
    If EngString = "" Then
        For i = 0 To 10
        Beep
        Next i
        MsgBox "Sorry, I can't translate your thoughts.  You'll need to type some words in the white box--- or I'll be forced to beep you again"
        
   Else
    'determines which of the two languages to look in depending on which option..
    'button was selected:
    If optPigLatin Then language = "PigLatin"
    If optSpanish Then language = "Spanish"
    
    'the following line finds the actual language file (on floppy) through...
    'the use of the "selectlanguage" function below
    FileName = selectlanguage(language)
    
    'the next line finds each word in turn that the user input in the text box...
    'with a loop and by...
    'calling the "UNtranslator" function, also written below
    For i = 0 To NumWords - 1
        Foreign(i) = UNtranslator(FileName, WordsinEngSent(i))
        If Foreign(i) = "" Then flagempty = True
    Next i

        'If no corresponding translated word can be found (i.e. the word...
        'is not in the dictionary file specified) a beep and an...
        'appropriate message box will occur.
        If flagempty Then
            For i = 0 To 20
            Beep
            Next i
            MsgBox "Sorry, either one or more words in that sentence doesn't exist or this program's resources can't translate something.  Try a different language."
            
            
            
        Else
            'When a correct translation can be made, it is printed in the Translated...
            'Sentence label
            For i = 0 To NumWords - 1
                lblTranslatedsentence = lblTranslatedsentence & " " & Foreign(i)
            Next i
            
        End If
    
    End If
    
End Sub


'The following is a function which tells the computer which file to look...
'under depending on which option button was selected

Private Function selectlanguage(language As String) As String

    Select Case language
        Case "PigLatin"
            selectlanguage = "A:PigLatin.Txt"
        Case "Spanish"
            selectlanguage = "A:Spanish.Txt"
    End Select
End Function

'the following function simply searches in the specified file for the...
'translated word, once the english word is known (from user input sentence)

Private Function UNtranslator(FileName As String, firstword As String) As String

    'Define Variables (Note that in all of the data files being used
    'the first word is ALWAYS the english word, and the second word is ALWAYS...
    'the translated word
        
        Dim englishwordfirst As String      'word before comma on line in specified file
        Dim foreignwordsecond As String     'word after comma
        Open FileName For Input As #1
                Do While Not EOF(1)
                    Input #1, englishwordfirst, foreignwordsecond
                    If UCase(firstword) = UCase(englishwordfirst) Then
                        UNtranslator = foreignwordsecond
                    End If
                Loop
        Close #1
        
End Function
