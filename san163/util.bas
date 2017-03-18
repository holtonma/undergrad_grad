Attribute VB_Name = "Module1"
Option Explicit

' This DLL allows toggling the lines on the Parallel port
Declare Function Inp Lib "InpOut.DLL" (ByVal Port%) As Integer
Declare Sub Out Lib "InpOut.DLL" (ByVal Port%, ByVal Value%)

' Dependants:
' Update this table when adding this file to a project
'
'|------------------------------------|
'| Project    |  This File  |  Copy   |
'|------------------------------------|
'|  Taper     |             |    X    |
'|  Level     |      X      |         |
'|  Tram      |      X      |         |
'|  Square    |      X      |         |
'|  Palind    |      X      |         |
'|  Spindle   |      X      |         |
'|  B-Axis    |      X      |         |
'|  CalApp    |      X      |         |
'|------------------------------------|

'Returns the path to this program a consistent form
'(i.e. if the program in in the root directory)
'
Function AppPath(ByVal FileName$) As String
    'VB puts a "\" on the root dir (ie "c:\")
    If Right$(App.Path, 1) = "\" Then
        AppPath = App.Path & FileName
    Else
        AppPath = App.Path & "\" & FileName
    End If
End Function

Sub CenterForm(Frm As Form)
    ' Center the form
    Frm.Left = (Screen.Width - Frm.Width) / 2
    Frm.Top = (Screen.Height - Frm.Height) / 2
End Sub

Function Deg2Rad(Degs As Double) As Double
    Deg2Rad = Degs * 3.14159265358 / 180
End Function

' Wait NumSecs Seconds
' Function only valid for NumSecs < 86400
Sub Delay(NumSecs As Double)
Dim StartTime As Double
Dim EndTime As Double
Dim DayOffset As Double

    ' Remember where we're at
    StartTime = Timer
    EndTime = StartTime + NumSecs
    
    ' See if we have to wait past midnight
    If EndTime >= 86400 Then
        
        ' Calculate stopping time tomorrow
        EndTime = EndTime - 86400
        
        ' Wait for Midnight
        While StartTime < Timer: DoEvents: Wend
    End If
    
    ' Finish waiting
    While Timer < EndTime: DoEvents: Wend
End Sub

Function FindListItem(ByVal FindItem As String, Flist As Control) As Integer
Dim Index As Integer
Dim TestString As String
Dim Pos As Integer

    ' If the list is empty then drop out
    If Flist.ListCount = 0 Then
        FindListItem = -1
        Exit Function
    End If

    ' Do a case insensitive search
    FindItem = UCase(Trim(FindItem))

    ' Test the TestString against every entry in the listbox
    For Index = 0 To Flist.ListCount - 1
        ' If the lisbox has tabs, then only compare up to the first tab
        Pos = InStr(Flist.List(Index), Chr(9))
        If Pos Then
            TestString = Left(Flist.List(Index), Pos - 1)
        Else
            TestString = Flist.List(Index)
        End If

        ' Do a case insensitive search
        If FindItem = UCase(Trim(TestString)) Then
            FindListItem = Index
            Exit Function
        End If
    Next Index

    ' The string was not found
    FindListItem = -1

End Function

' Returns the Portion after the "|" in a tag
' This is the format used for elastic controls
' when the tag is being used as a label.
Function GetTag(Mycontrol As Control) As Variant
Dim Pos As Integer

    ' Find the "|"
    Pos = InStr(Mycontrol.Tag, "|")
    
    ' Return everything after the "|"
    If Pos = 0 Or Pos = Len(Mycontrol.Tag) Then
        GetTag = ""
    Else
        GetTag = Right(Mycontrol.Tag, Len(Mycontrol.Tag) - Pos)
    End If

End Function

Sub HighLite(Mycontrol As Control)
    Mycontrol.SelStart = 0
    Mycontrol.SelLength = Len(Mycontrol.Text)
End Sub

Function IsEven(Number As Integer) As Boolean
    IsEven = ((Number Mod 2) = 0)
End Function

'This function takes a delimited line of data and
'Separates it into its fields.
'Label$() is dimentioned to the number of fields found
'and contains the individual fields.
'Inline$ is the line to be parsed, and Delim$ is the delimiter to use
'Returns number of fields found in Line
Function ParseLine(InLine As String, Labels() As String, Delim As String) As Long
Dim i%, X%, Y%
Dim MyString$
 
    ' If a blank string was sent then return 0 and drop out
    If Trim(InLine) = "" Then
        ParseLine = 0
        Exit Function
    End If

    'Locate the first delimiter and setup the loop
    i = 0
    X = 1
    Y% = InStr(InLine, Delim)
    While Y
        ReDim Preserve Labels$(i)
        Labels$(i) = Trim(Mid(InLine, X, Y - X))
        X = Y + Len(Delim)
        Y = InStr(X, InLine, Delim)
        i = i + 1
    Wend
    
    'This happens when there is no trailing delimeter
    If X <= Len(Trim(InLine)) Then
        Y = Len(InLine) + 1
        ReDim Preserve Labels$(i)
        Labels$(i) = Trim(Mid(InLine, X, Y - X))
    End If

    ParseLine = UBound(Labels$) + 1
         
End Function

' Prints text centered between X and W at Height Y
Sub PrintCenter(X!, W!, Y!, PrintText As String)
    Printer.CurrentX = X + (W - X - Printer.TextWidth(PrintText)) / 2
    Printer.CurrentY = Y
    Printer.Print PrintText
End Sub

' Prints line of Test information at current location like this:
'
' Title
' --------------------
' SubTitle
'
Sub PrintInfoLine(SubTitle$, Title$)
Dim FontPoints As Single
Dim FontHeight As Single
Dim FontBold As Integer
Dim FontItalic As Integer
Dim StartX As Single

    ' Save the initial settings
    FontPoints = Printer.FontSize
    FontBold = Printer.FontBold
    FontHeight = Printer.TextHeight("|")
    FontItalic = Printer.FontItalic
    StartX = Printer.CurrentX

    ' Print the line and sub title
    Printer.Line -Step(2, 0)
    Printer.CurrentX = StartX
    Printer.CurrentY = Printer.CurrentY + 0.02
    Printer.FontSize = 8.25
    Printer.FontBold = True
    Printer.FontItalic = True
    Printer.Print SubTitle;
    
    ' Print the title
    Printer.CurrentX = StartX
    Printer.FontSize = FontPoints
    Printer.FontItalic = False
    Printer.CurrentY = Printer.CurrentY - FontHeight - 0.02
    Printer.Print Title
    
    ' Restore the initial settings
    Printer.FontItalic = FontItalic
    Printer.FontBold = FontBold
    Printer.CurrentX = StartX
    Printer.CurrentY = Printer.CurrentY + 2 * FontHeight + 0.04

End Sub

'
' Prints Right-Justified text at X
Sub PrintRight(X As Single, PrintText As String)
    Printer.CurrentX = X - Printer.TextWidth(PrintText)
    Printer.Print PrintText;
End Sub

' Appends TagData to the MyControl.tag after the "|"
' Supplies the "|" if missing.
' This is the format used for elastic containers
' when the tag is being used as a label.
Sub PutTag(Mycontrol As Control, TagData As Variant)
Dim Pos As Integer
Dim OldTag As String

    ' Find the "|"
    Pos = InStr(Mycontrol.Tag, "|")

    ' Extract the original Tag before the "|"
    Select Case Pos
        Case 0 ' no "|"
            OldTag = Mycontrol.Tag
        
        Case 1 ' Only "|"
            OldTag = ""
            
        Case Else  ' More than "|"
            OldTag = Left(Mycontrol.Tag, Pos - 1)
    End Select
        
    ' Reset the tag
    Mycontrol.Tag = OldTag & "|" & TagData

End Sub

Function Rad2Deg(Rads As Double) As Double
    Rad2Deg = Rads * 180 / 3.14159265358
End Function

'This function take a number of minutes and seconds and
'reduces the minutes and seconds parts to lowest form.
'it returns the total number of seconds
Function ReduceMinutes(Minutes As Long, Seconds As Long) As Long
Dim Secs As Long
    Secs = 60 * Minutes + Seconds

    Minutes = Secs \ 60
    Seconds = Secs Mod 60
    ReduceMinutes = Secs
End Function

' Returns the last componet of the path
' Outpath is what preceeded it.
' If non-decomposable then return "" and OutPath = InPath
'
' Given:    "C:\vb\data\file.dat"
'           Return: "file.dat"
'           OutPath = "C:\vb\data"
'
' Given:    "C:\vb\data\"
'           Return: "data"
'           OutPath = "C:\vb\"
'
' Given:    "C:\"
'           Return: ""
'           OutPath = "C:\"
Function SplitPath(InPath As String, OutPath As String) As String
Dim X As Integer, Y As Integer

    ' Start by setting OutPtah to InPath
    OutPath = InPath
    
    ' If it is a root path then drop out
    If Len(InPath) < 4 Then SplitPath = "": Exit Function
    
    ' Strip final '\' if it is there
    If Right(InPath, 1) = "\" Then OutPath = Left$(InPath, Len(InPath) - 1)

    ' Find the position of the last "\"
    Y = 0
    Do
        X = InStr(Y + 1, OutPath, "\")
        If X Then Y = X
    Loop Until X = 0

    ' Split the InPath at the last "\"
    If Y Then
        SplitPath = Right(OutPath, Len(OutPath) - Y)
        OutPath = Left(OutPath, Y)
    Else
        OutPath = ""
        SplitPath = ""
    End If

End Function

