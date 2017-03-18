Attribute VB_Name = "Module1"
Option Explicit

' This sub draw
Public Sub DrawPizza(PizzaPic As PictureBox, NumPepperoni As Integer)
Dim i As Integer            ' Counter
Dim X As Integer            ' X coordinate of meat in TWIPs
Dim Y As Integer            ' Y coordinate of meat in TWIPs
Dim Radius As Integer       ' Radius of Pepperoni in TWIPs
Dim Offset As Integer       ' Distance from pizza center to center of pepperoni
Dim Degrees As Single       ' Compass location of pepperoni

    'Clear the pizza
    PizzaPic.Cls
    
    ' We want our Pizza to be filled in
    PizzaPic.FillStyle = vbFSSolid
    
    'Make the cheese orange
    PizzaPic.FillColor = &H80C0FF
    
    ' Draw the pizza outline
    PizzaPic.Circle (PizzaPic.Width / 2, PizzaPic.Height / 2), PizzaPic.Height / 2 - 50
    
    'Make the pepperoni Red
    PizzaPic.FillColor = &HC0&
    
    ' Draw each of the pepproni in a random location
    For i = 1 To NumPepperoni
                
        ' Locate the Pepperoni by specifying a random degree of rotation
        ' about the pizza center, and a random distance from the center
        Degrees = Rnd * 6.28  ' Woring in Radians here!
        Offset = Rnd * PizzaPic.Width / 2 - 200 ' Don't let it overlap the edge
        
        ' Do some trig to convert polar cordinates to cartesian
        X = Cos(Degrees) * Offset
        Y = Sin(Degrees) * Offset
        
        ' These cordinates are centered around the
        ' upper left corner of the picture box.
        ' Translate them to the center
        
        X = X + PizzaPic.Width / 2
        Y = Y + PizzaPic.Height / 2
                
        ' Assign a random pepperoni size beteewn 50 and 100 TWIPs
        Radius = Rnd * 50 + 50
    
        ' Draw the pepperoni
        PizzaPic.Circle (X, Y), Radius
    Next
End Sub

