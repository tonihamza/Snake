' WARNING:
' This code is divided into three parts and must be placed in three different modules:
' 1. The first part goes into "ThisWorkbook" (color selection of the snake).
' 2. The second part is the main module containing game functionalities like snake movement and apple generation.
' 3. The third part is the timer module that controls the snake's movement intervals.

' ------------------------------------------
' Part 1: This code goes in "ThisWorkbook"
' It is the code for the color selection of the snake.

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Set the snake's color based on the RGB values in the cells AM4, AO4, AQ4 for the first part,
    ' and AM5, AO5, AQ5 for the second part of the snake.
    Range("AS4").Interior.Color = RGB(Range("AM4").Value, Range("AO4").Value, Range("AQ4").Value)
    Range("AS5").Interior.Color = RGB(Range("AM5").Value, Range("AO5").Value, Range("AQ5").Value)
End Sub

' ------------------------------------------
' Part 2: Main Module - This module contains all the functionalities of the game
' such as apple generation and snake movement.

Public snake1(2000) As Range  ' Array to store the snake's body positions.
Public col As Integer         ' Stores the horizontal movement (column).
Public row As Integer         ' Stores the vertical movement (row).
Public apple0 As Range        ' Stores the position of the apple.
Public score As Integer       ' Stores the player's score.
Public R1 As Integer          ' Stores the red color value for the snake's head.
Public G1 As Integer          ' Stores the green color value for the snake's head.
Public B1 As Integer          ' Stores the blue color value for the snake's head.
Public R2 As Integer          ' Stores the red color value for the snake's body.
Public G2 As Integer          ' Stores the green color value for the snake's body.
Public B2 As Integer          ' Stores the blue color value for the snake's body.
Dim k As Long                 ' Length of the snake (number of segments).

Public Sub Snake()
    Dim i As Integer
    ' Reset the color of the tail segment.
    snake1(k).Interior.Color = RGB(255, 255, 255)
    
    ' Move the snake's body.
    For i = k To 1 Step -1
        Set snake1(i) = snake1(i - 1)
    Next i
    
    ' Check if the snake hits itself or the boundaries of the grid.
    If snake1(0).Offset(row, col).Interior.Color = RGB(R2, G2, B2) Or Intersect(snake1(0).Offset(row, col), Range("N8:AQ37")) Is Nothing Then
        StopTimer
        MsgBox ("GAME OVER")
        ' Reset the keys used for movement.
        Application.OnKey "{UP}"
        Application.OnKey "{DOWN}"
        Application.OnKey "{LEFT}"
        Application.OnKey "{RIGHT}"
        Exit Sub
    End If
    
    ' Check if the snake eats the apple.
    If snake1(0).Offset(row, col).Interior.Color = RGB(255, 255, 0) Then
        score = score + 10
        Set snake1(0) = snake1(0).Offset(row, col)
        k = k + 1
        ' Add a new segment to the snake.
        Set snake1(k) = snake1(k - 1).Offset(row * -1, col * -1)
        Apple
    Else
        ' Continue moving the snake.
        Set snake1(0) = snake1(0).Offset(row, col)
    End If
    
    Color
End Sub

Public Sub Start()
    ' Assign movement controls to arrow keys.
    Application.OnKey "{UP}", "Up"
    Application.OnKey "{DOWN}", "Down"
    Application.OnKey "{LEFT}", "Left"
    Application.OnKey "{RIGHT}", "Right"
    
    ' Clear the game grid.
    Dim i As Variant
    For Each i In Range("N8:AQ37")
        i.Interior.ColorIndex = xlColorIndexNone
    Next i
    
    ' Initialize the snake and game settings.
    k = 2
    score = 0
    Set snake1(0) = Range("AB18")
    Set snake1(1) = Range("AB19")
    Set snake1(2) = Range("AB20")
    Set apple0 = Range("N8")
    col = 0
    row = 0
    
    ' Load snake's head and body colors.
    R1 = Range("AM4").Value
    R2 = Range("AM5").Value
    G1 = Range("AO4").Value
    G2 = Range("AO5").Value
    B1 = Range("AQ4").Value
    B2 = Range("AQ5").Value
    
    Range("R4").Font.Size = 18
    Apple
    Color
    StartTimer
End Sub

Public Sub Color()
    ' Color the snake's body.
    Dim i As Variant
    For i = k To 1 Step -1
        snake1(i).Interior.Color = RGB(R2, G2, B2)
    Next i
    
    ' Color the snake's head.
    snake1(0).Interior.Color = RGB(R1, G1, B1)
    
    ' Color the apple.
    apple0.Interior.Color = RGB(255, 255, 0)
    
    ' Update the score display.
    Range("R4").Value = score
End Sub

Public Sub Apple()
    ' Generate random positions for the apple.
    col1 = Application.WorksheetFunction.RandBetween(0, 29)
    row1 = Application.WorksheetFunction.RandBetween(0, 29)
    
    ' Ensure the apple doesn't spawn on the snake's body.
    If Not Range("N8").Offset(row1, col1).Interior.Color = RGB(R1, G1, B1) And Not Range("N8").Offset(row1, col1).Interior.Color = RGB(R2, G2, B2) Then
        Set apple0 = Range("N8").Offset(row1, col1)
        Range("N8").Offset(row1, col1).Select
    Else
        Apple
    End If
End Sub

' Movement functions (Up, Down, Left, Right).
Public Sub Up()
    col = 0
    row = -1
    Snake
End Sub

Public Sub Down()
    col = 0
    row = 1
    Snake
End Sub

Public Sub Left()
    col = -1
    row = 0
    Snake
End Sub

Public Sub Right()
    col = 1
    row = 0
    Snake
End Sub

' ------------------------------------------
' Part 3: Timer Module - This module controls the game's timer.

Option Explicit

#If Win64 Then
    ' 64-bit declaration of Windows API functions for the timer.
    Public Declare PtrSafe Function SetTimer Lib "User32" ( _
        ByVal hwnd As LongLong, _
        ByVal nIDEvent As LongLong, _
        ByVal uElapse As LongLong, _
        ByVal lpTimerFunc As LongLong) As LongLong
    Public Declare PtrSafe Function KillTimer Lib "User32" ( _
        ByVal hwnd As LongLong, _
        ByVal nIDEvent As LongLong) As LongLong
    Public TimerID As LongLong
#Else
    ' 32-bit declaration of Windows API functions for the timer.
    Public Declare PtrSafe Function SetTimer Lib "User32" ( _
        ByVal hwnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long
    Public Declare PtrSafe Function KillTimer Lib "User32" ( _
        ByVal hwnd As Long, _
        ByVal nIDEvent As Long) As Long
    Public TimerID As Long
#End If

' Starts the game timer.
Sub StartTimer()
    If TimerID <> 0 Then
        KillTimer 0, TimerID
        TimerID = 0
    End If
    TimerID = SetTimer(0, 0, Range("P3").Value, AddressOf TimerEvent)
End Sub

' Timer event calls the Snake subroutine to move the snake.
Sub TimerEvent()
    On Error Resume Next
    Call Snake
    Exit Sub
End Sub

' Stops the game timer.
Sub StopTimer()
    KillTimer 0, TimerID
    TimerID = 0
End Sub
