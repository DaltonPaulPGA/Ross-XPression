' Initialize the ClockTimerWidget Name
Dim widgetName as String
widgetName = "GAME_CLOCK"

' Prompt user for time input in format HH:MM:SS, MM:SS, or SS.Z
Dim userInput As String
userInput = InputBox("Enter time to set clock (HH:MM:SS, MM:SS, or SS.Z)", "Set Clock Time")

' Check if user clicked Cancel or entered empty input
If userInput = "" Then
    msgBox("You must input a valid time to set the timer")
    Exit Sub
End If

' Split the input by colon to separate time components
Dim timeParts() As String
timeParts = Split(userInput, ":")

' Initialize hours, minutes, seconds, milliseconds
Dim hours As Long
Dim minutes As Long
Dim seconds As Long
Dim milliseconds As Long
hours = 0
minutes = 0
seconds = 0
milliseconds = 0

' Store raw input for display
Dim displayTime As String
displayTime = userInput

' Determine input format based on number of colons and presence of milliseconds
Dim lastPart() As String
If UBound(timeParts) = 2 Then
    ' Format is HH:MM:SS
    hours = CLng(timeParts(0))
    minutes = CLng(timeParts(1))
    seconds = CLng(timeParts(2))
ElseIf UBound(timeParts) = 1 Then
    ' Format is MM:SS
    minutes = CLng(timeParts(0))
    seconds = CLng(timeParts(1))
ElseIf UBound(timeParts) = 0 Then
    ' Format is SS.Z
    lastPart = Split(timeParts(0), ".")
    seconds = CLng(lastPart(0))
    If UBound(lastPart) = 1 Then
        If Len(lastPart(1)) = 1 Then
            milliseconds = CLng(lastPart(1)) * 100
        Else
            msgBox("Milliseconds must be a single digit (e.g., .2 for 200 ms)")
            Exit Sub
        End If
    End If
Else
    msgBox("Invalid time format. Use HH:MM:SS, MM:SS, or SS.Z")
    Exit Sub
End If

' Validate input
If hours < 0 Or minutes < 0 Or minutes > 59 Or seconds < 0 Or seconds > 59 Or milliseconds < 0 Or milliseconds > 999 Then
    msgBox("Invalid time values. Hours, minutes, seconds, and milliseconds must be non-negative, minutes/seconds < 60, milliseconds < 1000.")
    Exit Sub
End If

' Calculate total milliseconds
Dim totalMilliseconds As Long
totalMilliseconds = (hours * 3600000) + (minutes * 60000) + (seconds * 1000) + milliseconds

' Set the clockTimerWidget time
Dim gameClock As xpClockTimerWidget
Engine.GetWidgetByName(widgetName, gameClock)
gameClock.TimerValue = totalMilliseconds

' Confirm the action
Engine.debugMessage("Clock set to " & displayTime & " (" & totalMilliseconds & " milliseconds)", 0)
