Attribute VB_Name = "AppMain"
Option Explicit
'-----------------------------------------------------------------------
' Copyright : ICEnetware Ltd 2003 (www.ICEnetware.com)
' Module    : AppMain
' Created   : 06/12/2002
' Author    : GWilmot
' Purpose   : Main Module of the netapi32.dll Examples
'-----------------------------------------------------------------------
' Dependancies : FrmMain
' Assumptions  :
' Last Updated :
'-----------------------------------------------------------------------

Public Sub Main()
'-----------------------------------------------------------------------
' Procedure    : AppMain.Main
' Author       : GWilmot
' Date Created : 06/12/2002
'-----------------------------------------------------------------------
' Purpose      : This is starting point of the Application
' Assumptions  : That this has been CORRECTLY set in Project Properties
' Inputs       :
' Returns      :
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim lbDebugMode As Boolean      ' Flag to say if we are in debug mode
Dim ldStart As Double           ' Stores the start time for the splash screen

On Error GoTo Catch

' Perform app instance check before doing anything
If App.PrevInstance Then
    ' This just displays a critical message
    ' However a more sophisicated approach would be to switch to the
    ' other application version
    MsgBox "Application already started!", vbCritical
    End
End If

' Give the user some feedback
Screen.MousePointer = vbHourglass

' Add Splash Screen (Load/Show)
Load frmSplash
frmSplash.Show vbModeless
frmSplash.Refresh

' Make sure its displayed for at least a second
ldStart = Timer
Do
Loop Until Abs(Timer - ldStart) > 1

' Get DebugMode flag from the Command Line
If InStr(UCase$(Command$), "\DEBUGMODE") > 0 Then lbDebugMode = True

' Add in Application Main Form Load
Load frmMain

' Set properties
frmMain.DebugMode = lbDebugMode

' Only make ontop AFTER all error messages could have been displayed
frmSplash.SetTopMost True
frmSplash.Refresh

' Show Main form (Modelessly!)
frmMain.Show vbModeless

' Splash displayed over the top for half a second
ldStart = Timer
Do
Loop Until Abs(Timer - ldStart) > 0.5
' Drop Splash
Unload frmSplash
Set frmSplash = Nothing

Finally:
    ' Clean-up
    Screen.MousePointer = vbDefault
    Exit Sub

Catch:
    ' If a failure occurs during the start of an application - end (If you can't start then stop)
    Screen.MousePointer = vbDefault
    MsgBox "Fatal error occurred during Application Load!" & vbCrLf & Err.Description, vbCritical
    End

End Sub

