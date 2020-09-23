VERSION 5.00
Begin VB.Form frmClientUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Client Update"
   ClientHeight    =   4470
   ClientLeft      =   1530
   ClientTop       =   2010
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraComputers 
      Caption         =   "A&pplied to"
      Height          =   2835
      Left            =   60
      TabIndex        =   7
      Top             =   1530
      Width           =   5055
      Begin VB.CommandButton cmdUnSelect 
         Caption         =   "&Unselect"
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   780
         Width           =   1200
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select &All"
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   1200
      End
      Begin VB.ListBox lstClients 
         Height          =   2535
         Left            =   1410
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   180
         Width           =   3525
      End
   End
   Begin VB.Frame frmUpdate 
      Caption         =   "&Update Server Propert&ies"
      Height          =   1275
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   5055
      Begin VB.TextBox txtReplace 
         Height          =   345
         Left            =   1410
         MaxLength       =   80
         TabIndex        =   4
         Top             =   780
         Width           =   3525
      End
      Begin VB.TextBox txtFind 
         Height          =   345
         Left            =   1410
         MaxLength       =   80
         TabIndex        =   2
         Top             =   300
         Width           =   3525
      End
      Begin VB.Label lblData 
         Caption         =   "&Replace With:"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label lblData 
         Caption         =   "&From What:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5250
      TabIndex        =   6
      Top             =   630
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5250
      TabIndex        =   5
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmClientUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------
' Copyright : ICEnetware Ltd 2003 (www.ICEnetware.com)
' Module    : frmClientUpdate
' Created   : 07/11/2002
' Author    : GWilmot
' Purpose   : Provides a Global/Client Registry Search & Replace Facility
'-----------------------------------------------------------------------
' Dependancies : CBS_RegFunctions & CBS_NetFunctions
' Assumptions  : Running on Win NT/2000/XP
' Last Updated :
'-----------------------------------------------------------------------

' Interface Details
' =================

' Key Properties
' --------------
' DebugMode (Write) Sets true to set in Debugmode

' Key Public Methods
' ------------------
' RefreshForm - Configures for either as a Network Global Search & replace
'               or a single client search & replace

' Load Form/Set Refresh passing a client if selected/Show Modally

' Private Details
' ===============

' Dependancies
Private noRegFunctions As New CBS_RegFunctions  ' Registry Routines
Private noNetFunctions As New CBS_NetFunctions  ' Network Routines

' Variables
' ---------
Dim nsClient As String              ' Holds the current Client
Dim nbDebugMode As Boolean          ' Debugmode flag

Private Sub cmdOK_Click()
' OK this is where the update is triggered from

Dim lsFrom As String          ' Holds From
Dim lsWith As String          ' Holds With
Dim i As Long                 ' For counter
Dim lbContinue As Boolean     ' Flag to show if we can continue
Dim llNoUpdates As Long       ' Number of updates performed single client
Dim llNoUpdatesTotal As Long  ' Number of updates performed
Dim lbUnloadFlag As Boolean   ' Flag to say whether to unload

On Error GoTo Catch

' Set defaults
lsFrom = txtFind.Text
lsWith = txtReplace.Text

' Sanity Check
If lsFrom <> vbNullString And lsWith <> vbNullString Then
    
    ' See if we need to perform Global check
    If nsClient = vbNullString Then
        ' Where we check that at least one client is selected
        For i = 0 To lstClients.ListCount - 1
            If lstClients.Selected(i) Then
                lbContinue = True
                Exit For
            End If
        Next i
        If Not lbContinue Then
            MsgBox "Must Select at least one Client to be Updated!", vbExclamation
        End If
    Else
        lbContinue = True
    End If
    
    ' Now get the confirm
    If lbContinue Then
        Select Case MsgBox("Are you sure you want to perform the update?", vbYesNo + vbQuestion + vbDefaultButton2, "Update Registry")
        
            Case vbYes
            
                Me.MousePointer = vbHourglass
                
                ' OK see if we're doing a list
                If nsClient = vbNullString Then
                    ' Where we check that at least one client is selected
                    For i = 0 To lstClients.ListCount - 1
                        If lstClients.Selected(i) Then
                            ' OK update the selected machine
                            '*** Too scary to leave in
                            ' PerformUpdate lstClients.List(i), lsFrom, lsWith, llNoUpdatesTotal
                            llNoUpdatesTotal = llNoUpdatesTotal + llNoUpdates
                        End If
                    Next i
                Else
                    ' Just update our target machine
                    PerformUpdate nsClient, lsFrom, lsWith, llNoUpdatesTotal
                End If
                                                
                ' Tidy up
                MsgBox "Total Number of Replacements : " & CStr(llNoUpdatesTotal), vbInformation
                lbUnloadFlag = True
        
            Case vbNo
                ' just drop back
        End Select
    
    End If

Else
    MsgBox "Must set both 'From' and 'Replace' Server names!", vbExclamation
End If

Finally:
    ' Clean-up
    Me.MousePointer = vbDefault
    If lbUnloadFlag Then Unload Me
    Exit Sub

Catch:
    If nbDebugMode Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdOK_Click of Form frmClientUpdate"
    Else
        MsgBox "Problem Trying to Perform Replacement", vbExclamation
    End If
    Resume Finally
    
End Sub

Private Sub cmdCancel_Click()
' Just makes sure we unload
Unload Me
End Sub

Public Sub RefreshForm(ByVal ClientName As String)
'-----------------------------------------------------------------------
' Procedure    : frmClientUpdate.RefreshForm
' Author       : GWilmot
' Date Created : 07/11/2002
'-----------------------------------------------------------------------
' Purpose      : Updates the form either for a single client or multiple
' Assumptions  :
' Inputs       : Clientname
' Returns      :
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim lsServers() As String     ' Holds the list of servers
Dim i As Long                 ' For Counter

' Server types constants
Const SV_TYPE_ALL                 As Long = &HFFFFFFFF
Const GLOBAL_HEIGHT As Long = 4875
Const CLIENT_HEIGHT As Long = 1890

On Error GoTo Catch

' Copy to near properties
nsClient = ClientName

' Sort out caption & properties
If nsClient = vbNullString Then
    fraComputers.Enabled = True
    Me.Caption = Me.Caption + " : NETWORK GLOBAL"
    Me.Height = GLOBAL_HEIGHT
    
    ' OK lets load the computers
    If noNetFunctions.GetServers(vbNullString, lsServers, SV_TYPE_ALL) Then
        'ShellSortStrings lsServers
        For i = 1 To UBound(lsServers)
            lstClients.AddItem lsServers(i)
        Next i
    End If
Else
    fraComputers.Enabled = False
    Me.Caption = Me.Caption + " : " + nsClient
    Me.Height = CLIENT_HEIGHT
End If

' Centre the form
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

Finally:
    ' Clean-up
    Exit Sub

Catch:
    ' Only display developer error in debug mode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RefreshForm of Form frmClientUpdate"
    Resume Finally

End Sub

Public Property Let DebugMode(ByVal Flag As Boolean)
' Simple Flag to say if object in debug mode
nbDebugMode = Flag
End Property

Private Sub cmdSelectAll_Click()
' Just selects all workstations
Dim i As Long         ' for counter

On Error GoTo Catch

Me.MousePointer = vbHourglass

For i = 0 To lstClients.ListCount - 1
    lstClients.Selected(i) = True
Next i

Finally:
    ' Clean-up
    Me.MousePointer = vbDefault
    Exit Sub

Catch:

    Resume Finally

End Sub

Private Sub cmdUnSelect_Click()
' Just selects all workstations
Dim i As Long         ' for counter

On Error GoTo Catch

Me.MousePointer = vbHourglass

For i = 0 To lstClients.ListCount - 1
    lstClients.Selected(i) = False
Next i

Finally:
    ' Clean-up
    Me.MousePointer = vbDefault
    Exit Sub

Catch:

    Resume Finally

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Catch

'  Release the objects
Set noRegFunctions = Nothing
Set noNetFunctions = Nothing

Finally:
    Exit Sub
Catch:
    Resume Finally

End Sub

Public Sub PerformUpdate(ByVal Computer As String, ByVal Find As String, _
    ByVal ReplaceWith As String, ByRef ReplaceCount As Long)
'-----------------------------------------------------------------------
' Procedure    : frmClientUpdate.PerformUpdate
' Author       : GWilmot
' Date Created : 07/11/2002
'-----------------------------------------------------------------------
' Purpose      : Replaces Server name in a given machine
' Assumptions  : That we replace set parameters
' Inputs       : Computer, Text to find and replace with
' Returns      : Count of replacements
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim lsdata As String        ' Holds registry setting
Dim i As Long               ' for counter
Dim lsItems() As String     ' Holds items

On Error GoTo Catch

' Defaults
ReplaceCount = 0

'*** This is hardcoded to go thru set links and replace if required

' Product's

' Update Service Desk
If noRegFunctions.GetRemoteRegistry(Computer, "SOFTWARE\Enron Europe\Enron Direct\Service Desk", "Server", lsdata) Then
    If UCase$(lsdata) = UCase$(Find) Then
        If noRegFunctions.SetRemoteRegistry(Computer, "SOFTWARE\Enron Europe\Enron Direct\Service Desk", "Server", ReplaceWith) Then
            ReplaceCount = ReplaceCount + 1
        End If
    End If
End If

' Update Sales Desk
If noRegFunctions.GetRemoteRegistry(Computer, "SOFTWARE\Enron Europe\Enron Direct\Sales Desk", "Server", lsdata) Then
    If UCase$(lsdata) = UCase$(Find) Then
        If noRegFunctions.SetRemoteRegistry(Computer, "SOFTWARE\Enron Europe\Enron Direct\Sales Desk", "Server", ReplaceWith) Then
            ReplaceCount = ReplaceCount + 1
        End If
    End If
End If

' DSNs
If noRegFunctions.GetRemoteDSNs(Computer, lsItems) Then
    ' Need to get each server in turn
    For i = LBound(lsItems) To UBound(lsItems)
        If noRegFunctions.GetRemoteRegistry(Computer, "SOFTWARE\ODBC\ODBC.INI\" + lsItems(i), "Server", lsdata) Then
            If UCase$(lsdata) = UCase$(Find) Then
                If noRegFunctions.SetRemoteRegistry(Computer, "SOFTWARE\ODBC\ODBC.INI\" + lsItems(i), "Server", ReplaceWith) Then
                    ReplaceCount = ReplaceCount + 1
                End If
            End If
        End If
    Next i
End If

Finally:
    ' Clean-up
    Exit Sub

Catch:
    ' Just report back to developer
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PerformUpdate of Form frmClientUpdate"
    Resume Finally

End Sub
