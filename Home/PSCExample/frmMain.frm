VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Demo using netapi32.dll"
   ClientHeight    =   6195
   ClientLeft      =   2520
   ClientTop       =   3420
   ClientWidth     =   7650
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   7650
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Displays the progress of collecting the Schema"
      Top             =   5910
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   503
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12991
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDisplay 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5855
      Left            =   3450
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   40
      Width           =   4200
   End
   Begin MSComctlLib.TreeView trvNetwork 
      CausesValidation=   0   'False
      Height          =   5865
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   10345
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imgNetwork"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgNetwork 
      Left            =   1050
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuRefreshNetwork 
         Caption         =   "&Refresh Network"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpDummy01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------
' Copyright : ICEnetware Ltd 2003 (www.ICEnetware.com)
' Module    : frmMain
' Created   : 07/12/2002
' Author    : GWilmot
' Purpose   : Main form of application, displays Network information
'-----------------------------------------------------------------------
' Dependancies : ICE_NetFunctions
' Assumptions  : Running on Win NT/2000/XP
' Last Updated :
'-----------------------------------------------------------------------

' Interface Details
' =================

' Key Properties
' --------------
' DebugMode (Write) Sets true to set in Debugmode

' Approach : Load/Set DebugMode/Show - expects to exit from Application from this object

' Standard Properties
Private nbDebugMode As Boolean

' Dependancies
Private noNetFunctions As New ICE_NetFunctions  ' Network Routines

' Resize Variables
Private nlMinHeight As Long    ' Holds the minimum form Height
Private nlMinWidth As Long     ' Holds the minimum form Width
Private nlRightMargin As Long  ' Holds the right Margin offset
Private nlTreeHeight As Long   ' Holds the treeview margin offset

Private Sub Form_Load()

On Error GoTo Catch

' Updates the Network Tree
RefreshTree

' Set the form caption with version
Me.Caption = Me.Caption & " " & Format$(App.Major, "00") & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "000")

' Centre the form first
Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

' Set the Resize parameters (the defaults to current size - but could be made smaller)
nlMinHeight = Me.Height
nlMinWidth = Me.Width
' Get the offsets
nlTreeHeight = Me.Height - (trvNetwork.Top + trvNetwork.Height)
nlRightMargin = Me.Width - (txtDisplay.Left + txtDisplay.Width)

' Set the default information display
DisplayNetworkInformation

Finally:
    ' Clean-up
    Exit Sub

Catch:

    MsgBox "Fatal error occurred during application load!" & vbCrLf & Err.Description, vbCritical
    End

End Sub

Private Sub Form_Resize()
' Just re-sizes form
Dim lMeHeight As Long     ' Holds the Form height to base the resize on
Dim lMeWidth As Long      ' Holds the Form width to base the resize on

'*** Could put in splitter bar between tree & text box
'*** Could save in registry last settings

On Error GoTo Catch

Select Case Me.WindowState
    Case vbMinimized
    Case Else
        
        ' See if we're at minimums
        If Me.Height > nlMinHeight Then lMeHeight = Me.Height Else lMeHeight = nlMinHeight
        If Me.Width > nlMinWidth Then lMeWidth = Me.Width Else lMeWidth = nlMinWidth
    
        ' Form specific Code
        
        ' Right Margins
        txtDisplay.Width = lMeWidth - (txtDisplay.Left + nlRightMargin)
        
        ' TreeView height
        trvNetwork.Height = lMeHeight - (trvNetwork.Top + nlTreeHeight)
        txtDisplay.Height = lMeHeight - (txtDisplay.Top + nlTreeHeight)
        
End Select

Finally:
    ' Clean-up
    Exit Sub

Catch:

    ' Ignore error unless in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form frmMain"
    Resume Finally

End Sub

Private Sub Form_Unload(Cancel As Integer)
' As this is a simple application - Program exited from the Form UnLoad

On Error GoTo Catch

'  Release the objects
Set noNetFunctions = Nothing

Finally:
    End
    Exit Sub
Catch:
    MsgBox "Error occurred during application unload!" & vbCrLf & Err.Description, vbExclamation
    End

End Sub

Private Sub mnuContents_Click()
' Simple help that will load a text file

On Error GoTo Catch

' See if the file is there
If Dir$(App.Path & "\readme.txt") <> vbNullString Then
    ' just a simple shell load
    Shell "Notepad " & App.Path & "\help.txt", vbNormalFocus
Else
    MsgBox "No Help File!", vbExclamation
End If

Finally:
    Exit Sub
Catch:
    MsgBox "Problem Displaying Help!" + vbCrLf + Err.Description, vbExclamation
    Resume Finally

End Sub

Private Sub mnuExit_Click()
' Make it go through its unload
Form_Unload False
End Sub

Private Function ShellSortStrings(ByRef StringData() As String) As Boolean
'-----------------------------------------------------------------------
' Procedure    : frmMain.ShellSortStrings
' Author       : GWilmot
' Date Created : 05/12/2002
'-----------------------------------------------------------------------
' Purpose      : Sorts a single dimension array of strings in ascending
'                Order, case insensitive
' Assumptions  :
' Inputs       : String Array
' Returns      : True is succeeded & sorted Array, False if not
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim lsLocalData() As String    ' Holds a local copy of the data
Dim i As Long                  ' For Counter
Dim lbSwapped As Boolean       ' Flag to show that a value has been swapped
Dim lsBuffer As String         ' Holding point

On Error GoTo Catch

' Set up our local copy (that way if we error - we leave it how we found it)
ReDim lsLocalData(LBound(StringData) To UBound(StringData))
For i = LBound(StringData) To UBound(StringData)
    lsLocalData(i) = StringData(i)
Next i

Do
    ' Default
    lbSwapped = False
    For i = LBound(lsLocalData) To UBound(lsLocalData) - 1
        If UCase$(lsLocalData(i)) > UCase$(lsLocalData(i + 1)) Then
            ' Watch the shells....
            lsBuffer = lsLocalData(i)
            lsLocalData(i) = lsLocalData(i + 1)
            lsLocalData(i + 1) = lsBuffer
            lbSwapped = True
        End If
    Next i
Loop Until Not lbSwapped

' Copy back
For i = LBound(StringData) To UBound(StringData)
    StringData(i) = lsLocalData(i)
Next i

' Set the success flag
ShellSortStrings = True

Finally:
    ' Clean-up
    Exit Function

Catch:
    ' Ignore error unless in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShellSortStrings of Form frmMain"
    Resume Finally

End Function

Private Sub mnuHelpAbout_Click()
' Just loads the about form

On Error GoTo Catch

' Load & show
Load frmAbout
frmAbout.Show vbModal

Finally:
    ' Clean-up
    Set frmAbout = Nothing
    Exit Sub

Catch:
    ' Ignore error unless in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuHelpAbout_Click of Form frmMain"
    Resume Finally
End Sub

Private Sub mnuRefreshNetwork_Click()
' Just refreshes the network tree
RefreshTree
End Sub

Private Sub trvNetwork_NodeClick(ByVal Node As MSComctlLib.Node)
' Updates the txtDisplay according to the type of object currently selected

On Error GoTo Catch

' set defaults
txtDisplay.Text = vbNullString
Me.MousePointer = vbHourglass

If trvNetwork.SelectedItem.Key = "Network" Then
    DisplayNetworkInformation
Else
    Select Case trvNetwork.SelectedItem.Parent.Text
        Case "Network Root"
            DisplayNetworkInformation
        Case "Computers"
            DisplayComputerInformation trvNetwork.SelectedItem.Text
        Case "Users"
            DisplayUserInformation trvNetwork.SelectedItem.Text
        Case "Groups"
            DisplayGroupInformation trvNetwork.SelectedItem.Text
        Case "Local Groups"
            DisplayLocalGroupInformation trvNetwork.SelectedItem.Text
    End Select
End If

Finally:
    ' Clean-up
    Me.MousePointer = vbDefault
    Exit Sub

Catch:
    ' Ignore error unless in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure trvNetwork_NodeClick of Form frmMain"
    Resume Finally

End Sub

Private Sub DisplayNetworkInformation()
'-----------------------------------------------------------------------
' Procedure    : frmMain.DisplayNetworkInformation
' Author       : GWilmot
' Date Created : 05/12/2002
'-----------------------------------------------------------------------
' Purpose      : Just Displays the Network Information
' Assumptions  : Displays its own errors
' Inputs       : none
' Returns      : none
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim lsDisplay As String          ' This builds up the text to display

On Error GoTo Catch

lsDisplay = "NETWORK" & vbCrLf & vbCrLf & "Primary Domain Controller : " & noNetFunctions.GetPrimaryDomainController

txtDisplay.Text = lsDisplay

Finally:
    ' Clean-up
    Exit Sub

Catch:
    ' If in debugmode - display developer error otherwise display user error
    If nbDebugMode Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DisplayNetworkInformation of Form frmMain"
    Else
        MsgBox "Problem Displaying Network Information", vbExclamation
    End If
    Resume Finally

End Sub

Private Sub DisplayComputerInformation(ByVal Computer As String)
'-----------------------------------------------------------------------
' Procedure    : frmMain.DisplayComputerInformation
' Author       : GWilmot
' Date Created : 05/12/2002
'-----------------------------------------------------------------------
' Purpose      : Displays network information associated with a given
'                Computer name in the txtDisplay object by building a
'                Text string
' Assumptions  : Displays its own errors
' Inputs       : Computer name to display information for
' Returns      : none
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim lsDisplay As String          ' This builds up the text to display
Dim i As Long                    ' For counter
Dim lsItems() As String          ' Holds returned data

On Error GoTo Catch

' Display Header
lsDisplay = "COMPUTER : " & Computer & vbCrLf & vbCrLf

' Get Users
If noNetFunctions.GetWorkStationUsers(Computer, lsItems) Then
    ShellSortStrings lsItems
    lsDisplay = lsDisplay & "USERS : " & vbCrLf & vbCrLf
    For i = LBound(lsItems) To UBound(lsItems)
        lsDisplay = lsDisplay & lsItems(i) & vbCrLf
    Next i
    lsDisplay = lsDisplay & vbCrLf
Else
    lsDisplay = lsDisplay & "Failed to retrieve Users" & vbCrLf & vbCrLf
End If

' Get User Sessions
If noNetFunctions.GetUserSessions(Computer, lsItems) Then
    ShellSortStrings lsItems
    lsDisplay = lsDisplay & "SESSIONS : " & vbCrLf & vbCrLf
    For i = LBound(lsItems) To UBound(lsItems)
        lsDisplay = lsDisplay & lsItems(i) & vbCrLf
    Next i
    lsDisplay = lsDisplay & vbCrLf
Else
    lsDisplay = lsDisplay & "Failed to retrieve Sessions" & vbCrLf & vbCrLf
End If

' Get Shares
If noNetFunctions.GetWorkStationShares(Computer, lsItems) Then
    ShellSortStrings lsItems
    lsDisplay = lsDisplay & "SHARES : " & vbCrLf & vbCrLf
    For i = LBound(lsItems) To UBound(lsItems)
        lsDisplay = lsDisplay & lsItems(i) & vbCrLf
    Next i
    lsDisplay = lsDisplay & vbCrLf
Else
    lsDisplay = lsDisplay & "Failed to retrieve Shares" & vbCrLf & vbCrLf
End If

' Showtime
txtDisplay.Text = lsDisplay

Finally:
    ' Clean-up
    Exit Sub

Catch:
    ' If in debugmode - display developer error otherwise display user error
    If nbDebugMode Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DisplayComputerInformation of Form frmMain"
    Else
        MsgBox "Problem Displaying Computer Information", vbExclamation
    End If
    Resume Finally

End Sub

Private Sub DisplayUserInformation(ByVal User As String)
'-----------------------------------------------------------------------
' Procedure    : frmMain.DisplayUserInformation
' Author       : GWilmot
' Date Created : 05/12/2002
'-----------------------------------------------------------------------
' Purpose      : Displays network information associated with a given
'                User name in the txtDisplay object by building a
'                Text string
' Assumptions  : Displays its own errors
' Inputs       : User name to display information for
' Returns      : none
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim lsDisplay As String          ' This builds up the text to display
Dim lsUserData As String         ' Holds the basic user data

On Error GoTo Catch

lsDisplay = "USER : " & User & vbCrLf & vbCrLf

' Get the user data
noNetFunctions.GetNetUserInfo noNetFunctions.GetPrimaryDomainController, User, lsUserData
lsDisplay = lsDisplay & lsUserData

' Showtime
txtDisplay.Text = lsDisplay

Finally:
    ' Clean-up
    Exit Sub

Catch:
    ' If in debugmode - display developer error otherwise display user error
    If nbDebugMode Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DisplayUserInformation of Form frmMain"
    Else
        MsgBox "Problem Displaying User Information", vbExclamation
    End If
    Resume Finally

End Sub

Private Sub DisplayGroupInformation(ByVal Group As String)
'-----------------------------------------------------------------------
' Procedure    : frmMain.DisplayGroupInformation
' Author       : GWilmot
' Date Created : 05/12/2002
'-----------------------------------------------------------------------
' Purpose      : Displays network information associated with a given
'                Group name in the txtDisplay object by building a
'                Text string
' Assumptions  : Displays its own errors
' Inputs       : User name to display information for
' Returns      : None
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim lsDisplay As String          ' This builds up the text to display
Dim i As Long                    ' For counter
Dim lsItems() As String          ' Holds returned data
On Error GoTo Catch

lsDisplay = "GROUP : " & Group & vbCrLf & vbCrLf

' Get Group Users
If noNetFunctions.GetGroupUsers(noNetFunctions.GetPrimaryDomainController, Group, lsItems) Then
    ShellSortStrings lsItems
    lsDisplay = lsDisplay & "USERS : " & vbCrLf & vbCrLf
    For i = LBound(lsItems) To UBound(lsItems)
        lsDisplay = lsDisplay & lsItems(i) & vbCrLf
    Next i
    lsDisplay = lsDisplay & vbCrLf
End If

' Showtime
txtDisplay.Text = lsDisplay

Finally:
    ' Clean-up
    Exit Sub

Catch:
    ' If in debugmode - display developer error otherwise display user error
    If nbDebugMode Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DisplayGroupInformation of Form frmMain"
    Else
        MsgBox "Problem Displaying Group Information", vbExclamation
    End If
    Resume Finally

End Sub

Private Sub DisplayLocalGroupInformation(ByVal LocalGroup As String)
'-----------------------------------------------------------------------
' Procedure    : frmMain.DisplayLocalGroupInformation
' Author       : GWilmot
' Date Created : 05/12/2002
'-----------------------------------------------------------------------
' Purpose      : Displays network information associated with a given
'                Local Group name in the txtDisplay object by building a
'                Text string
' Assumptions  : Displays its own errors
' Inputs       : User name to display information for
' Returns      : None
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim lsDisplay As String          ' This builds up the text to display
Dim i As Long                    ' For counter
Dim lsItems() As String          ' Holds returned data

On Error GoTo Catch

lsDisplay = "LOCAL GROUP : " & LocalGroup & vbCrLf & vbCrLf

' Get Group Users
If noNetFunctions.GetLocalGroupUsers(noNetFunctions.GetPrimaryDomainController, LocalGroup, lsItems) Then
    ShellSortStrings lsItems
    lsDisplay = lsDisplay & "USERS : " & vbCrLf & vbCrLf
    For i = LBound(lsItems) To UBound(lsItems)
        lsDisplay = lsDisplay & lsItems(i) & vbCrLf
    Next i
    lsDisplay = lsDisplay & vbCrLf
End If

' Showtime
txtDisplay.Text = lsDisplay

Finally:
    ' Clean-up
    Exit Sub

Catch:
    ' If in debugmode - display developer error otherwise display user error
    If nbDebugMode Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DisplayLocalGroupInformation of Form frmMain"
    Else
        MsgBox "Problem Displaying Local Group Information", vbExclamation
    End If
    Resume Finally

End Sub

Private Sub RefreshTree()
'-----------------------------------------------------------------------
' Procedure    : noNetFunctions.RefreshTree
' Author       : GWilmot
' Date Created : 06/12/2002
'-----------------------------------------------------------------------
' Purpose      : To refresh the Tree View that display Network Entities
' Assumptions  : Dependant objects have been instantiated
' Inputs       : None
' Returns      : None
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim lsItems() As String         ' Holds the list of SQl-Users
Dim i As Long                   ' For Counter
Dim loNode As Node              ' Holds a tree node

On Error GoTo Catch

' Server types constants
Const SV_TYPE_ALL                 As Long = &HFFFFFFFF

' Tell the user to hang on if visible
If Me.Visible Then Me.MousePointer = vbHourglass

' Clears done any nodes
trvNetwork.Nodes.Clear

' Set Tree defaults
trvNetwork.Nodes.Add , , "Network", "Network Root", 5
trvNetwork.Nodes.Add "Network", tvwChild, "Computers", "Computers", 6
trvNetwork.Nodes.Add "Network", tvwChild, "Users", "Users", 3
trvNetwork.Nodes.Add "Network", tvwChild, "Groups", "Groups", 2
trvNetwork.Nodes.Add "Network", tvwChild, "Local Groups", "Local Groups", 2

Set loNode = trvNetwork.Nodes.Item(1)
loNode.Expanded = True

' OK Lets populate the tree with workstations
If noNetFunctions.GetServers(vbNullString, lsItems, SV_TYPE_ALL) Then
    ShellSortStrings lsItems
    For i = 1 To UBound(lsItems)
        trvNetwork.Nodes.Add "Computers", tvwChild, lsItems(i), lsItems(i), 1
    Next i
End If

' Gets users
If noNetFunctions.GetUsers(noNetFunctions.GetPrimaryDomainController, lsItems) Then
    ShellSortStrings lsItems
    For i = 1 To UBound(lsItems)
        trvNetwork.Nodes.Add "Users", tvwChild, lsItems(i), lsItems(i), 3
    Next i
End If

' Gets groups
If noNetFunctions.GetGroups(noNetFunctions.GetPrimaryDomainController, lsItems) Then
    ShellSortStrings lsItems
    For i = 1 To UBound(lsItems)
        trvNetwork.Nodes.Add "Groups", tvwChild, lsItems(i), lsItems(i), 2
    Next i
End If

' Gets local groups
If noNetFunctions.GetLocalGroups(noNetFunctions.GetPrimaryDomainController, lsItems) Then
    ShellSortStrings lsItems
    For i = 1 To UBound(lsItems)
        trvNetwork.Nodes.Add "Local Groups", tvwChild, "Local Groups" & lsItems(i), lsItems(i), 2
    Next i
End If

Finally:
    ' Clean-up
    If Me.Visible Then Me.MousePointer = vbDefault
    Exit Sub

Catch:
    ' If in debugmode - display developer error otherwise display user error
    If nbDebugMode Then
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RefreshTree of Form frmMain"
    Else
        If Me.Visible Then MsgBox "Problem Displaying Network Entities", vbExclamation
    End If
    Resume Finally

End Sub

Public Property Let DebugMode(ByVal Flag As Boolean)
' Simple Flag to say if object in debug mode
nbDebugMode = Flag
End Property

