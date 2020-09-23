VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4140
   ClientLeft      =   4080
   ClientTop       =   3270
   ClientWidth     =   5595
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSystemInfo 
      Caption         =   "&System Info..."
      Height          =   375
      Left            =   4140
      TabIndex        =   1
      Top             =   3510
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4140
      TabIndex        =   0
      Top             =   2970
      Width           =   1335
   End
   Begin VB.Label lblHyperlink 
      Caption         =   "www.ICEnetware.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1620
      TabIndex        =   7
      Top             =   1395
      Width           =   1695
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2670
      Left            =   90
      Picture         =   "Frmabout.frx":0000
      Stretch         =   -1  'True
      Top             =   135
      Width           =   1440
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "netapi32.dll"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   810
      Index           =   1
      Left            =   1620
      TabIndex        =   5
      Top             =   180
      Width           =   3480
   End
   Begin VB.Label lblUser 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   1620
      TabIndex        =   4
      Top             =   1935
      Width           =   3885
   End
   Begin VB.Label lblProduct 
      Caption         =   "This Product is licensed to:"
      Height          =   180
      Left            =   1620
      TabIndex        =   3
      Top             =   1710
      Width           =   3900
   End
   Begin VB.Label lblInformation 
      Height          =   315
      Left            =   1620
      TabIndex        =   2
      Top             =   1035
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   45
      X2              =   5445
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "netapi32.dll"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   810
      Index           =   0
      Left            =   1575
      TabIndex        =   6
      Top             =   180
      Width           =   3480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------
' Copyright : ICEnetware Ltd 2003 (www.ICEnetware.com)
' Form      : frmAbout
' Created   : 06/12/2002
' Author    : GWilmot
' Purpose   : Generic About form
'-----------------------------------------------------------------------
' Dependancies : None
' Assumptions  :
' Last Updated :
'-----------------------------------------------------------------------
Private nsSystemInfoPath As String        ' Holds the path name to run MSInfo
Private nlSystemInfoHandle As Long        ' Holds the handle of MSInfo App

Private Const HKEY_LOCAL_MACHINE As Long = &H80000002

' Used to get registery information
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwReserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName$, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdSystemInfo_Click()

On Error GoTo Catch

' This is the label for the start of the routine & is used by the error routine
ProcStart:

' Have we already loaded an instance?
If nlSystemInfoHandle = 0 Then
    ' No - create a new instance
    nlSystemInfoHandle = Shell(nsSystemInfoPath, vbNormalFocus)
Else
    ' Yes - Attempt to set the focus using the previously returned ID
    AppActivate nlSystemInfoHandle
End If

Finally:
    Exit Sub

Catch:
    If nlSystemInfoHandle = 0 Then
        ' If we try and run it for the FIRST time and an error occurs & inform the user & disable the button
        MsgBox "Problem in running MSINFO", vbExclamation
        cmdSystemInfo.Enabled = False
        Resume Finally
    Else
        ' We have already run it but the instance has probably been closed (damn users!)
        ' so lets start it as the first time & try again
        nlSystemInfoHandle = 0
        Resume ProcStart
    End If
        
End Sub


Private Sub Form_Load()

Dim lsName As String                ' User Name
Dim lsCompany As String             ' Company Name

On Error GoTo Catch

' Centre form
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

' Sort out the form Title according to the application name
Caption = "About " & App.Title

lblInformation = "(C) ICEnetware Ltd 2003"
    
' Get settings from the Registry
nsSystemInfoPath = nfsGetRegistryString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Shared Tools\MSInfo", "Path", "")
lsName = nfsGetRegistryString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner", "")
If lsName = vbNullString Then lsName = nfsGetRegistryString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner", "")
lsCompany = nfsGetRegistryString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOrganization", "")
If lsCompany = vbNullString Then lsCompany = nfsGetRegistryString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization", "")
    
' Sort out if we can enable the system information flag
If nsSystemInfoPath = vbNullString Then nsSystemInfoPath = nfsGetRegistryString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Shared Tools Location", "MSINFO", "")
If nsSystemInfoPath = vbNullString Then cmdSystemInfo.Enabled = False Else cmdSystemInfo.Enabled = True

' Sort out the labels
lblUser = lsName & vbCrLf & lsCompany

Exit Sub

Catch:
    MsgBox "Can't Display About Form", vbExclamation
    Unload Me
End Sub

Private Function nfsGetRegistryString(ByVal vlHinKey As Long, ByVal vsSubkey As String, ByVal vsValname As String, ByVal vsDefault As String) As String
'-----------------------------------------------------------------------
' Procedure    : frmSplash.nfsGetRegistryString
' Author       : GWilmot
' Date Created : 25/11/2002
'-----------------------------------------------------------------------
' Purpose      : Returns a value from the registry
' Assumptions  :
' Inputs       :
' Returns      :
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim lsValue As String * 512           ' Buffer to receive information
Dim lsReturnVal As String             ' Return value
Dim llhSubKey As Long                 ' Handle to the Registry
Dim lldwType As Long                  ' Type of data Returned
Dim llReply As Long                   ' Return from API call
Dim i As Integer                      ' Counter
Dim buffer As Long                    ' Setting the buffer to return

Const KEY_ALL_ACCESS As Long = 131135
Const ERROR_SUCCESS As Long = 0
Const REG_SZ As Long = 1

On Error GoTo Catch

buffer = 512

' Open the key
llReply = RegOpenKeyEx(vlHinKey, vsSubkey, 0, KEY_ALL_ACCESS, llhSubKey)

' Get the data
llReply = RegQueryValueEx(llhSubKey, vsValname, 0, lldwType, ByVal lsValue, buffer)

' See if we have anything returned and its datatype
If llReply = 0 And lldwType = REG_SZ Then
    i = InStr(lsValue, vbNullChar)
    
    If i > 0 Then
        lsReturnVal = Trim$(Left$(lsValue, i - 1))
    Else
        lsReturnVal = Trim$(Left$(lsValue, buffer))
    End If

    ' If nothing return the default
    If lsReturnVal = vbNullString Then lsReturnVal = vsDefault
Else
    ' If the wrong type return the default
    lsReturnVal = vsDefault
End If

' Tidy up
If vlHinKey = 0 Then i = RegCloseKey(llhSubKey)

' Set the data
nfsGetRegistryString = lsReturnVal

Finally:
    ' Clean-up
    Exit Function

Catch:
    ' Just set to the default
    nfsGetRegistryString = vsDefault
    Resume Finally

End Function
Private Sub lblHyperlink_Click()
'-----------------------------------------------------------------------
' Procedure    : frmAbout.lblHyperlink_Click
' Author       : GWilmot
' Date Created : 25/11/2002
'-----------------------------------------------------------------------
' Purpose      : Loads a linked page
' Assumptions  :
' Inputs       :
' Returns      :
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim loBrowser As Object        ' This will hold the browser object
                               ' Late bound DELIBRATELY to cope with different version's of IE

On Error GoTo Catch

Me.MousePointer = vbHourglass

' Set the colour of the foreground
lblHyperlink.ForeColor = &HFF00FF

' OK create the browser object
Set loBrowser = GetObject(vbNullString, "internetexplorer.application")
 
' Now set the address to our site
loBrowser.navigate "http://www.ICEnetware.com"

' Make visible
loBrowser.Visible = True

Set loBrowser = Nothing

Finally:
    ' Clean-up
    Me.MousePointer = vbDefault
    Exit Sub

Catch:
    ' ignore error
    Resume Finally

End Sub


