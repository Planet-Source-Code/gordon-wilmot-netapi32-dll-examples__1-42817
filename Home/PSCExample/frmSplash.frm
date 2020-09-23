VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3690
   ClientLeft      =   1440
   ClientTop       =   3150
   ClientWidth     =   7005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblInformation 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1650
      TabIndex        =   5
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label lblProduct 
      BackColor       =   &H00FFFFFF&
      Caption         =   "This Product is licensed to:"
      Height          =   180
      Left            =   1650
      TabIndex        =   4
      Top             =   2565
      Width           =   3900
   End
   Begin VB.Label lblUser 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   1605
      TabIndex        =   3
      Top             =   2790
      Width           =   5370
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1605
      TabIndex        =   2
      Top             =   1800
      Width           =   5385
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "netapi32.dll"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Index           =   0
      Left            =   1635
      TabIndex        =   1
      Top             =   1080
      Width           =   3180
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "netapi32.dll"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Index           =   1
      Left            =   1665
      TabIndex        =   0
      Top             =   1080
      Width           =   3180
   End
   Begin VB.Image imgLogo 
      Height          =   1845
      Index           =   1
      Left            =   -270
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   -675
      Width           =   7485
   End
   Begin VB.Image imgLogo 
      Height          =   3780
      Index           =   0
      Left            =   -630
      Picture         =   "frmSplash.frx":2BE4E
      Stretch         =   -1  'True
      Top             =   1125
      Width           =   2220
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------
' Copyright : ICEnetware Ltd 2003 (www.ICEnetware.com)
' Module    : frmSplash
' Created   : 07/12/2002
' Author    : GWilmot
' Purpose   : Main form of application, displays Network information
'-----------------------------------------------------------------------
' Dependancies :
' Assumptions  :
' Last Updated :
'-----------------------------------------------------------------------
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002

' Used to get registery information
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwReserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName$, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

' Used to set Top Most
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Load()
Dim lsName As String                ' User Name
Dim lsCompany As String             ' Company Name

On Error GoTo Catch

' Centre
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

' Set standard labels
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
lblInformation = "(C) ICEnetware Ltd 2003"
    
' Get settings from the Registry
lsName = nfsGetRegistryString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner", "")
lsCompany = nfsGetRegistryString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOrganization", "")
If lsName = vbNullString Then lsName = nfsGetRegistryString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner", "")
If lsCompany = vbNullString Then lsCompany = nfsGetRegistryString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization", "")
    
' Sort out the labels
lblUser = " " & lsName & vbCrLf & "  " & lsCompany
Finally:

Exit Sub

Catch:
    Resume Finally
    
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


Public Sub SetTopMost(ByVal Topmost As Boolean)
'-----------------------------------------------------------------------
' Procedure    : frmSplash.SetTopMost
' Author       : GWilmot
' Date Created : 25/11/2002
'-----------------------------------------------------------------------
' Purpose      : Forces the form to be topmost
' Assumptions  :
' Inputs       :
' Returns      :
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim llReply As Long       ' Holds API reply

On Error GoTo Catch

' Make the window topmost
If Topmost = True Then
   llReply = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, _
      0, flags)
Else
   llReply = SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, _
      0, 0, flags)
End If
         
Finally:
    ' Clean-up
    Exit Sub

Catch:
    ' Ignore error
    Resume Finally

End Sub



