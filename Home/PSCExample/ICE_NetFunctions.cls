VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICE_NetFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------
' Copyright : ICEnetware Ltd 2003 (www.ICEnetware.com)
' Module    : ICE_NetFunctions
' Created   : 06/12/2002
' Author    : GWilmot
' Purpose   : Contains Generic Network Methods
'-----------------------------------------------------------------------
' Dependancies : None Outside API Usage
' Assumptions  : Running on NT4/2000/XP OS's
' Last Updated :
'-----------------------------------------------------------------------
' Interface Details
' =================

' Key Properties
' --------------
' DebugMode : Sets the object into Developer/Support DebugMode

' Key Public Methods
' ------------------
' GetGroups : Gets a List of the Groups for a given domain
' GetGroupUsers : Gets a list of users for a given group
' GetLocalGroups : Returns a list of local groups for a given domain
' GetLocalGroupUsers : Gets a list of users for a given local group
' GetNetUserInfo : Returns User Information from a domain
' GetPrimaryDomainController : Returns the PDC for the current workstation
' GetServers : Gets a List of Servers based on a Domain
' GetUsers : Returns a list of users for a given Domain
' GetUserSessions : Gets the list of users/workstations for a given domain
' GetWorkStationShares : Returns the shares for a given workstation including hidden shares
' GetWorkStationUsers : Gets all the users for a given workstation

' Private Details
' ===============

' Netapi32.dll
Private Declare Function NetApiBufferFree Lib "netapi32.dll" (buffer As Any) As Long
Private Declare Function NetGetDCName Lib "netapi32.dll" (strServerName As Any, strDomainName As Any, pBuffer As Long) As Long
Private Declare Function NetUserGetInfo Lib "netapi32" (lpServer As Any, username As Byte, ByVal level As Long, lpBuffer As Long) As Long
Private Declare Function NetGroupEnum Lib "netapi32" (ByVal servername As Long, ByVal level As Long, buf As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, resume_handle As Long) As Long
Private Declare Function NetSessionEnum Lib "netapi32" (ByVal servername As Long, ByVal ClientName As Long, ByVal User As Long, ByVal level As Long, buf As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, resume_handle As Long) As Long
Private Declare Function NetGroupGetUsers Lib "netapi32" (ByVal servername As Long, ByVal groupname As Long, ByVal level As Long, buf As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, resume_handle As Long) As Long
Private Declare Function NetLocalGroupEnum Lib "netapi32" (ByVal servername As Long, ByVal level As Long, buf As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, resume_handle As Long) As Long
Private Declare Function NetLocalGroupGetMembers Lib "netapi32" (ByVal servername As Long, ByVal groupname As Long, ByVal level As Long, buf As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, resume_handle As Long) As Long
Private Declare Function NetServerEnum Lib "netapi32" (ByVal servername As Long, ByVal level As Long, buf As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, ByVal ServerType As Long, ByVal WorkStation As Long, resume_handle As Long) As Long
Private Declare Function NetShareEnum Lib "netapi32" (ByVal servername As Long, ByVal level As Long, buf As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, resume_handle As Long) As Long
Private Declare Function NetUserEnum Lib "netapi32" (ByVal servername As Long, ByVal level As Long, ByVal filter As Long, buf As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, resume_handle As Long) As Long
Private Declare Function NetWkstaUserEnum Lib "netapi32" (ByVal servername As Long, ByVal level As Long, buf As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, resume_handle As Long) As Long

' Kernel32 API Declares
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

' This is the simple structure
Private Type SERVER_INFO_100
    Platform As Long
    name     As Long
End Type

Private Type NAME_INFO
    name As Long
End Type

Private Type LOCAL_GROUP
    dummy1 As Long
    dummy2 As Long
    name As Long
End Type

' User information
Private Type USER_INFO_3_API
   ' Level 0 starts here
   name As Long
   ' Level 1 starts here
   Password As Long
   PasswordAge As Long
   Privilege As Long
   HomeDir As Long
   Comment As Long
   flags As Long
   ScriptPath As Long
   ' Level 2 starts here
   AuthFlags As Long
   FullName As Long
   UserComment As Long
   Parms As Long
   Workstations As Long
   LastLogon As Long
   LastLogoff As Long
   AcctExpires As Long
   MaxStorage As Long
   UnitsPerWeek As Long
   LogonHours As Long
   BadPwCount As Long
   NumLogons As Long
   LogonServer As Long
   CountryCode As Long
   CodePage As Long
   ' Level 3 starts here
   UserID As Long
   PrimaryGroupID As Long
   Profile As Long
   HomeDirDrive As Long
   PasswordExpired As Long
End Type

Private Type SESSION_INFO_10
    name As Long
    username As Long
    uptime As Long
    idle_time As Long
End Type

' Variables
' ---------
Private nbDebugMode As Boolean     ' Internal debug flag

Public Property Let DebugMode(ByVal Flag As Boolean)
' Simple Flag to say if object in debug mode
nbDebugMode = Flag
End Property

Public Function GetServers(ByVal Domain As String, ByRef Servers() As String, _
    ByVal ServerType As Long) As Boolean
'-----------------------------------------------------------------------
' Procedure    : ICE_NetFunctions.GetServers
' Author       : GWilmot
' Date Created : 07/12/2002
'-----------------------------------------------------------------------
' Purpose      : Gets a List of Servers based on a Domain
' Assumptions  :
' Inputs       : Domain - normally the PDC
'                Servers() - List of servers returned
'                ServerType - Type of server to search for (see list below)
' Returns      : False if Failed
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim llReply As Long                     ' Holds API Reply
Dim llEntriesRead As Long               ' Total number of entries returned by API call
Dim llTotalEntries As Long              ' Total number of entries to be recalled
Dim lhResume As Long                    ' Handle to resume the call if llEntriesRead<>llTotalEntries
Dim llBuffer As Long                    ' Structure Pointer
Dim llLevel As Long                     ' Structure type to retrieve
Dim llPrefMaxLen As Long                ' Max amount of data to return
Dim ltServerInfo As SERVER_INFO_100     ' Holds data returned
Dim i As Long                           ' For Counter
Dim llServerCount As Long               ' Counts the found servers
    
On Error GoTo Catch

' Types of server that can be returned
' They can be concatenated
Const SV_TYPE_WORKSTATION         As Long = &H1
Const SV_TYPE_SERVER              As Long = &H2
Const SV_TYPE_SQLSERVER           As Long = &H4
Const SV_TYPE_DOMAIN_CTRL         As Long = &H8
Const SV_TYPE_DOMAIN_BAKCTRL      As Long = &H10
Const SV_TYPE_TIME_SOURCE         As Long = &H20
Const SV_TYPE_AFP                 As Long = &H40
Const SV_TYPE_NOVELL              As Long = &H80
Const SV_TYPE_DOMAIN_MEMBER       As Long = &H100
Const SV_TYPE_PRINTQ_SERVER       As Long = &H200
Const SV_TYPE_DIALIN_SERVER       As Long = &H400
Const SV_TYPE_XENIX_SERVER        As Long = &H800
Const SV_TYPE_SERVER_UNIX         As Long = SV_TYPE_XENIX_SERVER
Const SV_TYPE_NT                  As Long = &H1000
Const SV_TYPE_WFW                 As Long = &H2000
Const SV_TYPE_SERVER_MFPN         As Long = &H4000
Const SV_TYPE_SERVER_NT           As Long = &H8000
Const SV_TYPE_POTENTIAL_BROWSER   As Long = &H10000
Const SV_TYPE_BACKUP_BROWSER      As Long = &H20000
Const SV_TYPE_MASTER_BROWSER      As Long = &H40000
Const SV_TYPE_DOMAIN_MASTER       As Long = &H80000
Const SV_TYPE_SERVER_OSF          As Long = &H100000
Const SV_TYPE_SERVER_VMS          As Long = &H200000
Const SV_TYPE_WINDOWS             As Long = &H400000  'Windows95 and above
Const SV_TYPE_DFS                 As Long = &H800000  'Root of a DFS tree
Const SV_TYPE_CLUSTER_NT          As Long = &H1000000 'NT Cluster
Const SV_TYPE_TERMINALSERVER      As Long = &H2000000 'Terminal Server
Const SV_TYPE_DCE                 As Long = &H10000000 'IBM DSS
Const SV_TYPE_ALTERNATE_XPORT     As Long = &H20000000 'rtn alternate transport
Const SV_TYPE_LOCAL_LIST_ONLY     As Long = &H40000000 'rtn local only
Const SV_TYPE_DOMAIN_ENUM         As Long = &H80000000
Const SV_TYPE_ALL                 As Long = &HFFFFFFFF

' API return codes
Const SV_SUCCESS As Long = 0
Const SV_MORE_DATA As Long = 234&

Const MAX_PREFERRED_LENGTH As Long = -1

' Defaults
llLevel = 100                       ' Basic Structure
llPrefMaxLen = MAX_PREFERRED_LENGTH ' No Limit on whot can be returned
ReDim Servers(0 To 0)               ' if Ubound(Servers) = 0 then its empty
    
' As a note a null domain name returns the default
    
Do
    ' Call depends on whether we have a domain name
    If Domain = vbNullString Then
        llReply = NetServerEnum(0&, llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, ServerType, 0&, lhResume)
    Else
        llReply = NetServerEnum(0&, llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, ServerType, StrPtr(Domain), lhResume)
    End If
            
    ' See if we succeeded
    If llReply = SV_SUCCESS Or llReply = SV_MORE_DATA Then
        
        ' Iterate through all the returned servers
        For i = 0 To llEntriesRead - 1
            
            ' Copy the data into the structure
            RtlMoveMemory ltServerInfo, ByVal llBuffer, Len(ltServerInfo)
            
            ' Increment our server count & redim array
            llServerCount = llServerCount + 1
            If llServerCount = 1 Then
                ReDim Servers(1 To 1)
            Else
                ReDim Preserve Servers(1 To llServerCount)
            End If
            
            ' OK get it to the array by taking the structure pointer
            ' and getting the string of the name
            Servers(llServerCount) = WinPointerToString(ltServerInfo.name)
            
            ' Adjust the pointer to the next entry
            llBuffer = llBuffer + Len(ltServerInfo)
        Next i
    End If
    
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer

' As we've set the maximum llPrefMaxLen = -1 then it SHOULD only require a single call plus
' a bug has been associated with the call & its resume handle (Can use undocumented call
' NetServerEnumEX if this is a problem)
Loop Until llReply <> SV_MORE_DATA

' Only return true if we got any
If UBound(Servers) > 0 Then GetServers = True
    
ProcResume:
    Exit Function

Finally:
    ' Pretty basic error handling
    ' we could catch the errors for instance
    ' Const SV_ACCESS_DENIED = 5            ' Access Denied
    ' Const SV_SERVER_NOT_ACCESSABLE = 1722 ' Server not available
    
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer
    Exit Function

Catch:
    ' Reports back to developer if in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetServers of Class Module ICE_NetFunctions"
    Resume Finally

End Function

Public Function GetWorkStationUsers(ByVal WorkStation As String, ByRef Users() As String) As Boolean
'-----------------------------------------------------------------------
' Procedure    : ICE_NetFunctions.GetWorkStationUsers
' Author       : GWilmot
' Date Created : 07/12/2002
'-----------------------------------------------------------------------
' Purpose      : Gets all the users for a given workstation
' Assumptions  :
' Inputs       : Workstation - Name of the workstation
'                Users() - Returned list of users
' Returns      : False if failed
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim llReply As Long                     ' Holds API Reply
Dim llEntriesRead As Long               ' Total number of entries returned by API call
Dim llTotalEntries As Long              ' Total number of entries to be recalled
Dim lhResume As Long                    ' Handle to resume the call if llEntriesRead<>llTotalEntries
Dim llBuffer As Long                    ' Structure Pointer
Dim llLevel As Long                     ' Structure type to retrieve
Dim llPrefMaxLen As Long                ' Max amount of data to return
Dim ltUserInfo As NAME_INFO             ' Holds data returned
Dim i As Long                           ' For Counter
Dim k As Long                           ' For Counter
Dim llUserCount As Long                 ' Counts the found Users
Dim lsUser As String                    ' Local copy of name
Dim lbFound As Boolean                  ' Flag to help prevent duplicates
    
On Error GoTo Catch
    
' API return codes
Const SV_SUCCESS As Long = 0
Const SV_MORE_DATA As Long = 234&

Const MAX_PREFERRED_LENGTH As Long = -1
Const ST_NAME_INFO As Long = 0

' Defaults
llLevel = ST_NAME_INFO      ' Basic Structure
llPrefMaxLen = MAX_PREFERRED_LENGTH ' No Limit on whot can be returned
ReDim Users(0 To 0)                 ' if Ubound(Users) = 0 then its empty
    
Do
    ' Call depends on whether we have a WorkStation name
    If WorkStation = vbNullString Then
        llReply = NetWkstaUserEnum(0&, llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    Else
        llReply = NetWkstaUserEnum(StrPtr(WorkStation), llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    End If
            
    ' See if we succeeded
    If llReply = SV_SUCCESS Or llReply = SV_MORE_DATA Then
        
        ' Iterate through all the returned Users
        For i = 0 To llEntriesRead - 1
            
            ' Copy the data into the struture
            RtlMoveMemory ltUserInfo, ByVal llBuffer, Len(ltUserInfo)
            
            ' Get the user name
            lsUser = WinPointerToString(ltUserInfo.name)
            ' set the default falg
            lbFound = False
            
            ' Only check if we have something in the array
            If llUserCount > 0 Then
                ' Walk through the names we alreday have
                For k = 1 To UBound(Users)
                    ' case sensitive test
                    If Users(k) = lsUser Then
                        lbFound = True
                        Exit For
                    End If
                Next k

            End If
            
            ' If the name is an original - add it
            If Not lbFound Then
                llUserCount = llUserCount + 1
                If llUserCount = 1 Then
                    ReDim Users(1 To 1)
                Else
                    ReDim Preserve Users(1 To llUserCount)
                End If
                Users(llUserCount) = lsUser
            End If
            
            ' Adjust the pointer to the next entry
            llBuffer = llBuffer + Len(ltUserInfo)
        Next i
    End If
    
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer

' As we've set the maximum llPrefMaxLen = -1 then it SHOULD only require a single call plus
' a bug has been associated with the call & its resume handle (Can use undocumented call
' NetUserEnumEX if this is a problem)
Loop Until llReply <> SV_MORE_DATA

' Only return true if we got any
If UBound(Users) > 0 Then GetWorkStationUsers = True
    
Finally:
    ' Clean-up
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer
    Exit Function

Catch:
    ' Reports back to developer if in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetWorkStationUsers of Class Module ICE_NetFunctions"
    Resume Finally

End Function

Public Function GetUsers(ByVal Domain As String, ByRef Users() As String) As Boolean
'-----------------------------------------------------------------------
' Procedure    : frmMain.GetUsers
' Author       : GWilmot
' Date Created : 04/12/2002
'-----------------------------------------------------------------------
' Purpose      : Returns a list of users for a given Domain
' Assumptions  :
' Inputs       : Domain - Normally the PDC
'                Users() - list of users
' Returns      : False if failed
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim llReply As Long                     ' Holds API Reply
Dim llEntriesRead As Long               ' Total number of entries returned by API call
Dim llTotalEntries As Long              ' Total number of entries to be recalled
Dim lhResume As Long                    ' Handle to resume the call if llEntriesRead<>llTotalEntries
Dim llBuffer As Long                    ' Structure Pointer
Dim llLevel As Long                     ' Structure type to retrieve
Dim llPrefMaxLen As Long                ' Max amount of data to return
Dim ltUserInfo As NAME_INFO     ' Holds data returned
Dim i As Long                           ' For Counter
Dim llUserCount As Long                 ' Counts the found Users
Dim llFilter As Long                    ' Can filter Users

On Error GoTo Catch

' API return codes
Const SV_SUCCESS As Long = 0
Const SV_MORE_DATA As Long = 234&

Const MAX_PREFERRED_LENGTH As Long = -1
Const ST_NAME_INFO As Long = 0
Const FT_ALL As Long = 0

' Defaults
llLevel = ST_NAME_INFO      ' Basic Structure
llFilter = FT_ALL                   ' No filtering of accounts
llPrefMaxLen = MAX_PREFERRED_LENGTH ' No Limit on whot can be returned
ReDim Users(0 To 0)                 ' if Ubound(Users) = 0 then its empty
    
Do
    ' Call depends on whether we have a Domain name
    If Domain = vbNullString Then
        llReply = NetUserEnum(0&, llLevel, llFilter, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    Else
        llReply = NetUserEnum(StrPtr(Domain), llLevel, llFilter, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    End If
            
    ' See if we succeeded
    If llReply = SV_SUCCESS Or llReply = SV_MORE_DATA Then
        
        ' Iterate through all the returned Users
        For i = 0 To llEntriesRead - 1
            
            ' Copy the data into the struture
            RtlMoveMemory ltUserInfo, ByVal llBuffer, Len(ltUserInfo)
            
            ' Increment our User count & redim array
            llUserCount = llUserCount + 1
            If llUserCount = 1 Then
                ReDim Users(1 To 1)
            Else
                ReDim Preserve Users(1 To llUserCount)
            End If
            
            ' OK get it to the array by taking the structure pointer
            ' and getting the string of the name
            Users(llUserCount) = WinPointerToString(ltUserInfo.name)
            
            ' Adjust the pointer to the next entry
            llBuffer = llBuffer + Len(ltUserInfo)
        Next i
    End If
    
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer

' As we've set the maximum llPrefMaxLen = -1 then it SHOULD only require a single call plus
Loop Until llReply <> SV_MORE_DATA

' Only return true if we got any
If UBound(Users) > 0 Then GetUsers = True
    
Finally:
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer
    Exit Function

Catch:
    ' Reports back to developer if in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetUsers of Form frmMain"
    Resume Finally

End Function
Public Function GetPrimaryDomainController() As String
'-----------------------------------------------------------------------
' Procedure    : ICE_NetFunctions.GetPrimaryDomainController
' Author       : GWilmot
' Date Created : 07/12/2002
'-----------------------------------------------------------------------
' Purpose      : Returns the PDC for the current workstation
' Assumptions  :
' Inputs       : None
' Returns      : The name of the PDC or nullstring if failed
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim lsServer As String   ' Holds the name of the server
Dim llBuffer As Long     ' Holds pointer to name
Dim laByte() As Byte     ' Holds name
Dim llGotNameOK As Long  ' Reply from the API call
Dim llBufferOK As Long   ' Reply from clear buffer call
   
On Error GoTo Catch
   
Const SUCCESS = 0

' This will return nothing if the machine is not in a domain


llGotNameOK = NetGetDCName(vbNullString, vbNullString, llBuffer)

If llGotNameOK = SUCCESS Then ' success
   
    ReDim laByte(256)
    
    ' llBuffer is a pointer so copy contents using API call
    RtlMoveMemory laByte(0), llBuffer, 256
    ' Convert from byte array to string
    lsServer = laByte
    ' Set to return name strip off trailing rubbish
    GetPrimaryDomainController = NullTrim(lsServer)
    ' Free llBuffer - not in other samples but mentioned in documentation
    llBufferOK = NetApiBufferFree(llBuffer)
    
Else
    GetPrimaryDomainController = vbNullString
End If
   
Finally:
    ' Clean-up
    Exit Function

Catch:

    ' Reports back to developer if in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetPrimaryDomainController of Class Module ICE_NetFunctions"
    Resume Finally

End Function
Private Function NullTrim(ByVal Convert As Variant) As String
'-----------------------------------------------------------------------
' Procedure    : ICE_NetFunctions.NullTrim
' Author       : GWilmot
' Date Created : 07/12/2002
'-----------------------------------------------------------------------
' Purpose      : Trims a string for null characters
' Assumptions  : Errors passed up to calling routine
' Inputs       : Value to be converted
' Returns      : Trimmed string
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim i As Integer          ' Counter
Dim lsString  As String   ' Holds variable once converted to string

'*** Errors passed up to calling routine

' Sort out any required convertion
lsString = vbNullString
Select Case VarType(Convert)
    Case vbByte, vbArray + vbByte
        '*** [GW need some work here]
        'lsString = CStr(Convert)
    Case vbString
        lsString = Convert
    Case vbEmpty
    Case vbNull
    Case vbInteger
        lsString = CStr(Convert)
    Case vbLong
        lsString = CStr(Convert)
    Case vbSingle
        lsString = CStr(Convert)
    Case vbDouble
        lsString = CStr(Convert)
    Case vbCurrency
        lsString = CStr(Convert)
    Case vbDate
        lsString = Format$(Convert, "General Date")
    Case vbBoolean
        lsString = Convert
    Case vbObject
    Case vbError
    Case vbVariant
    Case vbDataObject
    Case vbArray
End Select

' Now check for any Null terminate
i = InStr(lsString, vbNullChar)
If i > 0 Then
    NullTrim = Trim(Left$(lsString, i - 1))
Else
    NullTrim = Trim(lsString)
End If
    
End Function

Public Function GetGroups(ByVal Domain As String, ByRef Groups() As String) As Boolean
'-----------------------------------------------------------------------
' Procedure    : frmMain.GetGroups
' Author       : GWilmot
' Date Created : 04/12/2002
'-----------------------------------------------------------------------
' Purpose      : Gets a List of the Groups for a given domain
' Assumptions  :
' Inputs       : Domain - Normally the PDC
'                Groups() returned list of groups
' Returns      : False if failed
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim llReply As Long                     ' Holds API Reply
Dim llEntriesRead As Long               ' Total number of entries returned by API call
Dim llTotalEntries As Long              ' Total number of entries to be recalled
Dim lhResume As Long                    ' Handle to resume the call if llEntriesRead<>llTotalEntries
Dim llBuffer As Long                    ' Structure Pointer
Dim llLevel As Long                     ' Structure type to retrieve
Dim llPrefMaxLen As Long                ' Max amount of data to return
Dim ltGroupInfo As NAME_INFO            ' Holds data returned
Dim i As Long                           ' For Counter
Dim llUserCount As Long                 ' Counts the found Groups

On Error GoTo Catch

' API return codes
Const SV_SUCCESS As Long = 0
Const SV_MORE_DATA As Long = 234&

Const MAX_PREFERRED_LENGTH As Long = -1
Const ST_NAME_INFO As Long = 0
Const FT_ALL As Long = 0

' Defaults
llLevel = ST_NAME_INFO      ' Basic Structure
llPrefMaxLen = MAX_PREFERRED_LENGTH ' No Limit on whot can be returned
ReDim Groups(0 To 0)                ' if Ubound(Groups) = 0 then its empty
    
Do
    ' Call depends on whether we have a Domain name
    If Domain = vbNullString Then
        llReply = NetGroupEnum(0&, llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    Else
        llReply = NetGroupEnum(StrPtr(Domain), llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    End If
            
    ' See if we succeeded
    If llReply = SV_SUCCESS Or llReply = SV_MORE_DATA Then
        
        ' Iterate through all the returned Groups
        For i = 0 To llEntriesRead - 1
            
            ' Copy the data into the struture
            RtlMoveMemory ltGroupInfo, ByVal llBuffer, Len(ltGroupInfo)
            
            ' Increment our User count & redim array
            llUserCount = llUserCount + 1
            If llUserCount = 1 Then
                ReDim Groups(1 To 1)
            Else
                ReDim Preserve Groups(1 To llUserCount)
            End If
            
            ' OK get it to the array by taking the structure pointer
            ' and getting the string of the name
            Groups(llUserCount) = WinPointerToString(ltGroupInfo.name)
            
            ' Adjust the pointer to the next entry
            llBuffer = llBuffer + Len(ltGroupInfo)
        Next i
    End If
    
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer

' As we've set the maximum llPrefMaxLen = -1 then it SHOULD only require a single call plus
Loop Until llReply <> SV_MORE_DATA

' Only return true if we got any
If UBound(Groups) > 0 Then GetGroups = True

Finally:
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer
    Exit Function

Catch:
    ' Reports back to developer if in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetGroups of Form frmMain"
    Resume Finally

End Function
Public Function GetWorkStationShares(ByVal WorkStation As String, ByRef Shares() As String) As Boolean
'-----------------------------------------------------------------------
' Procedure    : frmMain.GetWorkStationShares
' Author       : GWilmot
' Date Created : 04/12/2002
'-----------------------------------------------------------------------
' Purpose      : Returns the shares for a given workstation including hidden shares
' Assumptions  :
' Inputs       : Workstation - Name of the computer
'                Shares() - list of the share names
' Returns      : False if failed
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim llReply As Long                     ' Holds API Reply
Dim llEntriesRead As Long               ' Total number of entries returned by API call
Dim llTotalEntries As Long              ' Total number of entries to be recalled
Dim lhResume As Long                    ' Handle to resume the call if llEntriesRead<>llTotalEntries
Dim llBuffer As Long                    ' Structure Pointer
Dim llLevel As Long                     ' Structure type to retrieve
Dim llPrefMaxLen As Long                ' Max amount of data to return
Dim ltShareInfo As NAME_INFO            ' Holds data returned
Dim i As Long                           ' For Counter
Dim llShareCount As Long                ' Counts the found Shares

On Error GoTo Catch

' API return codes
Const SV_SUCCESS As Long = 0
Const SV_MORE_DATA As Long = 234&

Const MAX_PREFERRED_LENGTH As Long = -1
Const ST_NAME_INFO As Long = 0

' Defaults
llLevel = ST_NAME_INFO      ' Basic Structure
llPrefMaxLen = MAX_PREFERRED_LENGTH ' No Limit on whot can be returned
ReDim Shares(0 To 0)                 ' if Ubound(Shares) = 0 then its empty
    
Do
    ' Call depends on whether we have a WorkStation name
    If WorkStation = vbNullString Then
        llReply = NetShareEnum(0&, llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    Else
        llReply = NetShareEnum(StrPtr(WorkStation), llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    End If
            
    ' See if we succeeded
    If llReply = SV_SUCCESS Or llReply = SV_MORE_DATA Then
        
        ' Iterate through all the returned Shares
        For i = 0 To llEntriesRead - 1
            
            ' Copy the data into the struture
            RtlMoveMemory ltShareInfo, ByVal llBuffer, Len(ltShareInfo)
            
            '*** Need to write some code to take out defaults &
            ' duplicates
            
            ' Increment our Share count & redim array
            llShareCount = llShareCount + 1
            If llShareCount = 1 Then
                ReDim Shares(1 To 1)
            Else
                ReDim Preserve Shares(1 To llShareCount)
            End If
            
            ' OK get it to the array by taking the structure pointer
            ' and getting the string of the name
            Shares(llShareCount) = WinPointerToString(ltShareInfo.name)
            
            ' Adjust the pointer to the next entry
            llBuffer = llBuffer + Len(ltShareInfo)
        Next i
    End If
    
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer

' As we've set the maximum llPrefMaxLen = -1 then it SHOULD only require a single call plus
Loop Until llReply <> SV_MORE_DATA

' Only return true if we got any
If UBound(Shares) > 0 Then GetWorkStationShares = True
    
Finally:
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer
    Exit Function

Catch:
    ' Reports back to developer if in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetWorkStationShares of Form frmMain"
    Resume Finally

End Function

Public Function GetLocalGroups(ByVal Domain As String, ByRef Groups() As String) As Boolean
'-----------------------------------------------------------------------
' Procedure    : frmMain.GetLocalGroups
' Author       : GWilmot
' Date Created : 05/12/2002
'-----------------------------------------------------------------------
' Purpose      : Returns a list of local groups for a given domain
' Assumptions  :
' Inputs       : Domain - Normally the PDC
'                Groups() returned list of groups
' Returns      : False if failed
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim llReply As Long                     ' Holds API Reply
Dim llEntriesRead As Long               ' Total number of entries returned by API call
Dim llTotalEntries As Long              ' Total number of entries to be recalled
Dim lhResume As Long                    ' Handle to resume the call if llEntriesRead<>llTotalEntries
Dim llBuffer As Long                    ' Structure Pointer
Dim llLevel As Long                     ' Structure type to retrieve
Dim llPrefMaxLen As Long                ' Max amount of data to return
Dim ltGroupInfo As NAME_INFO            ' Holds data returned
Dim i As Long                           ' For Counter
Dim llUserCount As Long                 ' Counts the found Groups

On Error GoTo Catch

' API return codes
Const SV_SUCCESS As Long = 0
Const SV_MORE_DATA As Long = 234&

Const MAX_PREFERRED_LENGTH As Long = -1
Const ST_NAME_INFO As Long = 0
Const FT_ALL As Long = 0

' Defaults
llLevel = ST_NAME_INFO      ' Basic Structure
llPrefMaxLen = MAX_PREFERRED_LENGTH ' No Limit on whot can be returned
ReDim Groups(0 To 0)                ' if Ubound(Groups) = 0 then its empty
    
Do
    ' Call depends on whether we have a Domain name
    If Domain = vbNullString Then
        llReply = NetLocalGroupEnum(0&, llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    Else
        llReply = NetLocalGroupEnum(StrPtr(Domain), llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    End If
            
    ' See if we succeeded
    If llReply = SV_SUCCESS Or llReply = SV_MORE_DATA Then
        
        ' Iterate through all the returned Groups
        For i = 0 To llEntriesRead - 1
            
            ' Copy the data into the struture
            RtlMoveMemory ltGroupInfo, ByVal llBuffer, Len(ltGroupInfo)
            
            ' Increment our User count & redim array
            llUserCount = llUserCount + 1
            If llUserCount = 1 Then
                ReDim Groups(1 To 1)
            Else
                ReDim Preserve Groups(1 To llUserCount)
            End If
            
            ' OK get it to the array by taking the structure pointer
            ' and getting the string of the name
            Groups(llUserCount) = WinPointerToString(ltGroupInfo.name)
            
            ' Adjust the pointer to the next entry
            llBuffer = llBuffer + Len(ltGroupInfo)
        Next i
    End If
    
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer

' As we've set the maximum llPrefMaxLen = -1 then it SHOULD only require a single call plus
Loop Until llReply <> SV_MORE_DATA

' Only return true if we got any
If UBound(Groups) > 0 Then GetLocalGroups = True


Finally:
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer
    Exit Function

Catch:
    ' Reports back to developer if in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetGroups of Form frmMain"
    Resume Finally
    
End Function

Public Sub GetNetUserInfo(ByVal Server As String, ByVal username As String, ByRef DisplayText As String)
'-----------------------------------------------------------------------
' Procedure    : ICE_NetFunctions.GetNetUserInfo
' Author       : GWilmot
' Date Created : 07/12/2002
'-----------------------------------------------------------------------
' Purpose      : Returns User Information from a domain
' Assumptions  :
' Inputs       : Server - Domain name normally the PDC
'                Username - Name of the Users
'                DisplayTest - details ready to be displayed
' Returns      :
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim llBuffer As Long               ' Pointer
Dim laUserName() As Byte           ' Username as byte array
Dim laServer() As Byte             ' Server name as byte array
Dim ltUserApi As USER_INFO_3_API   ' Data returned as pointers
Dim llReply As Long                ' Reply to API call
Dim llTotal As Long                ' Holds value from group call
Dim llRead As Long                 ' Holds groups returned
Dim i As Integer                   ' Counter
Dim lsDummy As String              ' Temp storage
Dim lsParameters As String         ' Holds the parameter string
Dim k As Integer                   ' Counter
Dim liCounter As Integer           ' Counts
Dim lpDate As Date                 ' Holds the source date
Dim lpDateV As Date                ' Holds the source date
Dim lfUserDetails As Form          ' Form used to display User deatils

Const flags& = 0
Const NERR_Success As Long = 0&

Const USER_PRIV_MASK = 3
Const USER_PRIV_GUEST = 0
Const USER_PRIV_USER = 1
Const USER_PRIV_ADMIN = 2

On Error GoTo Catch

' set defaults
lpDate = #1/1/1970#

' Configure the username
laUserName = username & vbNullChar

' See if we have a servername
If Server = vbNullString Then
    ' If not use the default
    llReply = NetUserGetInfo(ByVal 0&, laUserName(0), 3, llBuffer)
Else
    ' OK Configure the servername correctly
    If InStr(Server, "\\") = 1 Then
        laServer = Server & vbNullChar
    Else
        laServer = "\\" & Server & vbNullChar
    End If
    llReply = NetUserGetInfo(laServer(0), laUserName(0), 3, llBuffer)
End If
   
' Sooo did it work
If llReply = NERR_Success Then

    ' Use the pointer to fill the API structure
    RtlMoveMemory ltUserApi, ByVal llBuffer, Len(ltUserApi)
    
    DisplayText = "Fullname : " & WinPointerToString(ltUserApi.FullName) & vbCrLf
    DisplayText = DisplayText & "Password : " & WinPointerToString(ltUserApi.Password) & vbCrLf
    ' No of Secs....
    DisplayText = DisplayText & "Password Age : " & Format$(ltUserApi.PasswordAge / 86400, ".0 days") & vbCrLf & vbCrLf
        
    Select Case (ltUserApi.Privilege And USER_PRIV_MASK)
        Case USER_PRIV_GUEST
            DisplayText = DisplayText & "User Type : Guest" & vbCrLf
        Case USER_PRIV_USER
            DisplayText = DisplayText & "User Type : User" & vbCrLf
        Case USER_PRIV_ADMIN
            DisplayText = DisplayText & "User Type : Administrator" & vbCrLf
    End Select
    
    DisplayText = DisplayText & "Workstations : " & WinPointerToString(ltUserApi.Workstations) & vbCrLf
    DisplayText = DisplayText & "Home Directory : " & WinPointerToString(ltUserApi.HomeDir) & vbCrLf
    DisplayText = DisplayText & "Comments : " & WinPointerToString(ltUserApi.Comment) & vbCrLf & vbCrLf
    
    DisplayText = DisplayText & "Script Path : " & WinPointerToString(ltUserApi.ScriptPath) & vbCrLf
    DisplayText = DisplayText & "User Comments : " & WinPointerToString(ltUserApi.UserComment) & vbCrLf
    
    ' Date...Format(User.LastLogonDate, "long date")
    If ltUserApi.LastLogon <> 0 Then
        lpDateV = DateAdd("s", ltUserApi.LastLogon, lpDate)
        DisplayText = DisplayText & "Last Log On : " & Format$(lpDateV, "short date") & " " & Format$(lpDateV, "short time") & vbCrLf
    End If
    ' Date...Format(User.LastLogonDate, "long date")
    If ltUserApi.LastLogoff <> 0 Then
        lpDateV = DateAdd("s", ltUserApi.LastLogoff, lpDate)
        DisplayText = DisplayText & "Last Log Off : " & Format$(lpDateV, "short date") & " " & Format$(lpDateV, "short time") & vbCrLf
    End If
        
    DisplayText = DisplayText & "No. of Log Ons : " & CStr(ltUserApi.NumLogons) & vbCrLf
    DisplayText = DisplayText & "No. of Bad passwords : " & CStr(ltUserApi.BadPwCount) & vbCrLf
        
    If ltUserApi.AcctExpires = -1 Then
        lfUserDetails.txtAcctExpires = "Never"
        DisplayText = DisplayText & "Password : Never Expires " & vbCrLf
    Else
        lpDateV = DateAdd("s", ltUserApi.AcctExpires, lpDate)
        DisplayText = DisplayText & "Password Expires : " & Format$(lpDateV, "short date") + " " + Format$(lpDateV, "short time") & vbCrLf
    End If

End If

Finally:
    ' Clear down the buffer if required
    If llBuffer Then NetApiBufferFree llBuffer
    Exit Sub

Catch:
    ' Reports back to developer if in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetNetUserInfo of Form frmMain"
    Resume Finally

End Sub

Public Function GetGroupUsers(ByVal Domain As String, ByVal Group As String, ByRef GroupUsers() As String) As Boolean
'-----------------------------------------------------------------------
' Procedure    : frmMain.GetGroupUsers
' Author       : GWilmot
' Date Created : 05/12/2002
'-----------------------------------------------------------------------
' Purpose      : Gets a list of users for a given group
' Assumptions  :
' Inputs       : Domain - Normally the PDC
'                Group - Name of the group
'                GroupUsers() - list of group users
' Returns      : False if failed
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim llReply As Long                     ' Holds API Reply
Dim llEntriesRead As Long               ' Total number of entries returned by API call
Dim llTotalEntries As Long              ' Total number of entries to be recalled
Dim lhResume As Long                    ' Handle to resume the call if llEntriesRead<>llTotalEntries
Dim llBuffer As Long                    ' Structure Pointer
Dim llLevel As Long                     ' Structure type to retrieve
Dim llPrefMaxLen As Long                ' Max amount of data to return
Dim ltGroupInfo As NAME_INFO            ' Holds data returned
Dim i As Long                           ' For Counter
Dim llUserCount As Long                 ' Counts the found GroupUsers

On Error GoTo Catch

' API return codes
Const SV_SUCCESS As Long = 0
Const SV_MORE_DATA As Long = 234&

Const MAX_PREFERRED_LENGTH As Long = -1
Const ST_NAME_INFO As Long = 0
Const FT_ALL As Long = 0

' Defaults
llLevel = ST_NAME_INFO      ' Basic Structure
llPrefMaxLen = MAX_PREFERRED_LENGTH ' No Limit on whot can be returned
ReDim GroupUsers(0 To 0)                ' if Ubound(GroupUsers) = 0 then its empty
    
Do
    ' Call depends on whether we have a Domain name
    If Domain = vbNullString Then
        llReply = NetGroupGetUsers(0&, llLevel, StrPtr(Group), llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    Else
        llReply = NetGroupGetUsers(StrPtr(Domain), StrPtr(Group), llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    End If
    
    ' See if we succeeded
    If llReply = SV_SUCCESS Or llReply = SV_MORE_DATA Then
        
        ' Iterate through all the returned GroupUsers
        For i = 0 To llEntriesRead - 1
            
            ' Copy the data into the struture
            RtlMoveMemory ltGroupInfo, ByVal llBuffer, Len(ltGroupInfo)
            
            ' Increment our User count & redim array
            llUserCount = llUserCount + 1
            If llUserCount = 1 Then
                ReDim GroupUsers(1 To 1)
            Else
                ReDim Preserve GroupUsers(1 To llUserCount)
            End If
            
            ' OK get it to the array by taking the structure pointer
            ' and getting the string of the name
            GroupUsers(llUserCount) = WinPointerToString(ltGroupInfo.name)
            
            ' Adjust the pointer to the next entry
            llBuffer = llBuffer + Len(ltGroupInfo)
        Next i
    End If
    
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer

' As we've set the maximum llPrefMaxLen = -1 then it SHOULD only require a single call plus
Loop Until llReply <> SV_MORE_DATA

' Only return true if we got any
If UBound(GroupUsers) > 0 Then GetGroupUsers = True


Finally:
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer
    Exit Function

Catch:
    ' Reports back to developer if in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetGroupUsers of Form frmMain"
    Resume Finally

End Function

Public Function GetLocalGroupUsers(ByVal Domain As String, ByVal Group As String, ByRef GroupUsers() As String) As Boolean
'-----------------------------------------------------------------------
' Procedure    : frmMain.GetLocalGroupUsers
' Author       : GWilmot
' Date Created : 05/12/2002
'-----------------------------------------------------------------------
' Purpose      : Gets a list of users for a given local group
' Assumptions  :
' Inputs       : Domain - Normally the PDC
'                Group - Name of the local group
'                GroupUsers() - list of local group users
' Returns      : False if failed
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim llReply As Long                     ' Holds API Reply
Dim llEntriesRead As Long               ' Total number of entries returned by API call
Dim llTotalEntries As Long              ' Total number of entries to be recalled
Dim lhResume As Long                    ' Handle to resume the call if llEntriesRead<>llTotalEntries
Dim llBuffer As Long                    ' Structure Pointer
Dim llLevel As Long                     ' Structure type to retrieve
Dim llPrefMaxLen As Long                ' Max amount of data to return
Dim ltGroupInfo As LOCAL_GROUP          ' Holds data returned
Dim i As Long                           ' For Counter
Dim llUserCount As Long                 ' Counts the found GroupUsers

On Error GoTo Catch

' API return codes
Const SV_SUCCESS As Long = 0
Const SV_MORE_DATA As Long = 234&

Const MAX_PREFERRED_LENGTH As Long = -1
Const ST_NAME_INFO As Long = 1
Const FT_ALL As Long = 0

' Defaults
llLevel = ST_NAME_INFO      ' Basic Structure
llPrefMaxLen = MAX_PREFERRED_LENGTH ' No Limit on whot can be returned
ReDim GroupUsers(0 To 0)                ' if Ubound(GroupUsers) = 0 then its empty
    
Do
    ' Call depends on whether we have a Domain name
    If Domain = vbNullString Then
        llReply = NetLocalGroupGetMembers(0&, llLevel, StrPtr(Group), llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    Else
        llReply = NetLocalGroupGetMembers(StrPtr(Domain), StrPtr(Group), llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    End If
    
    ' See if we succeeded
    If llReply = SV_SUCCESS Or llReply = SV_MORE_DATA Then
        
        ' Iterate through all the returned GroupUsers
        For i = 0 To llEntriesRead - 1
            
            ' Copy the data into the struture
            RtlMoveMemory ltGroupInfo, ByVal llBuffer, Len(ltGroupInfo)
            
            ' Increment our User count & redim array
            llUserCount = llUserCount + 1
            If llUserCount = 1 Then
                ReDim GroupUsers(1 To 1)
            Else
                ReDim Preserve GroupUsers(1 To llUserCount)
            End If
            
            ' OK get it to the array by taking the structure pointer
            ' and getting the string of the name
            GroupUsers(llUserCount) = WinPointerToString(ltGroupInfo.name)
            
            ' Adjust the pointer to the next entry
            llBuffer = llBuffer + Len(ltGroupInfo)
        Next i
    End If
    
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer

' As we've set the maximum llPrefMaxLen = -1 then it SHOULD only require a single call plus
Loop Until llReply <> SV_MORE_DATA

' Only return true if we got any
If UBound(GroupUsers) > 0 Then GetLocalGroupUsers = True


Finally:
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer
    Exit Function

Catch:
    ' Reports back to developer if in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetGroupUsers of Form frmMain"
    Resume Finally

End Function

Private Function WinPointerToString(ByVal StringPtr As Long) As String
'-----------------------------------------------------------------------
' Procedure    : ICE_NetFunctions.WinPointerToString
' Author       : GWilmot
' Date Created : 07/12/2002
'-----------------------------------------------------------------------
' Purpose      : Returns the string associated with a pointer
' Assumptions  : Errors passed up to calling routine
' Inputs       : StringPtr - String Pointer in memory
' Returns      : The String
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------

Dim labyteBuffer() As Byte     ' Holds the data as a Byte Form
Dim llLength As Long           ' Holds the amount of data

'*** Errors passed up to calling routine

' If it ain't Zero
If StringPtr Then
    ' Get the amount of data
    llLength = lstrlenW(StringPtr) * 2
    ' If it ain't Zero
    If llLength Then
        ' Set the Buffer
        ReDim labyteBuffer(0 To (llLength - 1)) As Byte
        ' Fill it
        RtlMoveMemory labyteBuffer(0), ByVal StringPtr, llLength
        ' Set the return value
        WinPointerToString = labyteBuffer
    End If
End If

End Function

Public Function GetUserSessions(ByVal Domain As String, ByRef Sessions() As String) As Boolean
'-----------------------------------------------------------------------
' Procedure    : ICE_NetFunctions.GetUserSessions
' Author       : GWilmot
' Date Created : 07/12/2002
'-----------------------------------------------------------------------
' Purpose      : Gets the list of users/workstations for a given domain
' Assumptions  :
' Inputs       : Domain - Any computer
'                Sessions - list of User + Workstation
' Returns      : False if Failed
' Effects      :
' Last Updated :
'-----------------------------------------------------------------------
Dim llReply As Long                     ' Holds API Reply
Dim llEntriesRead As Long               ' Total number of entries returned by API call
Dim llTotalEntries As Long              ' Total number of entries to be recalled
Dim lhResume As Long                    ' Handle to resume the call if llEntriesRead<>llTotalEntries
Dim llBuffer As Long                    ' Structure Pointer
Dim llLevel As Long                     ' Structure type to retrieve
Dim llPrefMaxLen As Long                ' Max amount of data to return
Dim ltSessionsInfo As SESSION_INFO_10   ' Holds data returned

Dim i As Long                           ' For Counter
Dim llSessions As Long                  ' Counts the found Sessions

On Error GoTo Catch

' API return codes
Const SV_SUCCESS As Long = 0
Const SV_MORE_DATA As Long = 234&

Const MAX_PREFERRED_LENGTH As Long = -1
Const ST_SESSION_INFO_10 As Long = 10
Const FT_ALL As Long = 0

' Defaults
llLevel = ST_SESSION_INFO_10        ' Basic Structure
llPrefMaxLen = MAX_PREFERRED_LENGTH ' No Limit on whot can be returned
ReDim Sessions(0 To 0)              ' if Ubound(Sessions) = 0 then its empty
    
Do
    ' Call depends on whether we have a Domain name
    If Domain = vbNullString Then
        llReply = NetSessionEnum(0&, 0&, 0&, llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    Else
        llReply = NetSessionEnum(StrPtr(Domain), 0&, 0&, llLevel, llBuffer, llPrefMaxLen, _
            llEntriesRead, llTotalEntries, lhResume)
    End If
    
    ' See if we succeeded
    If llReply = SV_SUCCESS Or llReply = SV_MORE_DATA Then
        
        ' Iterate through all the returned Sessions
        For i = 0 To llEntriesRead - 1
            
            ' Copy the data into the struture
            RtlMoveMemory ltSessionsInfo, ByVal llBuffer, Len(ltSessionsInfo)
            
            ' Increment our User count & redim array
            llSessions = llSessions + 1
            If llSessions = 1 Then
                ReDim Sessions(1 To 1)
            Else
                ReDim Preserve Sessions(1 To llSessions)
            End If
            
            ' OK get it to the array by taking the structure pointer
            ' and getting the string of the name
            Sessions(llSessions) = WinPointerToString(ltSessionsInfo.username) & " : " & WinPointerToString(ltSessionsInfo.name)
            
            ' Adjust the pointer to the next entry
            llBuffer = llBuffer + Len(ltSessionsInfo)
        Next i
    End If
    
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer

' As we've set the maximum llPrefMaxLen = -1 then it SHOULD only require a single call
Loop Until llReply <> SV_MORE_DATA

' Only return true if we got any
If UBound(Sessions) > 0 Then GetUserSessions = True

Finally:
    ' See if we need to clear down the buffer
    If llBuffer Then NetApiBufferFree llBuffer
    Exit Function

Catch:
    ' Reports back to developer if in debugmode
    If nbDebugMode Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetUserSessions of Class Module ICE_NetFunctions"
    Resume Finally

End Function
