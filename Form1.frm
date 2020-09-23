VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Win2k/XP Task manager 'Processes' clone using Windows Terminal Services APIs"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   4260
      Left            =   105
      TabIndex        =   1
      Top             =   780
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   7514
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   435
      Left            =   195
      TabIndex        =   0
      Top             =   180
      Width           =   1110
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal sID As Long, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Private Declare Function WTSEnumerateProcesses Lib "wtsapi32.dll" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long
Private Declare Sub WTSFreeMemory Lib "wtsapi32.dll" (ByVal pMemory As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WTS_CURRENT_SERVER_HANDLE = 0&

Private Type WTS_PROCESS_INFO
    SessionID As Long
    ProcessID As Long
    pProcessName As Long
    pUserSid As Long
    End Type


Private Sub cmdRefresh_click()
    GetWTSProcesses
    End Sub

Private Function GetStringFromLP(ByVal StrPtr As Long) As String
    Dim b As Byte
    Dim tempStr As String
    Dim bufferStr As String
    Dim Done As Boolean

    Done = False
    Do
        ' Get the byte/character that StrPtr is pointing to.
        CopyMemory b, ByVal StrPtr, 1
        If b = 0 Then  ' If you've found a null character, then you're done.
            Done = True
        Else
            tempStr = Chr$(b)  ' Get the character for the byte's value
            bufferStr = bufferStr & tempStr 'Add it to the string
                
            StrPtr = StrPtr + 1  ' Increment the pointer to next byte/char
        End If
    Loop Until Done
    GetStringFromLP = bufferStr
    End Function

Private Sub Form_Load()
    ListView1.View = lvwReport

    'Add the Column Headers for your ListView Control
    ListView1.ColumnHeaders.Add 1, "SessionID", "Session ID"
    ListView1.ColumnHeaders.Add 2, "ProcessID", "Process ID"
    ListView1.ColumnHeaders.Add 3, "ProcessName", "Process Name"
    ListView1.ColumnHeaders.Add 4, "UserID", "User ID"
    ListView1.ColumnHeaders(4).Width = ListView1.Width - (ListView1.ColumnHeaders(1).Width * 3) - 300

    GetWTSProcesses
    End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    ListView1.SortKey = ColumnHeader.Index - 1
    ' Set Sorted to True to sort the list.
    ListView1.Sorted = True
    End Sub

Private Sub GetWTSProcesses()
   Dim RetVal As Long
   Dim Count As Long
   Dim i As Integer
   Dim lpBuffer As Long
   Dim p As Long
   Dim udtProcessInfo As WTS_PROCESS_INFO
   Dim itmAdd As ListItem

   ListView1.ListItems.Clear
   RetVal = WTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, Count)
   If RetVal Then ' WTSEnumerateProcesses was successful
      p = lpBuffer
        For i = 1 To Count
            ' Count is the number of Structures in the buffer
            ' WTSEnumerateProcesses returns a pointer, so copy it to a
            ' WTS_PROCESS_INO UDT so you can access its members
            CopyMemory udtProcessInfo, ByVal p, LenB(udtProcessInfo)
            ' Add items to the ListView control
            Set itmAdd = ListView1.ListItems.Add(i, , CStr(udtProcessInfo.SessionID))
                itmAdd.SubItems(1) = CStr(udtProcessInfo.ProcessID)
                ' Since pProcessName contains a pointer, call GetStringFromLP to get the
                ' variable length string it points to
                If udtProcessInfo.ProcessID = 0 Then
                    itmAdd.SubItems(2) = "System Idle Process"
                Else
                    itmAdd.SubItems(2) = GetStringFromLP(udtProcessInfo.pProcessName)
                End If
                
                'itmAdd.SubItems(3) = CStr(udtProcessInfo.pUserSid)
                itmAdd.SubItems(3) = GetUserName(udtProcessInfo.pUserSid)

                ' Increment to next WTS_PROCESS_INO structure in the buffer
                p = p + LenB(udtProcessInfo)
        Next i

        Set itmAdd = Nothing
        WTSFreeMemory lpBuffer   'Free your memory buffer
    Else
        ' Error occurred calling WTSEnumerateProcesses
        ' Check Err.LastDllError for error code
        MsgBox "Error occurred calling WTSEnumerateProcesses.  " & "Check the Platform SDK error codes in the MSDN Documentation" & " for more information.", vbCritical, "Error " & Err.LastDllError
    End If
    End Sub

Function GetUserName(sID As Long) As String
    On Error Resume Next
    Dim retname As String
    Dim retdomain As String
    retname = String(255, 0)
    retdomain = String(255, 0)
    LookupAccountSid vbNullString, sID, retname, 255, retdomain, 255, 0
    GetUserName = Left$(retdomain, InStr(retdomain, vbNullChar) - 1) & "\" & Left$(retname, InStr(retname, vbNullChar) - 1)
    End Function
