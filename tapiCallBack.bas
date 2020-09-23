Attribute VB_Name = "tapiCallBack"
Option Explicit
Global CallInfo As LINECALLINFO

Public Sub LINECALLBACK(ByVal hDevice As Long, ByVal dwMessage As Long, ByVal dwInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)

    'Handels messages from Tapi32
    
    Dim strTemp As String
    Dim hCall As Long
    
    If dwMessage = LINE_CALLSTATE Then
       hCall = PtrToLong(hDevice)
       Select Case dwParam1
          Case LINECALLSTATE_IDLE 'Call Terminated
            If hCall <> 0 Then
             lineDeallocateCall (hCall)
             frmTapiMon.lstStatus.AddItem ("Idle")
             frmTapiMon.lstStatus.TopIndex = _
             frmTapiMon.lstStatus.ListCount - 1
            End If
          Case LINECALLSTATE_DIALING ' Call Dialing
             frmTapiMon.lstStatus.AddItem ("Dialing Call")
             frmTapiMon.lstStatus.TopIndex = _
             frmTapiMon.lstStatus.ListCount - 1
          Case LINECALLSTATE_CONNECTED 'Service Connected
            If hCall <> 0 Then
             frmTapiMon.lstStatus.AddItem ("Connected")
             frmTapiMon.lstStatus.TopIndex = _
             frmTapiMon.lstStatus.ListCount - 1
             CallInfo.dwTotalSize = 4096
            End If
          Case LINECALLSTATE_PROCEEDING 'Call Proceeding (dialing)
            frmTapiMon.lstStatus.AddItem ("Proceeding")
            frmTapiMon.lstStatus.TopIndex = _
            frmTapiMon.lstStatus.ListCount - 1
          Case LINECALLSTATE_DISCONNECTED 'Disconnected
            frmTapiMon.lstStatus.AddItem ("Disconnected")
            frmTapiMon.lstStatus.TopIndex = _
            frmTapiMon.lstStatus.ListCount - 1
          Case LINECALLSTATE_BUSY 'Line Busy
            frmTapiMon.lstStatus.AddItem ("Line Busy")
            frmTapiMon.lstStatus.TopIndex = _
            frmTapiMon.lstStatus.ListCount - 1
           
       End Select
   End If
End Sub

Public Function PtrToLong(ByVal lngFnPtr As Long) As Long
'Convert Pointer into Long
    PtrToLong = lngFnPtr
End Function


