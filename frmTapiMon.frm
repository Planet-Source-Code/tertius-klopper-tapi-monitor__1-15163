VERSION 5.00
Begin VB.Form frmTapiMon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TAPI Monitor                                                                             "
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstStatus 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1980
      ItemData        =   "frmTapiMon.frx":0000
      Left            =   120
      List            =   "frmTapiMon.frx":0002
      TabIndex        =   0
      Top             =   720
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "Tapi Monitor will monitor outgoing calls from your modem"
      Height          =   375
      Left            =   1148
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmTapiMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This app may cause VB to give a exeption error, when running it
'in compiled mode. To be save first save all your work before
'running it
'By Tertius at tertiusklopper@hotmail.com for more information
'Was designed as part of an app I wrote to calculate the cost of
'a call
'Parts of program was translated from Delphi 4 as my first app
'was written in that.

'A Lot of the Declared functions and types in vbTapi.bas is not
'used in this app but it will help you to use the without
'declaring it yourself

'vbTapi.bas was found on Planet-Source-code.com but do not
'remember who made it (To the persone who did it THANK YOU)

Dim udtLineCall As LINECALLPARAMS
Dim lines As Long
Dim hInst As Long
Dim lineApp As Long
Dim lphLine As Long
Dim lphCall As Long
Dim adrCallBack As Long

Private Sub Form_Load()
Dim nDevs As Long
Dim tapiVer As Long
Dim extid As LINEEXTENSIONID

If lineInitialize(lineApp, hInst, AddressOf LINECALLBACK, 0, nDevs) < 0 Then
    lineApp = 0
ElseIf nDevs = 0 Then  'No Tapi Device
    lineShutdown (lineApp)
    lineApp = 0
ElseIf lineNegotiateAPIVersion(lineApp, 0, 65536, 65540, tapiVer, extid) < 0 Then 'Check for version
    lineShutdown (lineApp)
    lineApp = 0
    lphLine = 0
    'Open a line for monitor (here I use the first device, normally the modem
ElseIf lineOpen(lineApp, 0, lphLine, tapiVer, 0, 0, LINECALLPRIVILEGE_MONITOR, LINEMEDIAMODE_DATAMODEM, 0) < 0 Then
    lineShutdown (lineApp)
    lineApp = 0
    lphLine = 0
End If

If lineApp <> 0 Then
    lstStatus.AddItem ("Monitoring Calls...")
    lstStatus.TopIndex = lstStatus.ListCount - 1
Else
    lstStatus.AddItem ("Error!")
    lstStatus.TopIndex = lstStatus.ListCount - 1
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
If lphLine <> 0 Then
   lineClose (lphLine)
End If
If lineApp <> 0 Then
   lineShutdown (lineApp)
End If
End Sub
