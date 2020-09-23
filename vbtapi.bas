Attribute VB_Name = "vbtapi"
' The  Telephony  API  is jointly copyrighted by Intel and Microsoft.  You are
' granted  a royalty free worldwide, unlimited license to make copies, and use
' the   API/SPI  for  making  applications/drivers  that  interface  with  the
' specification provided that this paragraph and the Intel/Microsoft copyright
' statement is maintained as is in the text and source code files.
'
' Copyright 1992, 1993 Intel/Microsoft, all rights reserved.

'
' typedef of the LINE callback procedure
'
' Sub LINECALLBACK (ByVal hDevice As Long, ByVal dwMessage As Long, ByVal dwInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
'
' typedef of the PHONE callback procedure
'
' Sub PHONECALLBACK (ByVal hDevice As Long, ByVal dwMessage As Long, ByVal dwInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
'

' Messages for Phones and Lines

Global Const LINE_ADDRESSSTATE = 0&
Global Const LINE_CALLINFO = 1&
Global Const LINE_CALLSTATE = 2&
Global Const LINE_CLOSE = 3&
Global Const LINE_DEVSPECIFIC = 4&
Global Const LINE_DEVSPECIFICFEATURE = 5&
Global Const LINE_GATHERDIGITS = 6&
Global Const LINE_GENERATE = 7&
Global Const LINE_LINEDEVSTATE = 8&
Global Const LINE_MONITORDIGITS = 9&
Global Const LINE_MONITORMEDIA = 10&
Global Const LINE_MONITORTONE = 11&
Global Const LINE_REPLY = 12&
Global Const LINE_REQUEST = 13&
Global Const PHONE_BUTTON = 14&
Global Const PHONE_CLOSE = 15&
Global Const PHONE_DEVSPECIFIC = 16&
Global Const PHONE_REPLY = 17&
Global Const PHONE_STATE = 18&

' Define Simple Telephony Constants.

Global Const TAPI_REPLY = &H400& + 99&

Global Const TAPIERR_CONNECTED = 0&
Global Const TAPIERR_DROPPED = -1&
Global Const TAPIERR_NOREQUESTRECIPIENT = -2&
Global Const TAPIERR_REQUESTQUEUEFULL = -3&
Global Const TAPIERR_INVALDESTADDRESS = -4&
Global Const TAPIERR_INVALWINDOWHANDLE = -5&
Global Const TAPIERR_INVALDEVICECLASS = -6&
Global Const TAPIERR_INVALDEVICEID = -7&
Global Const TAPIERR_DEVICECLASSUNAVAIL = -8&
Global Const TAPIERR_DEVICEIDUNAVAIL = -9&
Global Const TAPIERR_DEVICEINUSE = -10&
Global Const TAPIERR_DESTBUSY = -11&
Global Const TAPIERR_DESTNOANSWER = -12&
Global Const TAPIERR_DESTUNAVAIL = -13&
Global Const TAPIERR_UNKNOWNWINHANDLE = -14&
Global Const TAPIERR_UNKNOWNREQUESTID = -15&
Global Const TAPIERR_REQUESTFAILED = -16&
Global Const TAPIERR_REQUESTCANCELLED = -17&
Global Const TAPIERR_INVALPOINTER = -18&

Global Const TAPIMAXDESTADDRESSSIZE = 80&
Global Const TAPIMAXAPPNAMESIZE = 40&
Global Const TAPIMAXCALLEDPARTYSIZE = 40&
Global Const TAPIMAXCOMMENTSIZE = 80&
Global Const TAPIMAXDEVICECLASSSIZE = 40&
Global Const TAPIMAXDEVICEIDSIZE = 40&

' Data types and values for Phones

Global Const PHONEBUTTONFUNCTION_UNKNOWN = &H0&
Global Const PHONEBUTTONFUNCTION_CONFERENCE = &H1&
Global Const PHONEBUTTONFUNCTION_TRANSFER = &H2&
Global Const PHONEBUTTONFUNCTION_DROP = &H3&
Global Const PHONEBUTTONFUNCTION_HOLD = &H4&
Global Const PHONEBUTTONFUNCTION_RECALL = &H5&
Global Const PHONEBUTTONFUNCTION_DISCONNECT = &H6&
Global Const PHONEBUTTONFUNCTION_CONNECT = &H7&
Global Const PHONEBUTTONFUNCTION_MSGWAITON = &H8&
Global Const PHONEBUTTONFUNCTION_MSGWAITOFF = &H9&
Global Const PHONEBUTTONFUNCTION_SELECTRING = &HA&
Global Const PHONEBUTTONFUNCTION_ABBREVDIAL = &HB&
Global Const PHONEBUTTONFUNCTION_FORWARD = &HC&
Global Const PHONEBUTTONFUNCTION_PICKUP = &HD&
Global Const PHONEBUTTONFUNCTION_RINGAGAIN = &HE&
Global Const PHONEBUTTONFUNCTION_PARK = &HF&
Global Const PHONEBUTTONFUNCTION_REJECT = &H10&
Global Const PHONEBUTTONFUNCTION_REDIRECT = &H11&
Global Const PHONEBUTTONFUNCTION_MUTE = &H12&
Global Const PHONEBUTTONFUNCTION_VOLUMEUP = &H13&
Global Const PHONEBUTTONFUNCTION_VOLUMEDOWN = &H14&
Global Const PHONEBUTTONFUNCTION_SPEAKERON = &H15&
Global Const PHONEBUTTONFUNCTION_SPEAKEROFF = &H16&
Global Const PHONEBUTTONFUNCTION_FLASH = &H17&
Global Const PHONEBUTTONFUNCTION_DATAON = &H18&
Global Const PHONEBUTTONFUNCTION_DATAOFF = &H19&
Global Const PHONEBUTTONFUNCTION_DONOTDISTURB = &H1A&
Global Const PHONEBUTTONFUNCTION_INTERCOM = &H1B&
Global Const PHONEBUTTONFUNCTION_BRIDGEDAPP = &H1C&
Global Const PHONEBUTTONFUNCTION_BUSY = &H1D&
Global Const PHONEBUTTONFUNCTION_CALLAPP = &H1E&
Global Const PHONEBUTTONFUNCTION_DATETIME = &H1F&
Global Const PHONEBUTTONFUNCTION_DIRECTORY = &H20&
Global Const PHONEBUTTONFUNCTION_COVER = &H21&
Global Const PHONEBUTTONFUNCTION_CALLID = &H22&
Global Const PHONEBUTTONFUNCTION_LASTNUM = &H23&
Global Const PHONEBUTTONFUNCTION_NIGHTSRV = &H24&
Global Const PHONEBUTTONFUNCTION_SENDCALLS = &H25&
Global Const PHONEBUTTONFUNCTION_MSGINDICATOR = &H26&
Global Const PHONEBUTTONFUNCTION_REPDIAL = &H27&
Global Const PHONEBUTTONFUNCTION_SETREPDIAL = &H28&
Global Const PHONEBUTTONFUNCTION_SYSTEMSPEED = &H29&
Global Const PHONEBUTTONFUNCTION_STATIONSPEED = &H2A&
Global Const PHONEBUTTONFUNCTION_CAMPON = &H2B&
Global Const PHONEBUTTONFUNCTION_SAVEREPEAT = &H2C&
Global Const PHONEBUTTONFUNCTION_QUEUECALL = &H2D&
Global Const PHONEBUTTONFUNCTION_NONE = &H2E&

Type PHONEBUTTONINFO
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwButtonMode As Long
    dwButtonFunction As Long

    dwButtonTextSize As Long
    dwButtonTextOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
End Type
Global Const PHONEBUTTONINFO_FIXEDSIZE = 36

' Note: the "_STR" Types are used to convert from data returned in variable-length strings
' to the fixed structure using LSET

Type PHONEBUTTONINFO_STR
    mem As String * PHONEBUTTONINFO_FIXEDSIZE
End Type

Global Const PHONEBUTTONMODE_DUMMY = &H1&
Global Const PHONEBUTTONMODE_CALL = &H2&
Global Const PHONEBUTTONMODE_FEATURE = &H4&
Global Const PHONEBUTTONMODE_KEYPAD = &H8&
Global Const PHONEBUTTONMODE_LOCAL = &H10&
Global Const PHONEBUTTONMODE_DISPLAY = &H20&

Global Const PHONEBUTTONSTATE_UP = &H1&
Global Const PHONEBUTTONSTATE_DOWN = &H2&

Type PHONEEXTENSIONID
    dwExtensionID0 As Long
    dwExtensionID1 As Long
    dwExtensionID2 As Long
    dwExtensionID3 As Long
End Type
Global Const PHONEEXTENSIONID_FIXEDSIZE = 16

Type PHONEEXTENSIONID_STR
    mem As String * PHONEEXTENSIONID_FIXEDSIZE
End Type

Type PHONECAPS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwProviderInfoSize As Long
    dwProviderInfoOffset As Long

    dwPhoneInfoSize As Long
    dwPhoneInfoOffset As Long

    dwPermanentPhoneID As Long
    dwPhoneNameSize As Long
    dwPhoneNameOffset As Long
    dwStringFormat As Long

    dwPhoneStates As Long
    dwHookSwitchDevs As Long
    dwHandsetHookSwitchModes As Long
    dwSpeakerHookSwitchModes As Long
    dwHeadsetHookSwitchModes As Long

    dwVolumeFlags As Long
    dwGainFlags As Long
    dwDisplayNumRows As Long
    dwDisplayNumColumns As Long
    dwNumRingModes As Long
    dwNumButtonLamps As Long

    dwButtonModesSize As Long
    dwButtonModesOffset As Long

    dwButtonFunctionsSize As Long
    dwButtonFunctionsOffset As Long

    dwLampModesSize As Long
    dwLampModesOffset As Long

    dwNumSetData As Long
    dwSetDataSize As Long
    dwSetDataOffset As Long

    dwNumGetData As Long
    dwGetDataSize As Long
    dwGetDataOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
End Type
Global Const PHONECAPS_FIXEDSIZE = 144

Type PHONECAPS_STR
    mem As String * PHONECAPS_FIXEDSIZE
End Type

Global Const PHONEERR_ALLOCATED = &H90000001
Global Const PHONEERR_BADDEVICEID = &H90000002
Global Const PHONEERR_INCOMPATIBLEAPIVERSION = &H90000003
Global Const PHONEERR_INCOMPATIBLEEXTVERSION = &H90000004
Global Const PHONEERR_INIFILECORRUPT = &H90000005
Global Const PHONEERR_INUSE = &H90000006
Global Const PHONEERR_INVALAPPHANDLE = &H90000007
Global Const PHONEERR_INVALAPPNAME = &H90000008
Global Const PHONEERR_INVALBUTTONLAMPID = &H90000009
Global Const PHONEERR_INVALBUTTONMODE = &H9000000A
Global Const PHONEERR_INVALBUTTONSTATE = &H9000000B
Global Const PHONEERR_INVALDATAID = &H9000000C
Global Const PHONEERR_INVALDEVICECLASS = &H9000000D
Global Const PHONEERR_INVALEXTVERSION = &H9000000E
Global Const PHONEERR_INVALHOOKSWITCHDEV = &H9000000F
Global Const PHONEERR_INVALHOOKSWITCHMODE = &H90000010
Global Const PHONEERR_INVALLAMPMODE = &H90000011
Global Const PHONEERR_INVALPARAM = &H90000012
Global Const PHONEERR_INVALPHONEHANDLE = &H90000013
Global Const PHONEERR_INVALPHONESTATE = &H90000014
Global Const PHONEERR_INVALPOINTER = &H90000015
Global Const PHONEERR_INVALPRIVILEGE = &H90000016
Global Const PHONEERR_INVALRINGMODE = &H90000017
Global Const PHONEERR_NODEVICE = &H90000018
Global Const PHONEERR_NODRIVER = &H90000019
Global Const PHONEERR_NOMEM = &H9000001A
Global Const PHONEERR_NOTOWNER = &H9000001B
Global Const PHONEERR_OPERATIONFAILED = &H9000001C
Global Const PHONEERR_OPERATIONUNAVAIL = &H9000001D
Global Const PHONEERR_RESOURCEUNAVAIL = &H9000001F
Global Const PHONEERR_REQUESTOVERRUN = &H90000020
Global Const PHONEERR_STRUCTURETOOSMALL = &H90000021
Global Const PHONEERR_UNINITIALIZED = &H90000022
Global Const PHONEERR_REINIT = &H90000023

Global Const PHONEHOOKSWITCHDEV_HANDSET = &H1&
Global Const PHONEHOOKSWITCHDEV_SPEAKER = &H2&
Global Const PHONEHOOKSWITCHDEV_HEADSET = &H4&

Global Const PHONEHOOKSWITCHMODE_ONHOOK = &H1&
Global Const PHONEHOOKSWITCHMODE_MIC = &H2&
Global Const PHONEHOOKSWITCHMODE_SPEAKER = &H4&
Global Const PHONEHOOKSWITCHMODE_MICSPEAKER = &H8&
Global Const PHONEHOOKSWITCHMODE_UNKNOWN = &H10&

Global Const PHONELAMPMODE_DUMMY = &H1&
Global Const PHONELAMPMODE_OFF = &H2&
Global Const PHONELAMPMODE_STEADY = &H4&
Global Const PHONELAMPMODE_WINK = &H8&
Global Const PHONELAMPMODE_FLASH = &H10&
Global Const PHONELAMPMODE_FLUTTER = &H20&
Global Const PHONELAMPMODE_BROKENFLUTTER = &H40&
Global Const PHONELAMPMODE_UNKNOWN = &H80&

Global Const PHONEPRIVILEGE_MONITOR = &H1&
Global Const PHONEPRIVILEGE_OWNER = &H2&

Global Const PHONESTATE_OTHER = &H1&
Global Const PHONESTATE_CONNECTED = &H2&
Global Const PHONESTATE_DISCONNECTED = &H4&
Global Const PHONESTATE_OWNER = &H8&
Global Const PHONESTATE_MONITORS = &H10&
Global Const PHONESTATE_DISPLAY = &H20&
Global Const PHONESTATE_LAMP = &H40&
Global Const PHONESTATE_RINGMODE = &H80&
Global Const PHONESTATE_RINGVOLUME = &H100&
Global Const PHONESTATE_HANDSETHOOKSWITCH = &H200&
Global Const PHONESTATE_HANDSETVOLUME = &H400&
Global Const PHONESTATE_HANDSETGAIN = &H800&
Global Const PHONESTATE_SPEAKERHOOKSWITCH = &H1000&
Global Const PHONESTATE_SPEAKERVOLUME = &H2000&
Global Const PHONESTATE_SPEAKERGAIN = &H4000&
Global Const PHONESTATE_HEADSETHOOKSWITCH = &H8000&
Global Const PHONESTATE_HEADSETVOLUME = &H10000
Global Const PHONESTATE_HEADSETGAIN = &H20000
Global Const PHONESTATE_SUSPEND = &H40000
Global Const PHONESTATE_RESUME = &H80000
Global Const PHONESTATE_DEVSPECIFIC = &H100000
Global Const PHONESTATE_REINIT = &H200000

Type PHONESTATUS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwStatusFlags As Long
    dwNumOwners As Long
    dwNumMonitors As Long
    dwRingMode As Long
    dwRingVolume As Long

    dwHandsetHookSwitchMode As Long
    dwHandsetVolume As Long
    dwHandsetGain As Long

    dwSpeakerHookSwitchMode As Long
    dwSpeakerVolume As Long
    dwSpeakerGain As Long

    dwHeadsetHookSwitchMode As Long
    dwHeadsetVolume As Long
    dwHeadsetGain As Long

    dwDisplaySize As Long
    dwDisplayOffset As Long

    dwLampModesSize As Long
    dwLampModesOffset As Long

    dwOwnerNameSize As Long
    dwOwnerNameOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
End Type
Global Const PHONESTATUS_FIXEDSIZE = 100

Type PHONESTATUS_STR
    mem As String * PHONESTATUS_FIXEDSIZE
End Type

Global Const PHONESTATUSFLAGS_CONNECTED = &H1&
Global Const PHONESTATUSFLAGS_SUSPENDED = &H2&

Global Const STRINGFORMAT_ASCII = &H1&
Global Const STRINGFORMAT_DBCS = &H2&
Global Const STRINGFORMAT_UNICODE = &H3&
Global Const STRINGFORMAT_BINARY = &H4&

Type VARSTRING
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwStringFormat As Long
    dwStringSize As Long
    dwStringOffset As Long
End Type
Global Const VARSTRING_FIXEDSIZE = 24

Type VARSTRING_STR
    mem As String * VARSTRING_FIXEDSIZE
End Type

' Data types and values for Lines

Global Const LINEADDRCAPFLAGS_FWDNUMRINGS = &H1&
Global Const LINEADDRCAPFLAGS_PICKUPGROUPID = &H2&
Global Const LINEADDRCAPFLAGS_SECURE = &H4&
Global Const LINEADDRCAPFLAGS_BLOCKIDDEFAULT = &H8&
Global Const LINEADDRCAPFLAGS_BLOCKIDOVERRIDE = &H10&
Global Const LINEADDRCAPFLAGS_DIALED = &H20&
Global Const LINEADDRCAPFLAGS_ORIGOFFHOOK = &H40&
Global Const LINEADDRCAPFLAGS_DESTOFFHOOK = &H80&
Global Const LINEADDRCAPFLAGS_FWDCONSULT = &H100&
Global Const LINEADDRCAPFLAGS_SETUPCONFNULL = &H200&
Global Const LINEADDRCAPFLAGS_AUTORECONNECT = &H400&
Global Const LINEADDRCAPFLAGS_COMPLETIONID = &H800&
Global Const LINEADDRCAPFLAGS_TRANSFERHELD = &H1000&
Global Const LINEADDRCAPFLAGS_TRANSFERMAKE = &H2000&
Global Const LINEADDRCAPFLAGS_CONFERENCEHELD = &H4000&
Global Const LINEADDRCAPFLAGS_CONFERENCEMAKE = &H8000&
Global Const LINEADDRCAPFLAGS_PARTIALDIAL = &H10000
Global Const LINEADDRCAPFLAGS_FWDSTATUSVALID = &H20000
Global Const LINEADDRCAPFLAGS_FWDINTEXTADDR = &H40000
Global Const LINEADDRCAPFLAGS_FWDBUSYNAADDR = &H80000
Global Const LINEADDRCAPFLAGS_ACCEPTTOALERT = &H100000
Global Const LINEADDRCAPFLAGS_CONFDROP = &H200000
Global Const LINEADDRCAPFLAGS_PICKUPCALLWAIT = &H400000

Type LINEADDRESSCAPS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwLineDeviceID As Long

    dwAddressSize As Long
    dwAddressOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long

    dwAddressSharing As Long
    dwAddressStates As Long
    dwCallInfoStates As Long
    dwCallerIDFlags As Long
    dwCalledIDFlags As Long
    dwConnectedIDFlags As Long
    dwRedirectionIDFlags As Long
    dwRedirectingIDFlags As Long
    dwCallStates As Long
    dwDialToneModes As Long
    dwBusyModes As Long
    dwSpecialInfo As Long
    dwDisconnectModes As Long

    dwMaxNumActiveCalls As Long
    dwMaxNumOnHoldCalls As Long
    dwMaxNumOnHoldPendingCalls As Long
    dwMaxNumConference As Long
    dwMaxNumTransConf As Long

    dwAddrCapFlags As Long
    dwCallFeatures As Long
    dwRemoveFromConfCaps As Long
    dwRemoveFromConfState As Long
    dwTransferModes As Long
    dwParkModes As Long

    dwForwardModes As Long
    dwMaxForwardEntries As Long
    dwMaxSpecificEntries As Long
    dwMinFwdNumRings As Long
    dwMaxFwdNumRings As Long

    dwMaxCallCompletions As Long
    dwCallCompletionConds As Long
    dwCallCompletionModes As Long
    dwNumCompletionMessages As Long
    dwCompletionMsgTextEntrySize As Long
    dwCompletionMsgTextSize As Long
    dwCompletionMsgTextOffset As Long
End Type
Global Const LINEADDRESSCAPS_FIXEDSIZE = 176

Type LINEADDRESSCAPS_STR
    mem As String * LINEADDRESSCAPS_FIXEDSIZE
End Type

Global Const LINEADDRESSMODE_ADDRESSID = &H1&
Global Const LINEADDRESSMODE_DIALABLEADDR = &H2&

Global Const LINEADDRESSSHARING_PRIVATE = &H1&
Global Const LINEADDRESSSHARING_BRIDGEDEXCL = &H2&
Global Const LINEADDRESSSHARING_BRIDGEDNEW = &H4&
Global Const LINEADDRESSSHARING_BRIDGEDSHARED = &H8&
Global Const LINEADDRESSSHARING_MONITORED = &H10&

Global Const LINEADDRESSSTATE_OTHER = &H1&
Global Const LINEADDRESSSTATE_DEVSPECIFIC = &H2&
Global Const LINEADDRESSSTATE_INUSEZERO = &H4&
Global Const LINEADDRESSSTATE_INUSEONE = &H8&
Global Const LINEADDRESSSTATE_INUSEMANY = &H10&
Global Const LINEADDRESSSTATE_NUMCALLS = &H20&
Global Const LINEADDRESSSTATE_FORWARD = &H40&
Global Const LINEADDRESSSTATE_TERMINALS = &H80&

Type LINEADDRESSSTATUS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwNumInUse As Long
    dwNumActiveCalls As Long
    dwNumOnHoldCalls As Long
    dwNumOnHoldPendCalls As Long
    dwAddressFeatures As Long

    dwNumRingsNoAnswer As Long
    dwForwardNumEntries As Long
    dwForwardSize As Long
    dwForwardOffset As Long

    dwTerminalModesSize As Long
    dwTerminalModesOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
End Type
Global Const LINEADDRESSSTATUS_FIXEDSIZE = 64

Type LINEADDRESSSTATUS_STR
    mem As String * LINEADDRESSSTATUS_FIXEDSIZE
End Type

Global Const LINEADDRFEATURE_FORWARD = &H1&
Global Const LINEADDRFEATURE_MAKECALL = &H2&
Global Const LINEADDRFEATURE_PICKUP = &H4&
Global Const LINEADDRFEATURE_SETMEDIACONTROL = &H8&
Global Const LINEADDRFEATURE_SETTERMINAL = &H10&
Global Const LINEADDRFEATURE_SETUPCONF = &H20&
Global Const LINEADDRFEATURE_UNCOMPLETECALL = &H40&
Global Const LINEADDRFEATURE_UNPARK = &H80&

Global Const LINEANSWERMODE_NONE = &H1&
Global Const LINEANSWERMODE_DROP = &H2&
Global Const LINEANSWERMODE_HOLD = &H4&

Global Const LINEBEARERMODE_VOICE = &H1&
Global Const LINEBEARERMODE_SPEECH = &H2&
Global Const LINEBEARERMODE_MULTIUSE = &H4&
Global Const LINEBEARERMODE_DATA = &H8&
Global Const LINEBEARERMODE_ALTSPEECHDATA = &H10&
Global Const LINEBEARERMODE_NONCALLSIGNALING = &H20&

Global Const LINEBUSYMODE_STATION = &H1&
Global Const LINEBUSYMODE_TRUNK = &H2&
Global Const LINEBUSYMODE_UNKNOWN = &H4&
Global Const LINEBUSYMODE_UNAVAIL = &H8&

Global Const LINECALLCOMPLCOND_BUSY = &H1&
Global Const LINECALLCOMPLCOND_NOANSWER = &H2&

Global Const LINECALLCOMPLMODE_CAMPON = &H1&
Global Const LINECALLCOMPLMODE_CALLBACK = &H2&
Global Const LINECALLCOMPLMODE_INTRUDE = &H4&
Global Const LINECALLCOMPLMODE_MESSAGE = &H8&

Global Const LINECALLFEATURE_ACCEPT = &H1&
Global Const LINECALLFEATURE_ADDTOCONF = &H2&
Global Const LINECALLFEATURE_ANSWER = &H4&
Global Const LINECALLFEATURE_BLINDTRANSFER = &H8&
Global Const LINECALLFEATURE_COMPLETECALL = &H10&
Global Const LINECALLFEATURE_COMPLETETRANSF = &H20&
Global Const LINECALLFEATURE_DIAL = &H40&
Global Const LINECALLFEATURE_DROP = &H80&
Global Const LINECALLFEATURE_GATHERDIGITS = &H100&
Global Const LINECALLFEATURE_GENERATEDIGITS = &H200&
Global Const LINECALLFEATURE_GENERATETONE = &H400&
Global Const LINECALLFEATURE_HOLD = &H800&
Global Const LINECALLFEATURE_MONITORDIGITS = &H1000&
Global Const LINECALLFEATURE_MONITORMEDIA = &H2000&
Global Const LINECALLFEATURE_MONITORTONES = &H4000&
Global Const LINECALLFEATURE_PARK = &H8000&
Global Const LINECALLFEATURE_PREPAREADDCONF = &H10000
Global Const LINECALLFEATURE_REDIRECT = &H20000
Global Const LINECALLFEATURE_REMOVEFROMCONF = &H40000
Global Const LINECALLFEATURE_SECURECALL = &H80000
Global Const LINECALLFEATURE_SENDUSERUSER = &H100000
Global Const LINECALLFEATURE_SETCALLPARAMS = &H200000
Global Const LINECALLFEATURE_SETMEDIACONTROL = &H400000
Global Const LINECALLFEATURE_SETTERMINAL = &H800000
Global Const LINECALLFEATURE_SETUPCONF = &H1000000
Global Const LINECALLFEATURE_SETUPTRANSFER = &H2000000
Global Const LINECALLFEATURE_SWAPHOLD = &H4000000
Global Const LINECALLFEATURE_UNHOLD = &H8000000

Type LINEDIALPARAMS
    dwDialPause As Long
    dwDialSpeed As Long
    dwDigitDuration As Long
    dwWaitForDialtone As Long
End Type
Global Const LINEDIALPARAMS_FIXEDSIZE = 16

Type LINEDIALPARAMS_STR
    mem As String * LINEDIALPARAMS_FIXEDSIZE
End Type

Type LINECALLINFO
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    hLine As Long
    dwLineDeviceID As Long
    dwAddressID As Long

    dwBearerMode As Long
    dwRate As Long
    dwMediaMode As Long

    dwAppSpecific As Long
    dwCallID As Long
    dwRelatedCallID As Long
    dwCallParamFlags As Long
    dwCallStates As Long

    dwMonitorDigitModes As Long
    dwMonitorMediaModes As Long
    DialParams As LINEDIALPARAMS

    dwOrigin As Long
    dwReason As Long
    dwCompletionID As Long
    dwNumOwners As Long
    dwNumMonitors As Long

    dwCountryCode As Long
    dwTrunk As Long

    dwCallerIDFlags As Long
    dwCallerIDSize As Long
    dwCallerIDOffset As Long
    dwCallerIDNameSize As Long
    dwCallerIDNameOffset As Long

    dwCalledIDFlags As Long
    dwCalledIDSize As Long
    dwCalledIDOffset As Long
    dwCalledIDNameSize As Long
    dwCalledIDNameOffset As Long

    dwConnectedIDFlags As Long
    dwConnectedIDSize As Long
    dwConnectedIDOffset As Long
    dwConnectedIDNameSize As Long
    dwConnectedIDNameOffset As Long

    dwRedirectionIDFlags As Long
    dwRedirectionIDSize As Long
    dwRedirectionIDOffset As Long
    dwRedirectionIDNameSize As Long
    dwRedirectionIDNameOffset As Long

    dwRedirectingIDFlags As Long
    dwRedirectingIDSize As Long
    dwRedirectingIDOffset As Long
    dwRedirectingIDNameSize As Long
    dwRedirectingIDNameOffset As Long

    dwAppNameSize As Long
    dwAppNameOffset As Long

    dwDisplayableAddressSize As Long
    dwDisplayableAddressOffset As Long

    dwCalledPartySize As Long
    dwCalledPartyOffset As Long

    dwCommentSize As Long
    dwCommentOffset As Long

    dwDisplaySize As Long
    dwDisplayOffset As Long

    dwUserUserInfoSize As Long
    dwUserUserInfoOffset As Long

    dwHighLevelCompSize As Long
    dwHighLevelCompOffset As Long

    dwLowLevelCompSize As Long
    dwLowLevelCompOffset As Long

    dwChargingInfoSize As Long
    dwChargingInfoOffset As Long

    dwTerminalModesSize As Long
    dwTerminalModesOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
End Type
Global Const LINECALLINFO_FIXEDSIZE = 296

Type LINECALLINFO_STR
    mem As String * LINECALLINFO_FIXEDSIZE
End Type

Global Const LINECALLINFOSTATE_OTHER = &H1&
Global Const LINECALLINFOSTATE_DEVSPECIFIC = &H2&
Global Const LINECALLINFOSTATE_BEARERMODE = &H4&
Global Const LINECALLINFOSTATE_RATE = &H8&
Global Const LINECALLINFOSTATE_MEDIAMODE = &H10&
Global Const LINECALLINFOSTATE_APPSPECIFIC = &H20&
Global Const LINECALLINFOSTATE_CALLID = &H40&
Global Const LINECALLINFOSTATE_RELATEDCALLID = &H80&
Global Const LINECALLINFOSTATE_ORIGIN = &H100&
Global Const LINECALLINFOSTATE_REASON = &H200&
Global Const LINECALLINFOSTATE_COMPLETIONID = &H400&
Global Const LINECALLINFOSTATE_NUMOWNERINCR = &H800&
Global Const LINECALLINFOSTATE_NUMOWNERDECR = &H1000&
Global Const LINECALLINFOSTATE_NUMMONITORS = &H2000&
Global Const LINECALLINFOSTATE_TRUNK = &H4000&
Global Const LINECALLINFOSTATE_CALLERID = &H8000&
Global Const LINECALLINFOSTATE_CALLEDID = &H10000
Global Const LINECALLINFOSTATE_CONNECTEDID = &H20000
Global Const LINECALLINFOSTATE_REDIRECTIONID = &H40000
Global Const LINECALLINFOSTATE_REDIRECTINGID = &H80000
Global Const LINECALLINFOSTATE_DISPLAY = &H100000
Global Const LINECALLINFOSTATE_USERUSERINFO = &H200000
Global Const LINECALLINFOSTATE_HIGHLEVELCOMP = &H400000
Global Const LINECALLINFOSTATE_LOWLEVELCOMP = &H800000
Global Const LINECALLINFOSTATE_CHARGINGINFO = &H1000000
Global Const LINECALLINFOSTATE_TERMINAL = &H2000000
Global Const LINECALLINFOSTATE_DIALPARAMS = &H4000000
Global Const LINECALLINFOSTATE_MONITORMODES = &H8000000

Type LINECALLLIST
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwCallsNumEntries As Long
    dwCallsSize As Long
    dwCallsOffset As Long
End Type
Global Const LINECALLLIST_FIXEDSIZE = 24

Type LINECALLLIST_STR
    mem As String * LINECALLLIST_FIXEDSIZE
End Type

Global Const LINECALLORIGIN_OUTBOUND = &H1&
Global Const LINECALLORIGIN_INTERNAL = &H2&
Global Const LINECALLORIGIN_EXTERNAL = &H4&
Global Const LINECALLORIGIN_UNKNOWN = &H10&
Global Const LINECALLORIGIN_UNAVAIL = &H20&
Global Const LINECALLORIGIN_CONFERENCE = &H40&

Global Const LINECALLPARAMFLAGS_SECURE = &H1&
Global Const LINECALLPARAMFLAGS_IDLE = &H2&
Global Const LINECALLPARAMFLAGS_BLOCKID = &H4&
Global Const LINECALLPARAMFLAGS_ORIGOFFHOOK = &H8&
Global Const LINECALLPARAMFLAGS_DESTOFFHOOK = &H10&

Type LINECALLPARAMS
    dwTotalSize As Long

    dwBearerMode As Long
    dwMinRate As Long
    dwMaxRate As Long
    dwMediaMode As Long

    dwCallParamFlags As Long
    dwAddressMode As Long
    dwAddressID As Long

    DialParams As LINEDIALPARAMS

    dwOrigAddressSize As Long
    dwOrigAddressOffset As Long

    dwDisplayableAddressSize As Long
    dwDisplayableAddressOffset As Long

    dwCalledPartySize As Long
    dwCalledPartyOffset As Long

    dwCommentSize As Long
    dwCommentOffset As Long

    dwUserUserInfoSize As Long
    dwUserUserInfoOffset As Long

    dwHighLevelCompSize As Long
    dwHighLevelCompOffset As Long

    dwLowLevelCompSize As Long
    dwLowLevelCompOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
    '
    mem As String * 2048 ' added by mca
End Type
Global Const LINECALLPARAMS_FIXEDSIZE = 112

Type LINECALLPARAMS_STR
    mem As String * LINECALLPARAMS_FIXEDSIZE
End Type

Global Const LINECALLPARTYID_BLOCKED = &H1&
Global Const LINECALLPARTYID_OUTOFAREA = &H2&
Global Const LINECALLPARTYID_NAME = &H4&
Global Const LINECALLPARTYID_ADDRESS = &H8&
Global Const LINECALLPARTYID_PARTIAL = &H10&
Global Const LINECALLPARTYID_UNKNOWN = &H20&
Global Const LINECALLPARTYID_UNAVAIL = &H40&

Global Const LINECALLPRIVILEGE_NONE = &H1&
Global Const LINECALLPRIVILEGE_MONITOR = &H2&
Global Const LINECALLPRIVILEGE_OWNER = &H4&

Global Const LINECALLREASON_DIRECT = &H1&
Global Const LINECALLREASON_FWDBUSY = &H2&
Global Const LINECALLREASON_FWDNOANSWER = &H4&
Global Const LINECALLREASON_FWDUNCOND = &H8&
Global Const LINECALLREASON_PICKUP = &H10&
Global Const LINECALLREASON_UNPARK = &H20&
Global Const LINECALLREASON_REDIRECT = &H40&
Global Const LINECALLREASON_CALLCOMPLETION = &H80&
Global Const LINECALLREASON_TRANSFER = &H100&
Global Const LINECALLREASON_REMINDER = &H200&
Global Const LINECALLREASON_UNKNOWN = &H400&
Global Const LINECALLREASON_UNAVAIL = &H800&

Global Const LINECALLSELECT_LINE = &H1&
Global Const LINECALLSELECT_ADDRESS = &H2&
Global Const LINECALLSELECT_CALL = &H4&

Global Const LINECALLSTATE_IDLE = &H1&
Global Const LINECALLSTATE_OFFERING = &H2&
Global Const LINECALLSTATE_ACCEPTED = &H4&
Global Const LINECALLSTATE_DIALTONE = &H8&
Global Const LINECALLSTATE_DIALING = &H10&
Global Const LINECALLSTATE_RINGBACK = &H20&
Global Const LINECALLSTATE_BUSY = &H40&
Global Const LINECALLSTATE_SPECIALINFO = &H80&
Global Const LINECALLSTATE_CONNECTED = &H100&
Global Const LINECALLSTATE_PROCEEDING = &H200&
Global Const LINECALLSTATE_ONHOLD = &H400&
Global Const LINECALLSTATE_CONFERENCED = &H800&
Global Const LINECALLSTATE_ONHOLDPENDCONF = &H1000&
Global Const LINECALLSTATE_ONHOLDPENDTRANSFER = &H2000&
Global Const LINECALLSTATE_DISCONNECTED = &H4000&
Global Const LINECALLSTATE_UNKNOWN = &H8000&

Type LINECALLSTATUS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwCallState As Long
    dwCallStateMode As Long
    dwCallPrivilege As Long
    dwCallFeatures As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
End Type
Global Const LINECALLSTATUS_FIXEDSIZE = 36

Type LINECALLSTATUS_STR
    mem As String * LINECALLSTATUS_FIXEDSIZE
End Type

Global Const LINEDEVCAPFLAGS_CROSSADDRCONF = &H1&
Global Const LINEDEVCAPFLAGS_HIGHLEVCOMP = &H2&
Global Const LINEDEVCAPFLAGS_LOWLEVCOMP = &H4&
Global Const LINEDEVCAPFLAGS_MEDIACONTROL = &H8&
Global Const LINEDEVCAPFLAGS_MULTIPLEADDR = &H10&
Global Const LINEDEVCAPFLAGS_CLOSEDROP = &H20&
Global Const LINEDEVCAPFLAGS_DIALBILLING = &H40&
Global Const LINEDEVCAPFLAGS_DIALQUIET = &H80&
Global Const LINEDEVCAPFLAGS_DIALDIALTONE = &H100&

Type LINEEXTENSIONID
    dwExtensionID0 As Long
    dwExtensionID1 As Long
    dwExtensionID2 As Long
    dwExtensionID3 As Long
End Type
Global Const LINEEXTENSIONID_FIXEDSIZE = 16

Type LINEEXTENSIONID_STR
    mem As String * LINEEXTENSIONID_FIXEDSIZE
End Type

Type LINEDEVCAPS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwProviderInfoSize As Long
    dwProviderInfoOffset As Long

    dwSwitchInfoSize As Long
    dwSwitchInfoOffset As Long

    dwPermanentLineID As Long
    dwLineNameSize As Long
    dwLineNameOffset As Long
    dwStringFormat As Long

    dwAddressModes As Long
    dwNumAddresses As Long
    dwBearerModes As Long
    dwMaxRate As Long
    dwMediaModes As Long

    dwGenerateToneModes As Long
    dwGenerateToneMaxNumFreq As Long
    dwGenerateDigitModes As Long
    dwMonitorToneMaxNumFreq As Long
    dwMonitorToneMaxNumEntries As Long
    dwMonitorDigitModes As Long
    dwGatherDigitsMinTimeout As Long
    dwGatherDigitsMaxTimeout As Long

    dwMedCtlDigitMaxListSize As Long
    dwMedCtlMediaMaxListSize As Long
    dwMedCtlToneMaxListSize As Long
    dwMedCtlCallStateMaxListSize As Long

    dwDevCapFlags As Long
    dwMaxNumActiveCalls As Long
    dwAnswerMode As Long
    dwRingModes As Long
    dwLineStates As Long

    dwUUIAcceptSize As Long
    dwUUIAnswerSize As Long
    dwUUIMakeCallSize As Long
    dwUUIDropSize As Long
    dwUUISendUserUserInfoSize As Long
    dwUUICallInfoSize As Long

    MinDialParams As LINEDIALPARAMS
    MaxDialParams As LINEDIALPARAMS
    DefaultDialParams As LINEDIALPARAMS

    dwNumTerminals As Long
    dwTerminalCapsSize As Long
    dwTerminalCapsOffset As Long
    dwTerminalTextEntrySize As Long
    dwTerminalTextSize As Long
    dwTerminalTextOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
    
    mem As String * 2048 ' ADDED BY MCA
End Type
Global Const LINEDEVCAPS_FIXEDSIZE = 236

Type LINEDEVCAPS_STR
    mem As String * LINEDEVCAPS_FIXEDSIZE
End Type

Global Const LINEDEVSTATE_OTHER = &H1&
Global Const LINEDEVSTATE_RINGING = &H2&
Global Const LINEDEVSTATE_CONNECTED = &H4&
Global Const LINEDEVSTATE_DISCONNECTED = &H8&
Global Const LINEDEVSTATE_MSGWAITON = &H10&
Global Const LINEDEVSTATE_MSGWAITOFF = &H20&
Global Const LINEDEVSTATE_INSERVICE = &H40&
Global Const LINEDEVSTATE_OUTOFSERVICE = &H80&
Global Const LINEDEVSTATE_MAINTENANCE = &H100&
Global Const LINEDEVSTATE_OPEN = &H200&
Global Const LINEDEVSTATE_CLOSE = &H400&
Global Const LINEDEVSTATE_NUMCALLS = &H800&
Global Const LINEDEVSTATE_NUMCOMPLETIONS = &H1000&
Global Const LINEDEVSTATE_TERMINALS = &H2000&
Global Const LINEDEVSTATE_ROAMMODE = &H4000&
Global Const LINEDEVSTATE_BATTERY = &H8000&
Global Const LINEDEVSTATE_SIGNAL = &H10000
Global Const LINEDEVSTATE_DEVSPECIFIC = &H20000
Global Const LINEDEVSTATE_REINIT = &H40000
Global Const LINEDEVSTATE_LOCK = &H80000

Type LINEDEVSTATUS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwNumOpens As Long
    dwOpenMediaModes As Long
    dwNumActiveCalls As Long
    dwNumOnHoldCalls As Long
    dwNumOnHoldPendCalls As Long
    dwLineFeatures As Long
    dwNumCallCompletions As Long
    dwRingMode As Long
    dwSignalLevel As Long
    dwBatteryLevel As Long
    dwRoamMode As Long

    dwDevStatusFlags As Long

    dwTerminalModesSize As Long
    dwTerminalModesOffset As Long

    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
End Type
Global Const LINEDEVSTATUS_FIXEDSIZE = 76

Type LINEDEVSTATUS_STR
    mem As String * LINEDEVSTATUS_FIXEDSIZE
End Type

Global Const LINEDEVSTATUSFLAGS_CONNECTED = &H1&
Global Const LINEDEVSTATUSFLAGS_MSGWAIT = &H2&
Global Const LINEDEVSTATUSFLAGS_INSERVICE = &H4&
Global Const LINEDEVSTATUSFLAGS_LOCKED = &H8&

Global Const LINEDIALTONEMODE_NORMAL = &H1&
Global Const LINEDIALTONEMODE_SPECIAL = &H2&
Global Const LINEDIALTONEMODE_INTERNAL = &H4&
Global Const LINEDIALTONEMODE_EXTERNAL = &H8&
Global Const LINEDIALTONEMODE_UNKNOWN = &H10&
Global Const LINEDIALTONEMODE_UNAVAIL = &H20&

Global Const LINEDIGITMODE_PULSE = &H1&
Global Const LINEDIGITMODE_DTMF = &H2&
Global Const LINEDIGITMODE_DTMFEND = &H4&

Global Const LINEDISCONNECTMODE_NORMAL = &H1&
Global Const LINEDISCONNECTMODE_UNKNOWN = &H2&
Global Const LINEDISCONNECTMODE_REJECT = &H4&
Global Const LINEDISCONNECTMODE_PICKUP = &H8&
Global Const LINEDISCONNECTMODE_FORWARDED = &H10&
Global Const LINEDISCONNECTMODE_BUSY = &H20&
Global Const LINEDISCONNECTMODE_NOANSWER = &H40&
Global Const LINEDISCONNECTMODE_BADADDRESS = &H80&
Global Const LINEDISCONNECTMODE_UNREACHABLE = &H100&
Global Const LINEDISCONNECTMODE_CONGESTION = &H200&
Global Const LINEDISCONNECTMODE_INCOMPATIBLE = &H400&
Global Const LINEDISCONNECTMODE_UNAVAIL = &H800&

Global Const LINEERR_ALLOCATED = &H80000001
Global Const LINEERR_BADDEVICEID = &H80000002
Global Const LINEERR_BEARERMODEUNAVAIL = &H80000003
Global Const LINEERR_CALLUNAVAIL = &H80000005
Global Const LINEERR_COMPLETIONOVERRUN = &H80000006
Global Const LINEERR_CONFERENCEFULL = &H80000007
Global Const LINEERR_DIALBILLING = &H80000008
Global Const LINEERR_DIALDIALTONE = &H80000009
Global Const LINEERR_DIALPROMPT = &H8000000A
Global Const LINEERR_DIALQUIET = &H8000000B
Global Const LINEERR_INCOMPATIBLEAPIVERSION = &H8000000C
Global Const LINEERR_INCOMPATIBLEEXTVERSION = &H8000000D
Global Const LINEERR_INIFILECORRUPT = &H8000000E
Global Const LINEERR_INUSE = &H8000000F
Global Const LINEERR_INVALADDRESS = &H80000010
Global Const LINEERR_INVALADDRESSID = &H80000011
Global Const LINEERR_INVALADDRESSMODE = &H80000012
Global Const LINEERR_INVALADDRESSSTATE = &H80000013
Global Const LINEERR_INVALAPPHANDLE = &H80000014
Global Const LINEERR_INVALAPPNAME = &H80000015
Global Const LINEERR_INVALBEARERMODE = &H80000016
Global Const LINEERR_INVALCALLCOMPLMODE = &H80000017
Global Const LINEERR_INVALCALLHANDLE = &H80000018
Global Const LINEERR_INVALCALLPARAMS = &H80000019
Global Const LINEERR_INVALCALLPRIVILEGE = &H8000001A
Global Const LINEERR_INVALCALLSELECT = &H8000001B
Global Const LINEERR_INVALCALLSTATE = &H8000001C
Global Const LINEERR_INVALCALLSTATELIST = &H8000001D
Global Const LINEERR_INVALCARD = &H8000001E
Global Const LINEERR_INVALCOMPLETIONID = &H8000001F
Global Const LINEERR_INVALCONFCALLHANDLE = &H80000020
Global Const LINEERR_INVALCONSULTCALLHANDLE = &H80000021
Global Const LINEERR_INVALCOUNTRYCODE = &H80000022
Global Const LINEERR_INVALDEVICECLASS = &H80000023
Global Const LINEERR_INVALDEVICEHANDLE = &H80000024
Global Const LINEERR_INVALDIGITLIST = &H80000026
Global Const LINEERR_INVALDIGITMODE = &H80000027
Global Const LINEERR_INVALDIGITS = &H80000028
Global Const LINEERR_INVALEXTVERSION = &H80000029
Global Const LINEERR_INVALGROUPID = &H8000002A
Global Const LINEERR_INVALLINEHANDLE = &H8000002B
Global Const LINEERR_INVALLINESTATE = &H8000002C
Global Const LINEERR_INVALLOCATION = &H8000002D
Global Const LINEERR_INVALMEDIALIST = &H8000002E
Global Const LINEERR_INVALMEDIAMODE = &H8000002F
Global Const LINEERR_INVALMESSAGEID = &H80000030
Global Const LINEERR_INVALPARAM = &H80000032
Global Const LINEERR_INVALPARKID = &H80000033
Global Const LINEERR_INVALPARKMODE = &H80000034
Global Const LINEERR_INVALPOINTER = &H80000035
Global Const LINEERR_INVALPRIVSELECT = &H80000036
Global Const LINEERR_INVALRATE = &H80000037
Global Const LINEERR_INVALREQUESTMODE = &H80000038
Global Const LINEERR_INVALTERMINALID = &H80000039
Global Const LINEERR_INVALTERMINALMODE = &H8000003A
Global Const LINEERR_INVALTIMEOUT = &H8000003B
Global Const LINEERR_INVALTONE = &H8000003C
Global Const LINEERR_INVALTONELIST = &H8000003D
Global Const LINEERR_INVALTONEMODE = &H8000003E
Global Const LINEERR_INVALTRANSFERMODE = &H8000003F
Global Const LINEERR_LINEMAPPERFAILED = &H80000040
Global Const LINEERR_NOCONFERENCE = &H80000041
Global Const LINEERR_NODEVICE = &H80000042
Global Const LINEERR_NODRIVER = &H80000043
Global Const LINEERR_NOMEM = &H80000044
Global Const LINEERR_NOREQUEST = &H80000045
Global Const LINEERR_NOTOWNER = &H80000046
Global Const LINEERR_NOTREGISTERED = &H80000047
Global Const LINEERR_OPERATIONFAILED = &H80000048
Global Const LINEERR_OPERATIONUNAVAIL = &H80000049
Global Const LINEERR_RATEUNAVAIL = &H8000004A
Global Const LINEERR_RESOURCEUNAVAIL = &H8000004B
Global Const LINEERR_REQUESTOVERRUN = &H8000004C
Global Const LINEERR_STRUCTURETOOSMALL = &H8000004D
Global Const LINEERR_TARGETNOTFOUND = &H8000004E
Global Const LINEERR_TARGETSELF = &H8000004F
Global Const LINEERR_UNINITIALIZED = &H80000050
Global Const LINEERR_USERUSERINFOTOOBIG = &H80000051
Global Const LINEERR_REINIT = &H80000052
Global Const LINEERR_ADDRESSBLOCKED = &H80000053
Global Const LINEERR_BILLINGREJECTED = &H80000054
Global Const LINEERR_INVALFEATURE = &H80000055
Global Const LINEERR_NOMULTIPLEINSTANCE = &H80000056

Global Const LINEFEATURE_DEVSPECIFIC = &H1&
Global Const LINEFEATURE_DEVSPECIFICFEAT = &H2&
Global Const LINEFEATURE_FORWARD = &H4&
Global Const LINEFEATURE_MAKECALL = &H8&
Global Const LINEFEATURE_SETMEDIACONTROL = &H10&
Global Const LINEFEATURE_SETTERMINAL = &H20&

Type lineForward
    dwForwardMode As Long

    dwCallerAddressSize As Long
    dwCallerAddressOffset As Long

    dwDestCountryCode As Long
    dwDestAddressSize As Long
    dwDestAddressOffset As Long
End Type
Global Const LINEFORWARD_FIXEDSIZE = 24

Type LINEFORWARD_STR
    mem As String * LINEFORWARD_FIXEDSIZE
End Type

Type LINEFORWARDLIST
    dwTotalSize As Long
    dwNumEntries As Long
End Type
Global Const LINEFORWARDLIST_FIXEDSIZE = 8

Type LINEFORWARDLIST_STR
    mem As String * LINEFORWARDLIST_FIXEDSIZE
End Type

Global Const LINEFORWARDMODE_UNCOND = &H1&
Global Const LINEFORWARDMODE_UNCONDINTERNAL = &H2&
Global Const LINEFORWARDMODE_UNCONDEXTERNAL = &H4&
Global Const LINEFORWARDMODE_UNCONDSPECIFIC = &H8&
Global Const LINEFORWARDMODE_BUSY = &H10&
Global Const LINEFORWARDMODE_BUSYINTERNAL = &H20&
Global Const LINEFORWARDMODE_BUSYEXTERNAL = &H40&
Global Const LINEFORWARDMODE_BUSYSPECIFIC = &H80&
Global Const LINEFORWARDMODE_NOANSW = &H100&
Global Const LINEFORWARDMODE_NOANSWINTERNAL = &H200&
Global Const LINEFORWARDMODE_NOANSWEXTERNAL = &H400&
Global Const LINEFORWARDMODE_NOANSWSPECIFIC = &H800&
Global Const LINEFORWARDMODE_BUSYNA = &H1000&
Global Const LINEFORWARDMODE_BUSYNAINTERNAL = &H2000&
Global Const LINEFORWARDMODE_BUSYNAEXTERNAL = &H4000&
Global Const LINEFORWARDMODE_BUSYNASPECIFIC = &H8000&

Global Const LINEGATHERTERM_BUFFERFULL = &H1&
Global Const LINEGATHERTERM_TERMDIGIT = &H2&
Global Const LINEGATHERTERM_FIRSTTIMEOUT = &H4&
Global Const LINEGATHERTERM_INTERTIMEOUT = &H8&
Global Const LINEGATHERTERM_CANCEL = &H10&

Global Const LINEGENERATETERM_DONE = &H1&
Global Const LINEGENERATETERM_CANCEL = &H2&

' This type is named differently than in TAPI.H because of the conflict with the function of the same name

Type LINEGENERATETONE_TYPE
    dwFrequency As Long
    dwCadenceOn As Long
    dwCadenceOff As Long
    dwVolume As Long
End Type
Global Const LINEGENERATETONE_FIXEDSIZE = 16

Type LINEGENERATETONE_STR
    mem As String * LINEGENERATETONE_FIXEDSIZE
End Type

Global Const LINEMAPPER = &HFFFFFFFF

Type LINEMEDIACONTROLCALLSTATE
    dwCallStates As Long
    dwMediaControl As Long
End Type
Global Const LINEMEDIACONTROLCALLSTATE_FIXEDSIZE = 8

Type LINEMEDIACONTROLCALLSTATE_STR
    mem As String * LINEMEDIACONTROLCALLSTATE_FIXEDSIZE
End Type

Type LINEMEDIACONTROLDIGIT
    dwDigit As Long
    dwDigitModes As Long
    dwMediaControl As Long
End Type
Global Const LINEMEDIACONTROLDIGIT_FIXEDSIZE = 12

Type LINEMEDIACONTROLDIGIT_STR
    mem As String * LINEMEDIACONTROLDIGIT_FIXEDSIZE
End Type

Type LINEMEDIACONTROLMEDIA
    dwMediaModes As Long
    dwDuration As Long
    dwMediaControl As Long
End Type
Global Const LINEMEDIACONTROLMEDIA_FIXEDSIZE = 12

Type LINEMEDIACONTROLMEDIA_STR
    mem As String * LINEMEDIACONTROLMEDIA_FIXEDSIZE
End Type

Type LINEMEDIACONTROLTONE
    dwAppSpecific As Long
    dwDuration As Long
    dwFrequency1 As Long
    dwFrequency2 As Long
    dwFrequency3 As Long
    dwMediaControl As Long
End Type
Global Const LINEMEDIACONTROLTONE_FIXEDSIZE = 24

Type LINEMEDIACONTROLTONE_STR
    mem As String * LINEMEDIACONTROLTONE_FIXEDSIZE
End Type

Global Const LINEMEDIACONTROL_NONE = &H1&
Global Const LINEMEDIACONTROL_START = &H2&
Global Const LINEMEDIACONTROL_RESET = &H4&
Global Const LINEMEDIACONTROL_PAUSE = &H8&
Global Const LINEMEDIACONTROL_RESUME = &H10&
Global Const LINEMEDIACONTROL_RATEUP = &H20&
Global Const LINEMEDIACONTROL_RATEDOWN = &H40&
Global Const LINEMEDIACONTROL_RATENORMAL = &H80&
Global Const LINEMEDIACONTROL_VOLUMEUP = &H100&
Global Const LINEMEDIACONTROL_VOLUMEDOWN = &H200&
Global Const LINEMEDIACONTROL_VOLUMENORMAL = &H400&

Global Const LINEMEDIAMODE_UNKNOWN = &H2&
Global Const LINEMEDIAMODE_INTERACTIVEVOICE = &H4&
Global Const LINEMEDIAMODE_AUTOMATEDVOICE = &H8&
Global Const LINEMEDIAMODE_DATAMODEM = &H10&
Global Const LINEMEDIAMODE_G3FAX = &H20&
Global Const LINEMEDIAMODE_TDD = &H40&
Global Const LINEMEDIAMODE_G4FAX = &H80&
Global Const LINEMEDIAMODE_DIGITALDATA = &H100&
Global Const LINEMEDIAMODE_TELETEX = &H200&
Global Const LINEMEDIAMODE_VIDEOTEX = &H400&
Global Const LINEMEDIAMODE_TELEX = &H800&
Global Const LINEMEDIAMODE_MIXED = &H1000&
Global Const LINEMEDIAMODE_ADSI = &H2000&

Type LINEMONITORTONE
    dwAppSpecific As Long
    dwDuration As Long
    dwFrequency1 As Long
    dwFrequency2 As Long
    dwFrequency3 As Long
End Type
Global Const LINEMONITORTONE_FIXEDSIZE = 20

Type LINEMONITORTONE_STR
    mem As String * LINEMONITORTONE_FIXEDSIZE
End Type

Global Const LINEPARKMODE_DIRECTED = &H1&
Global Const LINEPARKMODE_NONDIRECTED = &H2&

Global Const LINEREMOVEFROMCONF_NONE = &H1&
Global Const LINEREMOVEFROMCONF_LAST = &H2&
Global Const LINEREMOVEFROMCONF_ANY = &H3&

Type LINEREQMAKECALL
    szDestAddress As String * TAPIMAXDESTADDRESSSIZE
    szAppName As String * TAPIMAXAPPNAMESIZE
    szCalledParty As String * TAPIMAXCALLEDPARTYSIZE
    szComment As String * TAPIMAXCOMMENTSIZE
End Type
Global Const LINEREQMAKECALL_FIXEDSIZE = TAPIMAXDESTADDRESSSIZE + TAPIMAXAPPNAMESIZE + TAPIMAXCALLEDPARTYSIZE + TAPIMAXCOMMENTSIZE

Type LINEREQMAKECALL_STR
    mem As String * LINEREQMAKECALL_FIXEDSIZE
End Type

Type LINEREQMEDIACALL
    hWnd As Integer
    wRequestID As Integer
    szDeviceClass As String * TAPIMAXDEVICECLASSSIZE
    ucDeviceID As String * TAPIMAXDEVICEIDSIZE
    dwSize As Long
    dwSecure As Long
    szDestAddress As String * TAPIMAXDESTADDRESSSIZE
    szAppName As String * TAPIMAXAPPNAMESIZE
    szCalledParty As String * TAPIMAXCALLEDPARTYSIZE
    szComment As String * TAPIMAXCOMMENTSIZE
End Type
Global Const LINEREQMEDIACALL_FIXEDSIZE = 12 + TAPIMAXDEVICECLASSSIZE + TAPIMAXDEVICEIDSIZE + TAPIMAXDESTADDRESSSIZE + TAPIMAXAPPNAMESIZE + TAPIMAXCALLEDPARTYSIZE + TAPIMAXCOMMENTSIZE

Type LINEREQMEDIACALL_STR
    mem As String * LINEREQMEDIACALL_FIXEDSIZE
End Type

Global Const LINEREQUESTMODE_MAKECALL = &H1&
Global Const LINEREQUESTMODE_MEDIACALL = &H2&
Global Const LINEREQUESTMODE_DROP = &H4&

Global Const LINEROAMMODE_UNKNOWN = &H1&
Global Const LINEROAMMODE_UNAVAIL = &H2&
Global Const LINEROAMMODE_HOME = &H4&
Global Const LINEROAMMODE_ROAMA = &H8&
Global Const LINEROAMMODE_ROAMB = &H10&

Global Const LINESPECIALINFO_NOCIRCUIT = &H1&
Global Const LINESPECIALINFO_CUSTIRREG = &H2&
Global Const LINESPECIALINFO_REORDER = &H4&
Global Const LINESPECIALINFO_UNKNOWN = &H8&
Global Const LINESPECIALINFO_UNAVAIL = &H10&

Type LINETERMCAPS
    dwTermDev As Long
    dwTermModes As Long
    dwTermSharing As Long
End Type
Global Const LINETERMCAPS_FIXEDSIZE = 12

Type LINETERMCAPS_STR
    mem As String * LINETERMCAPS_FIXEDSIZE
End Type

Global Const LINETERMDEV_PHONE = &H1&
Global Const LINETERMDEV_HEADSET = &H2&
Global Const LINETERMDEV_SPEAKER = &H4&

Global Const LINETERMMODE_BUTTONS = &H1&
Global Const LINETERMMODE_LAMPS = &H2&
Global Const LINETERMMODE_DISPLAY = &H4&
Global Const LINETERMMODE_RINGER = &H8&
Global Const LINETERMMODE_HOOKSWITCH = &H10&
Global Const LINETERMMODE_MEDIATOLINE = &H20&
Global Const LINETERMMODE_MEDIAFROMLINE = &H40&
Global Const LINETERMMODE_MEDIABIDIRECT = &H80&

Global Const LINETERMSHARING_PRIVATE = &H1&
Global Const LINETERMSHARING_SHAREDEXCL = &H2&
Global Const LINETERMSHARING_SHAREDCONF = &H4&

Global Const LINETONEMODE_CUSTOM = &H1&
Global Const LINETONEMODE_RINGBACK = &H2&
Global Const LINETONEMODE_BUSY = &H4&
Global Const LINETONEMODE_BEEP = &H8&
Global Const LINETONEMODE_BILLING = &H10&

Global Const LINETRANSFERMODE_TRANSFER = &H1&
Global Const LINETRANSFERMODE_CONFERENCE = &H2&

Type LINETRANSLATEOUTPUT
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwDialableStringSize As Long
    dwDialableStringOffset As Long
    dwDisplayableStringSize As Long
    dwDisplayableStringOffset As Long

    dwCurrentCountry As Long
    dwDestCountry As Long
    dwTranslateResults As Long
End Type
Global Const LINETRANSLATEOUTPUT_FIXEDSIZE = 40

Type LINETRANSLATEOUTPUT_STR
    mem As String * LINETRANSLATEOUTPUT_FIXEDSIZE
End Type

Type LINETRANSLATECAPS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long

    dwNumLocations As Long
    dwLocationListSize As Long
    dwLocationListOffset As Long

    dwCurrentLocationID As Long

    dwNumCards As Long
    dwCardListSize As Long
    dwCardListOffset As Long

    dwCurrentPreferredCardID As Long
End Type
Global Const LINETRANSLATECAPS_FIXEDSIZE = 44

Type LINETRANSLATECAPS_STR
    mem As String * LINETRANSLATECAPS_FIXEDSIZE
End Type

Type LINELOCATIONENTRY
    dwPermanentLocationID As Long
    dwLocationNameSize As Long
    dwLocationNameOffset As Long
    dwCountryCode As Long
    dwCityCodeSize As Long
    dwCityCodeOffset As Long
    dwPreferredCardID As Long
End Type
Global Const LINELOCATIONENTRY_FIXEDSIZE = 28

Type LINELOCATIONENTRY_STR
    mem As String * LINELOCATIONENTRY_FIXEDSIZE
End Type

Type LINECARDENTRY
    dwPermanentCardID As Long
    dwCardNameSize As Long
    dwCardNameOffset As Long
End Type
Global Const LINECARDENTRY_FIXEDSIZE = 12

Type LINECARDENTRY_STR
    mem As String * LINECARDENTRY_FIXEDSIZE
End Type

Global Const LINETOLLLISTOPTION_ADD = &H1&
Global Const LINETOLLLISTOPTION_REMOVE = &H2&

Global Const LINETRANSLATEOPTION_CARDOVERRIDE = &H1&

Global Const LINETRANSLATERESULT_CANONICAL = &H1&
Global Const LINETRANSLATERESULT_INTERNATIONAL = &H2&
Global Const LINETRANSLATERESULT_LONGDISTANCE = &H4&
Global Const LINETRANSLATERESULT_LOCAL = &H8&
Global Const LINETRANSLATERESULT_INTOLLLIST = &H10&
Global Const LINETRANSLATERESULT_NOTINTOLLLIST = &H20&
Global Const LINETRANSLATERESULT_DIALBILLING = &H40&
Global Const LINETRANSLATERESULT_DIALQUIET = &H80&
Global Const LINETRANSLATERESULT_DIALDIALTONE = &H100&
Global Const LINETRANSLATERESULT_DIALPROMPT = &H200&

' Simple Telephony prototypes

Declare Function tapiRequestMakeCall Lib "TAPI32.DLL" (ByVal lpszDestAddress As String, ByVal lpszAppName As String, ByVal lpszCalledParty As String, ByVal lpszComment As String) As Long

Declare Function tapiRequestMediaCall Lib "TAPI32.DLL" (ByVal hWnd As Integer, ByVal wRequestID As Integer, ByVal lpszDeviceClass As String, ByVal lpDeviceID As String, ByVal dwSize As Long, ByVal dwSecure As Long, ByVal lpszDestAddress As String, ByVal lpszAppName As String, ByVal lpszCalledParty As String, ByVal lpszComment As String) As Long

Declare Function tapiRequestDrop Lib "TAPI32.DLL" (ByVal hWnd As Integer, ByVal wRequestID As Integer) As Long

Declare Function lineRegisterRequestRecipient Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwRegistrationInstance As Long, ByVal dwRequestMode As Long, ByVal bEnable As Long) As Long

Declare Function tapiGetLocationInfo Lib "TAPI32.DLL" (ByVal lpszCountryCode As String, ByVal lpszCityCode As String) As Long

' Tapi Address Translation procedures

Declare Function lineSetCurrentLocation Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwLocation As Long) As Long

Declare Function lineSetTollList Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal lpszAddressIn As String, ByVal dwTollListOption As Long) As Long

Declare Function lineTranslateAddress Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal dwAPIVersion As Long, ByVal lpszAddressIn As String, ByVal dwCard As Long, ByVal dwTranslateOptions As Long, lpTranslateOutput As Any) As Long

Declare Function lineGetTranslateCaps Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwAPIVersion As Long, lpTranslateCaps As Any) As Long

' Tapi function prototypes

Declare Function lineAccept Lib "TAPI32.DLL" (ByVal hCall As Long, lpsUserUserInfo As Any, ByVal dwSize As Long) As Long

Declare Function lineAddToConference Lib "TAPI32.DLL" (ByVal hConfCall As Long, ByVal hConsultCall As Long) As Long

Declare Function lineAnswer Lib "TAPI32.DLL" (ByVal hCall As Long, lpsUserUserInfo As Any, ByVal dwSize As Long) As Long

Declare Function lineBlindTransfer Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal lpszDestAddress As String, ByVal dwCountryCode As Long) As Long

Declare Function lineClose Lib "TAPI32.DLL" (ByVal hLine As Long) As Long

Declare Function lineCompleteCall Lib "TAPI32.DLL" (ByVal hCall As Long, lpdwCompletionID As Long, ByVal dwCompletionMode As Long, ByVal dwMessageID As Long) As Long

Declare Function lineCompleteTransfer Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal hConsultCall As Long, lphConfCall As Long, ByVal dwTransferMode As Long) As Long

Declare Function lineConfigDialog Lib "TAPI32.DLL" (ByVal dwDeviceID As Long, ByVal hwndOwner As Integer, ByVal lpszDeviceClass As String) As Long

Declare Function lineDeallocateCall Lib "TAPI32.DLL" (ByVal hCall As Long) As Long

Declare Function lineDevSpecific Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, ByVal hCall As Long, lpParams As Any, ByVal dwSize As Long) As Long

Declare Function lineDevSpecificFeature Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwFeature As Long, lpParams As Any, ByVal dwSize As Long) As Long

Declare Function lineDial Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal lpszDestAddress As String, ByVal dwCountryCode As Long) As Long

Declare Function lineDrop Lib "TAPI32.DLL" (ByVal hCall As Long, lpsUserUserInfo As Any, ByVal dwSize As Long) As Long

Declare Function lineForward Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal bAllAddresses As Long, ByVal dwAddressID As Long, lpForwardList As Any, ByVal dwNumRingsNoAnswer As Long, lphConsultCall As Long, lpCallParams As Any) As Long

Declare Function lineGatherDigits Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwDigitModes As Long, lpsDigits As Any, ByVal dwNumDigits As Long, ByVal lpszTerminationDigits As String, ByVal dwFirstDigitTimeout As Long, ByVal dwInterDigitTimeout As Long) As Long

Declare Function lineGenerateDigits Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwDigitMode As Long, ByVal lpszDigits As String, ByVal dwDuration As Long) As Long

Declare Function lineGenerateTone Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwToneMode As Long, ByVal dwDuration As Long, ByVal dwNumTones As Long, lpTones As LINEGENERATETONE_TYPE) As Long

Declare Function lineGetAddressCaps Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal dwAddressID As Long, ByVal dwAPIVersion As Long, ByVal dwExtVersion As Long, ByVal lpAddressCaps As String) As Long

Declare Function lineGetAddressID Lib "TAPI32.DLL" (ByVal hLine As Long, lpdwAddressID As Long, ByVal dwAddressMode As Long, lpsAddress As Any, ByVal dwSize As Long) As Long

Declare Function lineGetAddressStatus Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, lpAddressStatus As Any) As Long

Declare Function lineGetCallInfo Lib "TAPI32.DLL" (ByVal hCall As Long, lpCallInfo As Any) As Long

Declare Function lineGetCallStatus Lib "TAPI32.DLL" (ByVal hCall As Long, lpCallStatus As Any) As Long

Declare Function lineGetConfRelatedCalls Lib "TAPI32.DLL" (ByVal hCall As Long, lpCallList As Any) As Long

Declare Function lineGetDevCaps Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal dwAPIVersion As Long, ByVal dwExtVersion As Long, lpLineDevCaps As LINEDEVCAPS) As Long

Declare Function lineGetDevConfig Lib "TAPI32.DLL" (ByVal dwDeviceID As Long, lpDeviceConfig As Any, ByVal lpszDeviceClass As String) As Long

Declare Function lineGetNewCalls Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, ByVal dwSelect As Long, lpCallList As Any) As Long

Declare Function lineGetIcon Lib "TAPI32.DLL" (ByVal dwDeviceID As Long, ByVal lpszDeviceClass As String, lphIcon As Integer) As Long

Declare Function lineGetID Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, ByVal hCall As Long, ByVal dwSelect As Long, lpDeviceID As Any, ByVal lpszDeviceClass As String) As Long

Declare Function lineGetLineDevStatus Lib "TAPI32.DLL" (ByVal hLine As Long, lpLineDevStatus As Any) As Long

Declare Function lineGetNumRings Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, lpdwNumRings As Long) As Long

Declare Function lineGetRequest Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwRequestMode As Long, lpRequestBuffer As Any) As Long

Declare Function lineGetStatusMessages Lib "TAPI32.DLL" (ByVal hLine As Long, lpdwLineStates As Long, lpdwAddressStates As Long) As Long

Declare Function lineHandoff Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal lpszFileName As String, ByVal dwMediaMode As Long) As Long

Declare Function lineHold Lib "TAPI32.DLL" (ByVal hCall As Long) As Long

'Declare Function lineInitialize Lib "TAPI32.DLL" (lphLineApp As Long, ByVal hInstance As Integer, ByVal lpfnCallback As Long, ByVal lpszAppName As String, lpdwNumDevs As Long) As Long
Declare Function lineInitialize Lib "TAPI32.DLL" (lphLineApp As Long, ByVal hInstance As Long, ByVal lpfnCallback As Long, ByVal lpszAppName As String, lpdwNumDevs As Long) As Long

Declare Function lineMakeCall Lib "TAPI32.DLL" (ByVal hLine As Long, lphCall As Long, ByVal lpszDestAddress As String, ByVal dwCountryCode As Long, lpCallParams As Any) As Long

Declare Function lineMonitorDigits Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwDigitModes As Long) As Long

Declare Function lineMonitorMedia Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwMediaModes As Long) As Long

Declare Function lineMonitorTones Lib "TAPI32.DLL" (ByVal hCall As Long, lpToneList As Any, ByVal dwNumEntries As Long) As Long

Declare Function lineNegotiateAPIVersion Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal dwAPILowVersion As Long, ByVal dwAPIHighVersion As Long, lpdwAPIVersion As Long, lpExtensionID As LINEEXTENSIONID) As Long

Declare Function lineNegotiateExtVersion Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, ByVal dwAPIVersion As Long, ByVal dwExtLowVersion As Long, ByVal dwExtHighVersion As Long, lpdwExtVersion As Long) As Long

Declare Function lineOpen Lib "TAPI32.DLL" (ByVal hLineApp As Long, ByVal dwDeviceID As Long, lphLine As Long, ByVal dwAPIVersion As Long, ByVal dwExtVersion As Long, ByVal dwCallbackInstance As Long, ByVal dwPrivileges As Long, ByVal dwMediaModes As Long, lpCallParams As Any) As Long

Declare Function linePark Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwParkMode As Long, ByVal lpszDirAddress As String, lpNonDirAddress As Any) As Long

Declare Function linePickup Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, lphCall As Long, ByVal lpszDestAddress As String, ByVal lpszGroupID As String) As Long

Declare Function linePrepareAddToConference Lib "TAPI32.DLL" (ByVal hConfCall As Long, lphConsultCall As Long, lpCallParams As Any) As Long

Declare Function lineRedirect Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal lpszDestAddress As String, ByVal dwCountryCode As Long) As Long

Declare Function lineRemoveFromConference Lib "TAPI32.DLL" (ByVal hCall As Long) As Long

Declare Function lineSecureCall Lib "TAPI32.DLL" (ByVal hCall As Long) As Long

Declare Function lineSendUserUserInfo Lib "TAPI32.DLL" (ByVal hCall As Long, lpsUserUserInfo As Any, ByVal dwSize As Long) As Long

Declare Function lineSetAppSpecific Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwAppSpecific As Long) As Long

Declare Function lineSetCallParams Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwBearerMode As Long, ByVal dwMinRate As Long, ByVal dwMaxRate As Long, lpDialParams As LINEDIALPARAMS) As Long

Declare Function lineSetCallPrivilege Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwCallPrivilege As Long) As Long

Declare Function lineSetDevConfig Lib "TAPI32.DLL" (ByVal dwDeviceID As Long, lpDeviceConfig As Any, ByVal dwSize As Long, ByVal lpszDeviceClass As String) As Long

Declare Function lineSetMediaControl Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, ByVal hCall As Long, ByVal dwSelect As Long, lpDigitList As Any, ByVal dwDigitNumEntries As Long, lpMediaList As Any, ByVal dwMediaNumEntries As Long, lpToneList As Any, ByVal dwToneNumEntries As Long, lpCallStateList As Any, ByVal dwCallStateNumEntries As Long) As Long

Declare Function lineSetMediaMode Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal dwMediaModes As Long) As Long

Declare Function lineSetNumRings Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, ByVal dwNumRings As Long) As Long

Declare Function lineSetStatusMessages Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwLineStates As Long, ByVal dwAddressStates As Long) As Long

Declare Function lineSetTerminal Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, ByVal hCall As Long, ByVal dwSelect As Long, ByVal dwTerminalModes As Long, ByVal dwTerminalID As Long, ByVal bEnable As Long) As Long

Declare Function lineSetupConference Lib "TAPI32.DLL" (ByVal hCall As Long, ByVal hLine As Long, lphConfCall As Long, lphConsultCall As Long, ByVal dwNumParties As Long, lpCallParams As Any) As Long

Declare Function lineSetupTransfer Lib "TAPI32.DLL" (ByVal hCall As Long, lphConsultCall As Long, lpCallParams As Any) As Long

Declare Function lineShutdown Lib "TAPI32.DLL" (ByVal hLineApp As Long) As Long

Declare Function lineSwapHold Lib "TAPI32.DLL" (ByVal hActiveCall As Long, ByVal hHeldCall As Long) As Long

Declare Function lineUncompleteCall Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwCompletionID As Long) As Long

Declare Function lineUnhold Lib "TAPI32.DLL" (ByVal hCall As Long) As Long

Declare Function lineUnpark Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwAddressID As Long, lphCall As Long, ByVal lpszDestAddress As String) As Long

Declare Function phoneClose Lib "TAPI32.DLL" (ByVal hPhone As Long) As Long

Declare Function phoneConfigDialog Lib "TAPI32.DLL" (ByVal dwDeviceID As Long, ByVal hwndOwner As Integer, ByVal lpszDeviceClass As String) As Long

Declare Function phoneDevSpecific Lib "TAPI32.DLL" (ByVal hPhone As Long, lpParams As Any, ByVal dwSize As Long) As Long

Declare Function phoneGetButtonInfo Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwButtonLampID As Long, lpButtonInfo As Any) As Long

Declare Function phoneGetData Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwDataID As Long, lpData As Any, ByVal dwSize As Long) As Long

Declare Function phoneGetDevCaps Lib "TAPI32.DLL" (ByVal hPhoneApp As Long, ByVal dwDeviceID As Long, ByVal dwAPIVersion As Long, ByVal dwExtVersion As Long, lpPhoneCaps As Any) As Long

Declare Function phoneGetDisplay Lib "TAPI32.DLL" (ByVal hPhone As Long, lpDisplay As Any) As Long

Declare Function phoneGetGain Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwHookSwitchDev As Long, lpdwGain As Long) As Long

Declare Function phoneGetHookSwitch Lib "TAPI32.DLL" (ByVal hPhone As Long, lpdwHookSwitchDevs As Long) As Long

Declare Function phoneGetIcon Lib "TAPI32.DLL" (ByVal dwDeviceID As Long, ByVal lpszDeviceClass As String, lphIcon As Integer) As Long

Declare Function phoneGetID Lib "TAPI32.DLL" (ByVal hPhone As Long, lpDeviceID As Any, ByVal lpszDeviceClass As String) As Long

Declare Function phoneGetLamp Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwButtonLampID As Long, lpdwLampMode As Long) As Long

Declare Function phoneGetRing Lib "TAPI32.DLL" (ByVal hPhone As Long, lpdwRingMode As Long, lpdwVolume As Long) As Long

Declare Function phoneGetStatus Lib "TAPI32.DLL" (ByVal hPhone As Long, lpPhoneStatus As Any) As Long

Declare Function phoneGetStatusMessages Lib "TAPI32.DLL" (ByVal hPhone As Long, lpdwPhoneStates As Long, lpdwButtonModes As Long, lpdwButtonStates As Long) As Long

Declare Function phoneGetVolume Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwHookSwitchDev As Long, lpdwVolume As Long) As Long

Declare Function phoneInitialize Lib "TAPI32.DLL" (lphPhoneApp As Long, ByVal hInstance As Long, ByVal lpfnCallback As Long, ByVal lpszAppName As String, lpdwNumDevs As Long) As Long

Declare Function phoneNegotiateAPIVersion Lib "TAPI32.DLL" (ByVal hPhoneApp As Long, ByVal dwDeviceID As Long, ByVal dwAPILowVersion As Long, ByVal dwAPIHighVersion As Long, lpdwAPIVersion As Long, lpExtensionID As PHONEEXTENSIONID) As Long

Declare Function phoneNegotiateExtVersion Lib "TAPI32.DLL" (ByVal hPhoneApp As Long, ByVal dwDeviceID As Long, ByVal dwAPIVersion As Long, ByVal dwExtLowVersion As Long, ByVal dwExtHighVersion As Long, lpdwExtVersion As Long) As Long

Declare Function phoneOpen Lib "TAPI32.DLL" (ByVal hPhoneApp As Long, ByVal dwDeviceID As Long, lphPhone As Long, ByVal dwAPIVersion As Long, ByVal dwExtVersion As Long, ByVal dwCallbackInstance As Long, ByVal dwPrivilege As Long) As Long

Declare Function phoneSetButtonInfo Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwButtonLampID As Long, lpButtonInfo As Any) As Long

Declare Function phoneSetData Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwDataID As Long, lpData As Any, ByVal dwSize As Long) As Long

Declare Function phoneSetDisplay Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwRow As Long, ByVal dwColumn As Long, lpsDisplay As Any, ByVal dwSize As Long) As Long

Declare Function phoneSetGain Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwHookSwitchDev As Long, ByVal dwGain As Long) As Long

Declare Function phoneSetHookSwitch Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwHookSwitchDevs As Long, ByVal dwHookSwitchMode As Long) As Long

Declare Function phoneSetLamp Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwButtonLampID As Long, ByVal dwLampMode As Long) As Long

Declare Function phoneSetRing Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwRingMode As Long, ByVal dwVolume As Long) As Long

Declare Function phoneSetStatusMessages Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwPhoneStates As Long, ByVal dwButtonModes As Long, ByVal dwButtonStates As Long) As Long

Declare Function phoneSetVolume Lib "TAPI32.DLL" (ByVal hPhone As Long, ByVal dwHookSwitchDev As Long, ByVal dwVolume As Long) As Long

Declare Function phoneShutdown Lib "TAPI32.DLL" (ByVal hPhoneApp As Long) As Long

