Attribute VB_Name = "ModMixer"
'******************************************************************
'***************Copyright PSST 2002********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'Mixer API straight from MS with a few tiny tweaks to
'the Types for compatability with VB

'Note: not all these constants and declarations are
'used in this demo,I thought they might be helpful however.
Option Explicit
Public Const MMSYSERR_NOERROR = 0
Public Const MAXPNAMELEN = 32
Public Const MIXER_GETCONTROLDETAILSF_LISTTEXT = &H1&
Public Const MIXER_GETCONTROLDETAILSF_QUERYMASK = &HF&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ALL = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Public Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF&
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETLINEINFOF_DESTINATION = &H0&
Public Const MIXER_GETLINEINFOF_LINEID = &H2&
Public Const MIXER_GETLINEINFOF_QUERYMASK = &HF&
Public Const MIXER_GETLINEINFOF_SOURCE = &H1&
Public Const MIXER_GETLINEINFOF_TARGETTYPE = &H4&
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_OBJECTF_AUX = &H50000000
Public Const MIXER_OBJECTF_HANDLE = &H80000000
Public Const MIXER_OBJECTF_MIDIIN = &H40000000
Public Const MIXER_OBJECTF_MIDIOUT = &H30000000
Public Const MIXER_OBJECTF_MIXER = &H0&
Public Const MIXER_OBJECTF_WAVEIN = &H20000000
Public Const MIXER_OBJECTF_WAVEOUT = &H10000000
Public Const MIXER_SETCONTROLDETAILSF_CUSTOM = &H1&
Public Const MIXER_SETCONTROLDETAILSF_QUERYMASK = &HF&
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXERCONTROL_CONTROLF_DISABLED = &H80000000
Public Const MIXERCONTROL_CONTROLF_MULTIPLE = &H2&
Public Const MIXERCONTROL_CONTROLF_UNIFORM = &H1&
Public Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Public Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Public Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Public Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Public Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Public Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Public Const MIXERCONTROL_CT_CLASS_TIME = &H60000000
Public Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Public Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0&
Public Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Public Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Public Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
Public Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0&
Public Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
Public Const MIXERCONTROL_CT_SUBCLASS_MASK = &HF000000
Public Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Public Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0&
Public Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000 ' 10ths
Public Const MIXERCONTROL_CT_UNITS_MASK = &HFF0000
Public Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000 ' 10ths
Public Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERLINE_LINEF_ACTIVE = &H1&
Public Const MIXERLINE_LINEF_DISCONNECTED = &H8000&
Public Const MIXERLINE_LINEF_SOURCE = &H80000000
Public Const MIXERLINE_TARGETTYPE_AUX = 5
Public Const MIXERLINE_TARGETTYPE_MIDIIN = 4
Public Const MIXERLINE_TARGETTYPE_MIDIOUT = 3
Public Const MIXERLINE_TARGETTYPE_UNDEFINED = 0
Public Const MIXERLINE_TARGETTYPE_WAVEIN = 2
Public Const MIXERLINE_TARGETTYPE_WAVEOUT = 1
Public Const MIXERR_BASE = 1024
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXER_OBJECTF_HMIDIIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIIN)
Public Const MIXER_OBJECTF_HMIDIOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)
Public Const MIXER_OBJECTF_HMIXER = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)
Public Const MIXER_OBJECTF_HWAVEIN = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)
Public Const MIXER_OBJECTF_HWAVEOUT = (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_BASS_BOOST = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + &H2277)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEANMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_BUTTON = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BUTTON Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_CUSTOM = (MIXERCONTROL_CT_CLASS_CUSTOM Or MIXERCONTROL_CT_UNITS_CUSTOM)
Public Const MIXERCONTROL_CONTROLTYPE_DECIBELS = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_DECIBELS)
Public Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
Public Const MIXERCONTROL_CONTROLTYPE_LOUDNESS = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 4)
Public Const MIXERCONTROL_CONTROLTYPE_MICROTIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MICROSECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_MILLITIME = (MIXERCONTROL_CT_CLASS_TIME Or MIXERCONTROL_CT_SC_TIME_MILLISECS Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MIXER = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MONO = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
Public Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Public Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)
Public Const MIXERCONTROL_CONTROLTYPE_ONOFF = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
Public Const MIXERCONTROL_CONTROLTYPE_SLIDER = (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_PERCENT = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_PERCENT)
Public Const MIXERCONTROL_CONTROLTYPE_QSOUNDPAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_SIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_SIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_SRS_MTS = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 6)
Public Const MIXERCONTROL_CONTROLTYPE_SRS_ONOFF = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 7)
Public Const MIXERCONTROL_CONTROLTYPE_SRS_SYNTHSELECT = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 8)
Public Const MIXERCONTROL_CONTROLTYPE_STEREOENH = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 5)
Public Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNED = (MIXERCONTROL_CT_CLASS_NUMBER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Public Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 5)
Public Const MIXERLINE_COMPONENTTYPE_DST_LAST = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_DST_LINE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_DST_MONITOR = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 0)
Public Const MIXERLINE_COMPONENTTYPE_DST_VOICEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_SRC_ANALOG = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)
Public Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)
Public Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Public Const MIXERLINE_COMPONENTTYPE_SRC_DIGITAL = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LAST = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 10)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 0)
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Public Const MIXERR_INVALCONTROL = (MIXERR_BASE + 1)
Public Const MIXERR_INVALLINE = (MIXERR_BASE + 0)
Public Const MIXERR_INVALVALUE = (MIXERR_BASE + 2)
Public Const MIXERR_LASTERROR = (MIXERR_BASE + 2)
Public Type MIXERCAPS
    wMid As Integer          ' manufacturer id
    wPid As Integer          ' product id
    vDriverVersion As Long      ' version of the driver
    szPname As String * MAXPNAMELEN  ' product name
    fdwSupport As Long       ' misc. support bits
    cDestinations As Long     ' count of destinations
End Type

Public Type MIXERCONTROL
     cbStruct As Long           '  size in Byte of MIXERCONTROL
     dwControlID As Long        '  unique control id for mixer device
     dwControlType As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
     fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
     cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
     szShortName As String * MIXER_SHORT_NAME_CHARS  ' short name of control
     szName As String * MIXER_LONG_NAME_CHARS        ' long name of control
     lMinimum As Long           '  Minimum value
     lMaximum As Long           '  Maximum value
     reserved(10) As Long       '  reserved structure space
End Type
      

Public Type MIXERCONTROLDETAILS
    cbStruct As Long    ' size Byte of MIXERCONTROLDETAILS
    dwControlID As Long  ' control id to get/set details on
    cChannels As Long   ' number of channels paDetails array
    item As Long              ' hwndOwner or cMultipleItems
    cbDetails As Long   ' size of _one_ details_XX struct
    paDetails As Long   ' pointer to array of details_XX structs
End Type

Public Type MIXERCONTROLDETAILS_BOOLEAN
    fValue As Long
End Type

Public Type MIXERCONTROLDETAILS_LISTTEXT
    dwParam1 As Long
    dwParam2 As Long
    szName As String * MIXER_LONG_NAME_CHARS
End Type

Public Type MIXERCONTROLDETAILS_SIGNED
    lValue As Long
End Type

Public Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long
End Type
Public Type Target  ' for use MIXERLINE and others (embedded structure)
    
    dwType As Long         ' MIXERLINE_TARGETTYPE_xxxx
    dwDeviceID As Long       ' target device ID of device type
    wMid As Integer          ' of target device
    wPid As Integer          '    "
    vDriverVersion As Long      '    "
    szPname As String * MAXPNAMELEN
End Type

Public Type MIXERLINE
    cbStruct As Long        ' size of MIXERLINE structure
    dwDestination As Long     ' zero based destinationdex
    dwSource As Long        ' zero based sourcedex (if source)
    dwLineID As Long        ' unique line id for mixer device
    fdwLine As Long        ' state/information about line
    dwUser As Long         ' driver specificformation
    dwComponentType As Long    ' component type line connects to
    cChannels As Long       ' number of channels line supports
    cConnections As Long      ' number of connections (possible)
    cControls As Long       ' number of controls at this line
    szShortName As String * MIXER_SHORT_NAME_CHARS
    szName As String * MIXER_LONG_NAME_CHARS
    lpTarget As Target
End Type

Public Type MIXERLINECONTROLS
    cbStruct As Long
    dwLineID As Long
    dwControl As Long
    cControls As Long
    cbmxctrl As Long
    pamxctrl As Long
End Type


Public Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Public Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Public Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId As Long, pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long
Public Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, pumxID As Long, ByVal fdwId As Long) As Long
Public Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Public Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Public Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function mixerMessage Lib "winmm.dll" (ByVal hmx As Long, ByVal uMsg As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long
Public Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Public Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
'Memory API
Public Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Public Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Function GetDeviceName() As String
    'Name of soundcard
    Dim MxCaps As MIXERCAPS, z As Long
    z = mixerGetDevCaps(0, MxCaps, Len(MxCaps))
    GetDeviceName = Left(MxCaps.szPname, InStr(1, MxCaps.szPname, Chr(0)) - 1)
End Function
Public Function GetVolumeControl(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL, Optional ByRef mCTLname As String, Optional ByRef m_IsStereo As Boolean) As Boolean
    'gets info about one control and puts it into
    'a structure we can work with
    Dim mxlc As MIXERLINECONTROLS
    Dim mxl As MIXERLINE
    Dim rc As Long
    Dim hmem As Long
    mCTLname = ""
    mxl.cbStruct = Len(mxl)
    mxl.dwComponentType = componentType
    rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
    If (MMSYSERR_NOERROR = rc) Then
        mxlc.cbStruct = Len(mxlc)
        mxlc.dwLineID = mxl.dwLineID
        mxlc.dwControl = ctrlType
        mxlc.cControls = 1
        mxlc.cbmxctrl = Len(mxc)
        hmem = GlobalAlloc(&H40, Len(mxc))
        mxlc.pamxctrl = GlobalLock(hmem)
        mxc.cbStruct = Len(mxc)
        rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
        m_IsStereo = mxl.cChannels > 1
        If (MMSYSERR_NOERROR = rc) Then
            CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
            GetVolumeControl = True
        Else
            GetVolumeControl = False
        End If
        GlobalFree hmem
        Exit Function
    End If
    GetVolumeControl = False
End Function
