VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl BoboMixer 
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   ScaleHeight     =   3375
   ScaleWidth      =   1440
   Begin VB.PictureBox PicFocus 
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   3480
      Width           =   855
   End
   Begin VB.CheckBox chMute 
      Caption         =   "Mute"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2985
      Width           =   975
   End
   Begin MSComctlLib.Slider slVolume 
      Height          =   1365
      Left            =   360
      TabIndex        =   3
      Top             =   1575
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   2408
      _Version        =   393216
      Orientation     =   1
      Max             =   100
      SelStart        =   50
      TickStyle       =   2
      TickFrequency   =   17
      Value           =   50
   End
   Begin MSComctlLib.Slider slBalance 
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   705
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393216
      Min             =   -100
      Max             =   100
      TickFrequency   =   100
   End
   Begin VB.Image ImgLeft 
      Height          =   240
      Left            =   120
      Picture         =   "BoboMixer.ctx":0000
      Top             =   765
      Width           =   240
   End
   Begin VB.Image ImgRight 
      Height          =   240
      Left            =   1080
      Picture         =   "BoboMixer.ctx":014A
      Top             =   765
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Volume:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1365
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Balance:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   465
      Width           =   735
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   120
      X2              =   1200
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   120
      X2              =   1200
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label lblControl 
      Caption         =   "Volume Control"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   105
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   1395
      X2              =   1395
      Y1              =   3345
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   1410
      X2              =   1410
      Y1              =   3345
      Y2              =   0
   End
End
Attribute VB_Name = "BoboMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2002********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive
Option Explicit
Private Const TBM_GETTOOLTIPS = WM_USER + 30
Private Const TTM_ACTIVATE = WM_USER + 1
Dim Mix_Ctl As MIXERCONTROL
Dim mMute As MIXERCONTROL
Dim m_Mixer As Long
Dim DontSet As Boolean
Dim hmem As Long
Dim m_IsStereo As Boolean
Public Sub Initialise(mType As Long)
    Dim temp As String, z As Long
    'Get the control as allocated from frmMixer
    GetVolumeControl m_Mixer, mType, MIXERCONTROL_CONTROLTYPE_VOLUME, Mix_Ctl, temp, m_IsStereo
    GetVolumeControl m_Mixer, mType, MIXERCONTROL_CONTROLTYPE_MUTE, mMute
    'No Focus Rectangles thanks all the same!
    HookSlider slVolume.hwnd
    HookSlider slBalance.hwnd
    HookCheck chMute.hwnd
    'What are we called?
    MixerName = temp
    'Stereo?
    slBalance.Enabled = m_IsStereo
    'What is the current Volume
    GetVolume
    'What is the current Balance
    If m_IsStereo Then GetBalance
    'Is the Mute on?
    GetMute
    'Not tooltips for the sliders
    z = SendMessage(slVolume.hwnd, TBM_GETTOOLTIPS, 0&, ByVal 0&)
    If z <> 0 Then SendMessage z, TTM_ACTIVATE, 0&, ByVal 0&
    z = SendMessage(slBalance.hwnd, TBM_GETTOOLTIPS, 0&, ByVal 0&)
    If z <> 0 Then SendMessage z, TTM_ACTIVATE, 0&, ByVal 0&
End Sub

Public Property Get MixerName() As String
    MixerName = lblControl.Caption
End Property
Public Property Let MixerName(ByVal vNewValue As String)
    lblControl.Caption = vNewValue
End Property
Public Property Get Mixer() As Long
    'Needed for the Callback to identify this Control
    Mixer = Mix_Ctl.dwControlID
End Property
Public Property Get Muter() As Long
    'Needed for the Callback to identify this Control
    Muter = mMute.dwControlID
End Property
Public Property Get BalanceWindow() As Long
    'Needed for the Hook so frmMixer can unhook on Form_Unload
    BalanceWindow = slBalance.hwnd
End Property
Public Property Get VolumeWindow() As Long
    'Needed for the Hook so frmMixer can unhook on Form_Unload
    VolumeWindow = slVolume.hwnd
End Property
Public Property Get CheckWindow() As Long
    'Needed for the Hook so frmMixer can unhook on Form_Unload
    CheckWindow = chMute.hwnd
End Property
Public Property Get IsStereo() As Boolean
    IsStereo = m_IsStereo
End Property

Public Property Let IsStereo(ByVal vNewValue As Boolean)
    m_IsStereo = vNewValue
End Property
Private Sub chMute_Click()
    SetMute CBool(chMute.Value)
End Sub
Public Function GetVolume() As Long
    'Slightly modified MS code to get the current volume
    Dim z As Long
    Dim MxDetails As MIXERCONTROLDETAILS
    Dim MxVolume As MIXERCONTROLDETAILS_UNSIGNED
    Dim mMax As Long
    Dim mValue As Long
    Dim mTick As Long
    If DontSet Then Exit Function
    On Error Resume Next
    MxDetails.cbStruct = Len(MxDetails)
    MxDetails.dwControlID = Mix_Ctl.dwControlID
    MxDetails.cChannels = 1
    MxDetails.item = 0
    MxDetails.cbDetails = Len(MxVolume)
    MxDetails.paDetails = 0
    hmem = GlobalAlloc(&H40, Len(MxVolume))
    MxDetails.paDetails = GlobalLock(hmem)
    z = mixerGetControlDetails(m_Mixer, MxDetails, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr MxVolume, MxDetails.paDetails, Len(MxVolume)
    GlobalFree hmem
    'Adjust the slider accordingly
    mMax = Mix_Ctl.lMaximum
    mTick = mMax / 7
    If z = 0 Then
        GetVolume = MxVolume.dwValue
        mValue = mMax - MxVolume.dwValue
    Else
        GetVolume = 0
        mValue = 0
    End If
    DontSet = True 'Stops interaction with the balance slider while we change
    If slVolume.Max <> mMax Then slVolume.Max = mMax
    If slVolume.TickFrequency <> mTick Then slVolume.TickFrequency = mTick
    If slVolume.Value <> mValue Then slVolume.Value = mValue
    DontSet = False
End Function
Public Sub SetVolume(mValue As Long)
    Dim z As Long
    Dim MxDetails As MIXERCONTROLDETAILS
    Dim MxVolume As MIXERCONTROLDETAILS_UNSIGNED
    If m_IsStereo Then
        'To adjust volume we normally actually adjust the balance
        SetBalance slBalance.Value
    Else
        'unless of course we are mono
        MxDetails.cbStruct = Len(MxDetails)
        MxDetails.dwControlID = Mix_Ctl.dwControlID
        MxDetails.cChannels = 1
        MxDetails.item = 0
        MxDetails.cbDetails = Len(MxVolume)
        hmem = GlobalAlloc(&H40, Len(MxVolume))
        MxDetails.paDetails = GlobalLock(hmem)
        MxVolume.dwValue = mValue
        CopyPtrFromStruct MxDetails.paDetails, MxVolume, Len(MxVolume)
        z = mixerSetControlDetails(m_Mixer, MxDetails, MIXER_SETCONTROLDETAILSF_VALUE)
        GlobalFree (hmem)
    End If
End Sub
Public Function GetBalance() As Long
    Dim z As Long
    Dim MxDetails As MIXERCONTROLDETAILS
    Dim MxVolume(1) As MIXERCONTROLDETAILS_UNSIGNED
    Dim mMax As Long
    Dim mValue As Long
    Dim mTick As Long, tmpVol As Long
    Dim Cheated As Boolean
    If DontSet Then Exit Function
    On Error Resume Next
    tmpVol = (slVolume.Max - slVolume.Value)
    If tmpVol = 0 Then
        SetVolume 655
        Cheated = True
    End If
    MxDetails.item = Mix_Ctl.cMultipleItems
    MxDetails.dwControlID = Mix_Ctl.dwControlID
    MxDetails.cbStruct = Len(MxDetails)
    MxDetails.cbDetails = Len(MxVolume(0))
    MxDetails.cChannels = 2
    hmem = GlobalAlloc(&H40, Len(MxVolume(0)))
    MxDetails.paDetails = GlobalLock(hmem)
    z = mixerGetControlDetails(m_Mixer, MxDetails, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr MxVolume(0).dwValue, MxDetails.paDetails, Len(MxVolume(1)) * MxDetails.cChannels
    GlobalFree hmem
    GetBalance = (MxVolume(1).dwValue - MxVolume(0).dwValue)
    tmpVol = (slVolume.Max - slVolume.Value)
    If tmpVol = 0 Then tmpVol = 655
    mValue = ((MxVolume(1).dwValue - MxVolume(0).dwValue) / tmpVol) * 100
    DontSet = True
    If slBalance.Value <> mValue Then slBalance.Value = mValue
    If Cheated Then SetVolume 0
    DontSet = False
End Function
Public Sub SetBalance(mValue As Long)
    Dim z As Long
    Dim MxDetails As MIXERCONTROLDETAILS
    Dim MxVolume(1) As MIXERCONTROLDETAILS_UNSIGNED
    Dim volL As Long, volR As Long, tmpVol As Long
    tmpVol = slVolume.Max - slVolume.Value
    tmpVol = IIf(tmpVol = 0, 655, tmpVol)
    volR = tmpVol * (IIf(mValue >= 0, 1, (100 + mValue) / 100))
    volL = tmpVol * (IIf(mValue <= 0, 1, (100 - mValue) / 100))
    MxDetails.item = Mix_Ctl.cMultipleItems
    MxDetails.dwControlID = Mix_Ctl.dwControlID
    MxDetails.cbStruct = Len(MxDetails)
    MxDetails.cbDetails = Len(MxVolume(0))
    MxDetails.cChannels = 2
    hmem = GlobalAlloc(&H40, Len(MxVolume(0)))
    MxDetails.paDetails = GlobalLock(hmem)
    MxVolume(1).dwValue = volR
    MxVolume(0).dwValue = volL
    'two channels
    CopyPtrFromStruct MxDetails.paDetails, MxVolume(1).dwValue, Len(MxVolume(0)) * MxDetails.cChannels
    CopyPtrFromStruct MxDetails.paDetails, MxVolume(0).dwValue, Len(MxVolume(1)) * MxDetails.cChannels
    z = mixerSetControlDetails(m_Mixer, MxDetails, MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree hmem
End Sub
Public Function GetMute() As Boolean
    Dim z As Long
    Dim MxDetails As MIXERCONTROLDETAILS
    Dim MxVolume As MIXERCONTROLDETAILS_UNSIGNED
    MxDetails.cbStruct = Len(MxDetails)
    MxDetails.dwControlID = mMute.dwControlID
    MxDetails.cChannels = 1
    MxDetails.item = 0
    MxDetails.cbDetails = Len(MxVolume)
    MxDetails.paDetails = 0
    hmem = GlobalAlloc(&H40, Len(MxVolume))
    MxDetails.paDetails = GlobalLock(hmem)
    z = mixerGetControlDetails(m_Mixer, MxDetails, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr MxVolume, MxDetails.paDetails, Len(MxVolume)
    GlobalFree (hmem)
    GetMute = MxVolume.dwValue
    If GetMute Then
        chMute.Value = 1
    Else
        chMute.Value = 0
    End If
End Function
Public Sub SetMute(ByVal vNewValue As Boolean)
    Dim MxDetails As MIXERCONTROLDETAILS
    Dim MxVolume As MIXERCONTROLDETAILS_UNSIGNED
    MxDetails.cbStruct = Len(MxDetails)
    MxDetails.dwControlID = mMute.dwControlID
    MxDetails.cChannels = 1
    MxDetails.item = 0
    MxDetails.cbDetails = Len(MxVolume)
    hmem = GlobalAlloc(&H40, Len(MxVolume))
    MxDetails.paDetails = GlobalLock(hmem)
    MxVolume.dwValue = vNewValue
    CopyPtrFromStruct MxDetails.paDetails, MxVolume, Len(MxVolume)
    mixerSetControlDetails m_Mixer, MxDetails, MIXER_SETCONTROLDETAILSF_VALUE
    GlobalFree (hmem)
End Sub
Private Sub slBalance_Scroll()
    SetBalance slBalance.Value
End Sub
Private Sub slVolume_Scroll()
    Dim z As Long
    'Vertical sliders are upside down
    z = slVolume.Max - slVolume.Value
    SetVolume z
End Sub
Private Sub UserControl_Resize()
    UserControl.Width = 1440
    UserControl.Height = 3375
End Sub



