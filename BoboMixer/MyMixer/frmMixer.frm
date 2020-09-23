VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMixer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bobo Mixer"
   ClientHeight    =   3720
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   1560
   Icon            =   "frmMixer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   1560
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   3420
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2699
         EndProperty
      EndProperty
   End
   Begin MyMixer.BoboMixer BoboMixer1 
      Height          =   3375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   5953
   End
   Begin VB.Image ImgIcon 
      Height          =   240
      Left            =   360
      Picture         =   "frmMixer.frx":0442
      Top             =   3960
      Width           =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   960
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   1320
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu mnuHelpSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2002********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'How it all works..briefly
'The Concept: We get a device (there can be more than one soundcard or
'a soundcard can have multiple "Devices"). This is called
'a "Line". Each "Line" has at least one "Control" - eg CD Player
'Wave, Midi, etc. Each "Control" has different elements - Volume,
'Balance and Mute etc.

'To see this demo do it's stuff, run the App and open "Volume Control"
'(sndvol32.exe) at the same time and adjust sliders on either App
'and the other will respond accordingly.
Option Explicit
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim Mix_Ctl As MIXERCONTROL
Dim mMute As MIXERCONTROL
Dim m_Mixer As Long
Dim mName As String
Private Sub Form_Load()
    Dim z As Long, mxName As String, cnt As Long, q As Long, tmp As String
    Me.Icon = ImgIcon.Picture
    mixerOpen m_Mixer, 0, Me.hwnd, 0, CALLBACK_WINDOW
    For z = 0 To 5
        If GetVolumeControl(m_Mixer, GetCompType(z), MIXERCONTROL_CONTROLTYPE_VOLUME, Mix_Ctl) Then
            'Load a Usercontrol for each control
            If cnt <> 0 Then
                Load BoboMixer1(BoboMixer1.Count)
                BoboMixer1(cnt).Left = BoboMixer1(cnt - 1).Left + BoboMixer1(cnt - 1).Width
                BoboMixer1(cnt).Visible = True
            End If
            'Let the Usercontrol do the real work
            BoboMixer1(cnt).Initialise GetCompType(z)
            BoboMixer1(cnt).MixerName = mName
            mName = ""
            cnt = cnt + 1
        End If
    Next
    If cnt = 0 Then
        mixerClose m_Mixer
        MsgBox "Failed to open mixer"
        End
    Else
        'Loaded at least one control so start the hook
        'to receive messages from the mixer
        Hook Me.hwnd
        Me.Width = BoboMixer1(cnt - 1).Left + BoboMixer1(cnt - 1).Width
        SB.Panels(1).Text = GetDeviceName
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim z As Long
    'clean up each of the hooks for each usercontrol
    For z = 0 To BoboMixer1.Count - 1
        UnHookSlider BoboMixer1(z).BalanceWindow
        UnHookSlider BoboMixer1(z).VolumeWindow
        UnHookCheck BoboMixer1(z).CheckWindow
    Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Line1.X2 = Me.Width
    Line2.X2 = Me.Width
    Me.Height = BoboMixer1(0).Top + BoboMixer1(0).Height + SB.Height + 630
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'clean up after ourselves
    Unhook Me.hwnd
    mixerClose m_Mixer

End Sub

Private Sub mnuHelpAbout_Click()
    'seeing as this is a clone, make the copy complete!
    ShellAbout Me.hwnd, App.Title, "©PSST Software 2002" & vbCrLf & "www.psst.com.au", ByVal 0&
End Sub

Private Sub mnuHelpTopics_Click()
    'just for fun - windows help for the mixer
    'you should have this file
    Dim Path As String, strSave As String
    strSave = String(200, Chr$(0))
    Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\Help"
    If Dir(Path & "\sndvol32.chm") <> "" Then
        ShellExecute 0, vbNullString, Path & "\sndvol32.chm", "", Path, vbNormalNoFocus
    Else
        MsgBox "Failed to locate help file"
    End If
End Sub

Private Sub mnuOptionsExit_Click()
    Unload Me
End Sub
Public Function GetCompType(mIndex As Long) As Long
    Select Case mIndex
        Case 0
            GetCompType = MIXERLINE_COMPONENTTYPE_DST_SPEAKERS
            mName = "Volume Control"
        Case 1
            GetCompType = MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT
            mName = "Wave"
        Case 2
            GetCompType = MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER
            mName = "Midi"
        Case 3
            GetCompType = MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE
            mName = "Mic"
        Case 4
            GetCompType = MIXERLINE_COMPONENTTYPE_SRC_LINE
            mName = "Line-In"
        Case 5
            GetCompType = MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC
            mName = "CD Player"
    End Select
'        Other possibilities...
'            MIXERLINE_COMPONENTTYPE_DST_DIGITAL
'            MIXERLINE_COMPONENTTYPE_DST_HEADPHONES
'            MIXERLINE_COMPONENTTYPE_DST_LINE
'            MIXERLINE_COMPONENTTYPE_DST_MONITOR
'            MIXERLINE_COMPONENTTYPE_DST_TELEPHONE
'            MIXERLINE_COMPONENTTYPE_DST_UNDEFINED
'            MIXERLINE_COMPONENTTYPE_DST_VOICEIN
'            MIXERLINE_COMPONENTTYPE_DST_WAVEIN
'            MIXERLINE_COMPONENTTYPE_SRC_ANALOG
'            MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY
'            MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE
'            MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED
'            MIXERLINE_COMPONENTTYPE_SRC_DIGITAL
'            MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER

End Function

