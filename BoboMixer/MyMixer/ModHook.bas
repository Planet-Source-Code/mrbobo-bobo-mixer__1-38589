Attribute VB_Name = "WindowHook"
'******************************************************************
'***************Copyright PSST 2002********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive
Option Explicit
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const GWL_WNDPROC = (-4)
Public Const CALLBACK_WINDOW = &H10000
Private Const WM_SETFOCUS = &H7
Private Const MM_MIXM_CONTROL_CHANGE = &H3D1
Public Const WM_USER = &H400
Public lpPrevWndProc As Long
Public lpPrevSliderProc As Long
Public lpPrevCheckProc As Long
'This is the main hook for Mixer messages we need to respond to
Public Sub Hook(mHwnd As Long)
    lpPrevWndProc = SetWindowLong(mHwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub Unhook(mHwnd As Long)
    SetWindowLong mHwnd, GWL_WNDPROC, lpPrevWndProc
End Sub
Function WindowProc(ByVal mHwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim z As Long
    Select Case uMsg
        Case MM_MIXM_CONTROL_CHANGE
            'Mixer message - either Volume/Balance or Mute has
            'been changed by another application
            For z = 0 To frmMixer.BoboMixer1.Count - 1
                'which control? Volume/Balance or Mute?
                If lParam = frmMixer.BoboMixer1(z).Mixer Or lParam = frmMixer.BoboMixer1(z).Muter Then
                    frmMixer.BoboMixer1(z).GetVolume
                    If frmMixer.BoboMixer1(z).IsStereo Then frmMixer.BoboMixer1(z).GetBalance
                    frmMixer.BoboMixer1(z).GetMute
                    Exit For
                End If
            Next
    End Select
    WindowProc = CallWindowProc(lpPrevWndProc, mHwnd, uMsg, wParam, lParam)
End Function
'These hooks are just for looks - they stop the Focus Rectangle
'on the Sliders and Checkboxes
Public Sub HookSlider(mHwnd As Long)
    lpPrevSliderProc = SetWindowLong(mHwnd, GWL_WNDPROC, AddressOf SliderProc)
End Sub
Public Sub UnHookSlider(mHwnd As Long)
    SetWindowLong mHwnd, GWL_WNDPROC, lpPrevSliderProc
End Sub
Public Function SliderProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Select Case uMsg
        Case WM_SETFOCUS
            Exit Function
    End Select
    SliderProc = CallWindowProc(lpPrevSliderProc, hwnd&, uMsg&, wParam&, lParam&)
End Function
Public Sub HookCheck(mHwnd As Long)
    lpPrevCheckProc = SetWindowLong(mHwnd, GWL_WNDPROC, AddressOf CheckProc)
End Sub
Public Sub UnHookCheck(mHwnd As Long)
    SetWindowLong mHwnd, GWL_WNDPROC, lpPrevCheckProc
End Sub
Public Function CheckProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Select Case uMsg
        Case WM_SETFOCUS
            Exit Function
    End Select
    CheckProc = CallWindowProc(lpPrevCheckProc, hwnd&, uMsg&, wParam&, lParam&)
End Function


