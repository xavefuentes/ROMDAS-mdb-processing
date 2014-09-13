Attribute VB_Name = "Module1"
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Const NV_CLOSEMSGBOX As Long = &H5000&
    Private sLastTitle As String


Public Function ACmsgbox(AutoCloseSeconds As Long, prompt As String, Optional buttons As Long, _
    Optional title As String, Optional helpfile As String, _
    Optional context As Long) As Long
    sLastTitle = title
    SetTimer Screen.ActiveForm.hWnd, NV_CLOSEMSGBOX, AutoCloseSeconds * 1000, AddressOf TimerProc
    ACmsgbox = MsgBox(prompt, buttons, title, helpfile, context)
End Function


Private Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Dim hMessageBox As Long
    KillTimer hWnd, idEvent
    
    Select Case idEvent
        Case NV_CLOSEMSGBOX
        hMessageBox = FindWindow("#32770", sLastTitle)
        If hMessageBox Then
            Call SetForegroundWindow(hMessageBox)
            SendKeys "{enter}"
        End If
        sLastTitle = vbNullString
    End Select
End Sub
