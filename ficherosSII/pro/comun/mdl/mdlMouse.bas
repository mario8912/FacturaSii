Attribute VB_Name = "mdlMouse"
Option Explicit

' declaraciones del api
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    wParam As Any, lParam As Any) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Constantes
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const GWL_WNDPROC = (-4)
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_VSCROLL As Integer = &H115

Private qFrm As Form
Private qGrid As DataGrid

Dim PrevProc As Long
Public Sub RRA(Obj As Object, Optional frm As Form)
    Set qGrid = Obj
    Set qFrm = frm
    
    PrevProc = SetWindowLong(Obj.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub RRD(Obj As Object)
    SetWindowLong Obj.hwnd, GWL_WNDPROC, PrevProc
End Sub

Public Sub HookForm(Obj As Object)
    PrevProc = SetWindowLong(Obj.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookForm(Obj As Object)
    SetWindowLong Obj.hwnd, GWL_WNDPROC, PrevProc
End Sub

' Procedimiento qie intercepta los mensajes de windows, en este caso para _
  interceptar el uso del Scroll del mouse
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function WindowProc(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long

    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
    
    If uMsg = WM_MOUSEWHEEL Then
        If wParam < 0 Then
            ' envia mediante SendMessage el comando para mover el Scroll hacia abajo
            If Not qFrm Is Nothing Then
                qFrm.qGrid.Scroll 0, 1
                Set qFrm = Nothing
            Else
                SendMessage hwnd, WM_VSCROLL, ByVal 1, ByVal 0
            End If
        Else
            ' Mueve el scroll hacia arriba
            If Not qFrm Is Nothing Then
                qFrm.qGrid.Scroll 0, -1
                Set qFrm = Nothing
            Else
            SendMessage hwnd, WM_VSCROLL, ByVal 0, ByVal 0
            End If
        End If
    End If
End Function


