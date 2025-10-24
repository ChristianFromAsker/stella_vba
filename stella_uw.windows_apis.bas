Attribute VB_Name = "windows_apis"
Option Compare Database
Option Explicit

    #If Win64 Then
        Public Type POINTAPI
            x As Long
            y As Long
        End Type
        Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hwnd As LongPtr, lpPoint As POINTAPI) As Long
        Declare PtrSafe Function apiGetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
        Declare PtrSafe Function apiMoveWindow Lib "user32" Alias "MoveWindow" ( _
            ByVal hwnd As LongPtr, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal nWidth As Long, _
            ByVal nHeight As Long, _
            ByVal bRepaint As Long _
         ) As Long
         Declare PtrSafe Function apiGetParent Lib "user32" Alias "GetParent" (ByVal hwnd As LongPtr) As Long
         Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
    #Else
        Declare Function apiGetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
        Declare Function apiGetParent Lib "user32" Alias "GetParent" (ByVal hWnd As LongPtr) As Long
        Declare Function apiMoveWindow Lib "user32" Alias "MoveWindow" ( _
            ByVal hWnd As LongPtr, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal nWidth As Long, _
            ByVal nHeight As Long, _
            ByVal bRepaint As Long _
         ) As Long
    #End If

    ' DPI
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long

    ' Work area (desktop minus taskbar)
    Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
        ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

    ' Fallback: full screen size
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const SPI_GETWORKAREA As Long = 48
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const TPI As Long = 1440 ' twips per inch

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Sub GetWorkAreaPx(ByRef leftPx As Long, ByRef topPx As Long, ByRef widthPx As Long, ByRef heightPx As Long)
    Dim r As RECT
    If SystemParametersInfo(SPI_GETWORKAREA, 0, r, 0) <> 0 Then
        leftPx = r.Left
        topPx = r.Top
        widthPx = r.Right - r.Left
        heightPx = r.Bottom - r.Top
    Else
        ' Fallback to full screen if SPI_GETWORKAREA fails
        leftPx = 0
        topPx = 0
        widthPx = GetSystemMetrics(SM_CXSCREEN)
        heightPx = GetSystemMetrics(SM_CYSCREEN)
    End If
End Sub

Public Function PositionFormFillVertical(Optional ByVal widthFraction As Double = 0.6 _
, Optional ByVal anchor As String = "center")

    Dim leftPx As Long
    Dim topPx As Long
    Dim workWpx As Long
    Dim workHpx As Long
    Dim dpiX As Long
    Dim dpiY As Long
    Dim workWtw As Long
    Dim workHtw As Long
    Dim workLeftTw As Long
    Dim workTopTw As Long
    Dim widthTw As Long
    Dim leftTw As Long
    Dim topTw As Long

    ' 1) Get desktop work area in pixels
    GetWorkAreaPx leftPx, topPx, workWpx, workHpx

    ' 2) Get screen DPI (per-axis; respects Windows display scaling)
    dpiX = ScreenDPI(True)
    dpiY = ScreenDPI(False)
    If dpiX = 0 Then dpiX = 96
    If dpiY = 0 Then dpiY = 96

    ' 3) Convert to twips
    workLeftTw = PxToTwX(leftPx, dpiX)
    workTopTw = PxToTwY(topPx, dpiY)
    workWtw = PxToTwX(workWpx, dpiX)
    workHtw = PxToTwY(workHpx, dpiY)

    ' 4) Decide target size/position
    If widthFraction <= 0 Or widthFraction > 1 Then widthFraction = 0.6
    widthTw = CLng(workWtw * widthFraction + 0.5)
    topTw = workTopTw ' align to top edge of work area

    Select Case LCase$(anchor)
        Case "left"
            leftTw = workLeftTw
        Case "right"
            leftTw = workLeftTw + (workWtw - widthTw)
        Case Else ' "center"
            leftTw = workLeftTw + (workWtw - widthTw) \ 2
    End Select

    PositionFormFillVertical = workHtw
    
End Function

Function ScreenDPI(Optional ByVal isX As Boolean = True) As Long
    Dim hdc As LongPtr
    Dim dpi As Long
    
    hdc = GetDC(0)
    If hdc Then
        If isX Then
            dpi = GetDeviceCaps(hdc, LOGPIXELSX)
        Else
            dpi = GetDeviceCaps(hdc, LOGPIXELSY)
        End If
        ReleaseDC 0, hdc
    End If
    ScreenDPI = dpi
End Function

Private Function PxToTwX(px As Long, dpiX As Long) As Long
    PxToTwX = CLng(px * (TPI / dpiX) + 0.5)
End Function

Private Function PxToTwY(px As Long, dpiY As Long) As Long
    PxToTwY = CLng(px * (TPI / dpiY) + 0.5)
End Function


Function AccessMoveSize(iX As Integer, iY As Integer, iWidth As Integer, iHeight As Integer)
    apiMoveWindow GetAccesshWnd(), iX, iY, iWidth, iHeight, True
End Function
Function GetAccesshWnd()
    Dim hwnd As LongPtr
    Dim hWndAccess As LongPtr
    hwnd = apiGetActiveWindow()
    hWndAccess = hwnd
    While hwnd <> 0
        hWndAccess = hwnd
        hwnd = apiGetParent(hwnd)
    Wend
    GetAccesshWnd = hWndAccess
End Function
Function MouseX(Optional ByVal hwnd As LongPtr) As Long
'Get mouse X coordinates in pixels. If a window handle is passed, the result is relative to the client area
' of that window, otherwise the result is relative to the screen
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hwnd Then ScreenToClient hwnd, lpPoint
    MouseX = lpPoint.x
End Function
Function MouseY(Optional ByVal hwnd As LongPtr) As Long
' Get mouse Y coordinates in pixels
' If a window handle is passed, the result is relative to the client area
' of that window, otherwise the result is relative to the screen
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hwnd Then ScreenToClient hwnd, lpPoint
    MouseY = lpPoint.y
End Function


