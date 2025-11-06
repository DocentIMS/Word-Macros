Attribute VB_Name = "ZZ_dpi"
Option Explicit

Private Const LOGPIXELSX As Long = 88

#If VBA7 Then
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function EnumDisplayMonitors Lib "user32" (ByVal hdc As LongPtr, ByRef lprcClip As Any, ByVal lpfnEnum As LongPtr, ByVal dwData As Long) As Long
#Else
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hdc As Long, ByRef lprcClip As Any, ByVal lpfnEnum As Long, ByVal dwData As Long) As Long
#End If

Public Function GetDpi() As Long
    #If VBA7 Then
        Dim hdcScreen As LongPtr
        Dim hwnd As LongPtr
    #Else
        Dim hdcScreen As Long
        Dim hwnd As Long
    #End If
    
    hwnd = GetActiveWindow()
    hdcScreen = GetDC(hwnd)

    Dim iDPI As Long
    iDPI = -1

    If (hdcScreen) Then
        iDPI = GetDeviceCaps(hdcScreen, LOGPIXELSX)
        ReleaseDC hwnd, hdcScreen
    End If

    GetDpi = iDPI
End Function
