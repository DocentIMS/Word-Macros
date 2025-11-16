Attribute VB_Name = "Tools_Colors"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
#End If
Public Type typColor
    Red As Long
    Green As Long
    Blue As Long
    Long As Long
    Hex As String
    Inverse As Long
    TooDark As Boolean
'    Hex2 As String
End Type
Function FullColor(ByVal Color) As typColor
    With FullColor
        .Long = ToLongColor(Color)
        .Hex = CleanHex(CStr(Hex(.Long)))
        .Red = (.Long And &HFF)
        .Green = (.Long \ &H100&) And &HFF&
        .Blue = (.Long \ &H10000) And &HFF&
        .Inverse = &HFFFFFF - .Long
        .TooDark = (0.2126 * .Red + 0.7152 * .Green + 0.0722 * .Blue) < 60
'        .Hex2 = (.Red * 256 * 256) + (.Green * 256) + (.Blue)
    End With
End Function
'############################################################################################
'    Other system color constants
'

'The Office Theme is not exposed in the VBA object model.
'
'You can retrieve it from the registry. The DWORD value
'
'HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\UI Theme
'
'stores the theme value (note: 16.0 is for Office 2016; for Office 2013 it is 15.0).
'
'In Office 2016, the values are:
'
'0 = Colorful
'3 = Dark Gray
'4 = Black
'5 = White
'
'In Office 2013, the values are:
'
'0 = White
'1 = Light Grey
'2 = Dark Grey
'
'(Don't you love the consistency?)
'
'Example for Office 2016:
'
'Dim strValue As String
'Dim lngTheme As Long
'strValue = "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\UI Theme"
'lngTheme = CreateObject("WScript.Shell").RegRead(strValue)
Function IsDarkModeSelected() As Boolean
    Select Case GetReg("UI Theme", "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Common") ', REG_DWORD)
    Case 2, 3, 4: IsDarkModeSelected = True
    End Select
End Function
'Function HexColorToDec(ByVal HexColor As String, Optional Switched As Boolean) As Long
'    On Error Resume Next
''    Dim RGB As Variant
''    RGB = HexToRGB(HexColor)
'    With ClrToRGB(HexColor)
'        If Switched Then
'            HexColorToDec = (.Red * 256 * 256) + (.Green * 256) + (.Blue)
'        Else
'            HexColorToDec = (.Blue * 256 * 256) + (.Green * 256) + (.Red)
'        End If
'    End With
''    HexColorToDec = (RGB(0) * 256 * 256) + (RGB(1) * 256) + (RGB(2))
'    'HexColorToDec = CDec(CleanHex(HexColor))
'    'Blue x 256 x 256 + Green x 256 + Red
'End Function
Private Function CleanHex(ByVal HexColor As String) As String
    If Left$(HexColor, 1) = "#" Then HexColor = Right$(HexColor, Len(HexColor) - 1)
    Do While Len(HexColor) < 6: HexColor = HexColor & "0": Loop
    HexColor = Right$(HexColor, 2) & Mid$(HexColor, 3, 2) & Left$(HexColor, 2)
    Select Case Left$(HexColor, 2)
    Case "0x": HexColor = "&H" & Right$(HexColor, Len(HexColor) - 2)
    Case "&H"
    Case ""
    Case Else: If Not IsNumeric(HexColor) Then HexColor = "&H" & HexColor
    End Select
    CleanHex = HexColor
End Function
Private Function ToLongColor(ByVal Color) As Long
    If Not IsNumeric(Color) Then Color = CleanHex(Color)
    ToLongColor = Color
    If ToLongColor < 0 Then ToLongColor = GetSysColor(CLng("&H" & Right(Hex(Color), 6)))
End Function
'Function TooDark(ByVal Color) As Boolean
'    With ToLongColor(Color)
'        TooDark = (0.2126 * .Red + 0.7152 * .Green + 0.0722 * .Blue) < 60 ' 128
'    End With
'End Function
'Function HexToRGB(ByVal HexColor As String) As Variant
'    Dim r As Integer
'    Dim g As Integer
'    Dim b As Integer
'    HexColor = CleanHex(HexColor)
'    r = CInt(CleanHex(Mid(HexColor, 3, 2)))
'    g = CInt(CleanHex(Mid(HexColor, 5, 2)))
'    b = CInt(CleanHex(Mid(HexColor, 7, 2)))
'    HexToRGB = Array(r, g, b)
'End Function
'const hexToRgb = (hex) =>
'  (value =>
'    value.length === 3
'      ? value.split('').map(c => parseInt(c.repeat(2), 16))
'      : value.match(/.{1,2}/g).map(v => parseInt(v, 16)))
'  (hex.replace('#', ''));
'
'// Luma - https://stackoverflow.com/a/12043228/1762224
'const isHexTooDark = (hexColor) =>
'  (([r, g, b]) =>
'    (0.2126 * r + 0.7152 * g + 0.0722 * b) < 40)
'  (hexToRgb(hexColor));
'
'// Brightness - https://stackoverflow.com/a/51567564/1762224
'const isHexTooLight = (hexColor) =>
'  (([r, g, b]) =>
'    (((r * 299) + (g * 587) + (b * 114)) / 1000) > 155)
'  (hexToRgb(hexColor));


