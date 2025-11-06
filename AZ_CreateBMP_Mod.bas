Attribute VB_Name = "AZ_CreateBMP_Mod"
'*******************************
' // This code Sets the BackColor of
' // Pages on a Multipage Control.(Excel)
'*******************************
Option Explicit
 
'=============================
' // Private Declarations..
'=============================
 
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
 
Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As LongPtr
End Type
 
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biRUsed As Long
    biRImportant As Long
End Type
 
' A BITMAPINFO structure for bitmaps with no color palette.
Private Type BITMAPINFO_NoColors
    bmiHeader As BITMAPINFOHEADER
End Type
 
Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type
 
Private Type MemoryBitmap
    hdc As LongPtr
    hbm As LongPtr
    oldhDC As LongPtr
    wid As Long
    hgt As Long
    bitmap_info As BITMAPINFO_NoColors
End Type
'#If VBA7 Then
''    #If Win64 Then
'    Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" _
'    (ByVal hdc As Long) _
'    As Long
'
'    Private Declare PtrSafe Function SelectObject Lib "gdi32" _
'    (ByVal hdc As Long, ByVal hObject As Long) _
'    As Long
'
'    Private Declare PtrSafe Function DeleteDC Lib "gdi32" _
'    (ByVal hdc As Long) _
'    As Long
'
'    Private Declare PtrSafe Function DeleteObject Lib "gdi32" _
'    (ByVal hObject As Long) _
'    As Long
'
'    Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" _
'    (ByVal hdc As Long, pBitmapInfo As BITMAPINFO_NoColors, _
'    ByVal un As Long, ByVal lplpVoid As Long, _
'    ByVal handle As Long, ByVal dw As Long) _
'    As Long
'
'    Private Declare PtrSafe Function GetDIBits Lib "gdi32" _
'    (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal _
'    nStartScan As Long, ByVal nNumScans As Long, _
'    lpBits As Any, lpBI As BITMAPINFO_NoColors, _
'    ByVal wUsage As Long) _
'    As Long
'
'    Private Declare PtrSafe Function SetRect Lib "user32" _
'    (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, _
'    ByVal X2 As Long, ByVal Y2 As Long) As Long
'
'    Private Declare PtrSafe Function SetBkMode Lib "gdi32.dll" _
'    (ByVal hdc As Long, ByVal nBkMode As Long) _
'    As Long
'
'    Private Declare PtrSafe Function CreateBrushIndirect Lib "gdi32" _
'    (lpLogBrush As LOGBRUSH) As Long
'    Private Declare PtrSafe Function FillRect Lib "user32" _
'    (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'#Else
'    Private Declare Function CreateCompatibleDC Lib "gdi32" _
'    (ByVal hdc As Long) _
'    As Long
'
'    Private Declare Function SelectObject Lib "gdi32" _
'    (ByVal hdc As Long, ByVal hObject As Long) _
'    As Long
'
'    Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'
'    Private Declare Function DeleteObject Lib "gdi32" _
'    (ByVal hObject As Long) _
'    As Long
'
'    Private Declare Function CreateDIBSection Lib "gdi32" _
'    (ByVal hdc As Long, pBitmapInfo As BITMAPINFO_NoColors, _
'    ByVal un As Long, ByVal lplpVoid As Long, _
'    ByVal handle As Long, ByVal dw As Long) _
'    As Long
'
'    Private Declare Function GetDIBits Lib "gdi32" _
'    (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal _
'    nStartScan As Long, ByVal nNumScans As Long, _
'    lpBits As Any, lpBI As BITMAPINFO_NoColors, _
'    ByVal wUsage As Long) _
'    As Long
'
'    Private Declare Function SetRect Lib "user32" _
'    (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, _
'    ByVal X2 As Long, ByVal Y2 As Long) As Long
'
'    Private Declare Function SetBkMode Lib "gdi32.dll" _
'    (ByVal hdc As Long, ByVal nBkMode As Long) _
'    As Long
'
'    Private Declare Function CreateBrushIndirect Lib "gdi32" _
'    (lpLogBrush As LOGBRUSH) As Long
'    Private Declare Function FillRect Lib "user32" _
'    (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'#End If
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "User32.dll" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
    Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
    Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" (ByVal hdc As LongPtr, pBitmapInfo As BITMAPINFO_NoColors, _
        ByVal un As Long, ByVal lplpVoid As LongPtr, ByVal handle As LongPtr, ByVal dw As Long) As LongPtr
    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Private Declare PtrSafe Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As LongPtr
    Private Declare PtrSafe Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare PtrSafe Function SetBkMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal nBkMode As Long) As Long
    Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hdc As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
    Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, _
        ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO_NoColors, ByVal wUsage As Long) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long

  #Else
    Private Declare Function OpenClipboard Lib "User32.dll" (ByVal hwnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "User32.dll" () As Long
    Private Declare Function CloseClipboard Lib "User32.dll" () As Long
    Private Declare Function IsClipboardFormatAvailable Lib "User32.dll" (ByVal wFormat As Long) As Long
    Private Declare Function GetClipboardData Lib "User32.dll" (ByVal wFormat As Long) As Long
    Private Declare Function SetClipboardData Lib "User32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
    Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Any) As Long
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO_NoColors, _
        ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
    Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
    Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
    Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
    Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
        ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO_NoColors, ByVal wUsage As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
  #End If
Private Const DIB_RGB_COLORS = 0&
Private Const BI_RGB = 0&
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const POINTSPERINCH As Long = 72
'=============================
' // Public Routines.
'=============================
Public Function Generate1ColorBMP(Color As Long, Optional Width As Single = 50, Optional Height As Single = 50) As IPicture
    Dim sBMPFile As String
    sBMPFile = Environ("Temp") & "\Temp.bmp"
    Dim memory_bitmap As MemoryBitmap
 
    ' Create the memory bitmap.
    memory_bitmap = MakeMemoryBitmap(Width, Height)
 
    ' Draw on the bitmap.
    DrawOnMemoryBitmap memory_bitmap, Color
 
    ' Save the bmp.
    SaveMemoryBitmap memory_bitmap, sBMPFile
 
    ' load the bmp onto the page.
    Set Generate1ColorBMP = LoadPicture(sBMPFile)
 
    ' Delete the memory bitmap.
    DeleteMemoryBitmap memory_bitmap
 
    ' Delete BMP file.
    Kill sBMPFile
 
End Function

Public Sub SetBackColor(Page As MSForms.Page, Color As Long)
    Dim sBMPFile As String
    sBMPFile = Environ("Temp") & "\Temp.bmp"
    Dim memory_bitmap As MemoryBitmap
 
    ' Create the memory bitmap.
    memory_bitmap = MakeMemoryBitmap(Page.Parent.Parent.Width, Page.Parent.Parent.Height)
    
    ' Draw on the bitmap.
    DrawOnMemoryBitmap memory_bitmap, Color
 
    ' Save the bmp.
    Call SaveMemoryBitmap(memory_bitmap, sBMPFile)
 
    ' load the bmp onto the page.
    Set Page.Picture = LoadPicture(sBMPFile)
 
    ' Delete the memory bitmap.
    DeleteMemoryBitmap memory_bitmap
 
    ' Delete BMP file.
    Kill sBMPFile
 
End Sub

 
 
 
'=============================
' // Private Routines.
'=============================
 
' Make a memory bitmap according to the MultiPage size.
Private Function MakeMemoryBitmap(Width As Single, Height As Single) As MemoryBitmap
 
    Dim result As MemoryBitmap
    Dim bytes_per_scanLine As Long
    Dim pad_per_scanLine As Long
    Dim new_font As Long
 
    ' Create the device context.
    result.hdc = CreateCompatibleDC(0)
 
 
    ' Define the bitmap.
    With result.bitmap_info.bmiHeader
        .biBitCount = 32
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(result.bitmap_info.bmiHeader)
        .biWidth = Width 'Page.Parent.Parent.Width 'wid
        .biHeight = Height 'Page.Parent.Parent.Height ' hgt
        bytes_per_scanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
        pad_per_scanLine = bytes_per_scanLine - (((.biWidth * .biBitCount) + 7) \ 8)
        .biSizeImage = bytes_per_scanLine * Abs(.biHeight)
    End With
 
    ' Create the bitmap.
    result.hbm = CreateDIBSection( _
    result.hdc, result.bitmap_info, _
    DIB_RGB_COLORS, ByVal 0&, _
    ByVal 0&, ByVal 0&)
 
    ' Make the device context use the bitmap.
    result.oldhDC = SelectObject(result.hdc, result.hbm)
 
    ' Return the MemoryBitmap structure.
    result.wid = Width ' Page.Parent.Parent.Width
    result.hgt = Height 'Page.Parent.Parent.Height
 
    MakeMemoryBitmap = result
 
End Function
 
Private Sub DrawOnMemoryBitmap(memory_bitmap As MemoryBitmap, Color As Long)
 
   Dim lb As LOGBRUSH, tRect As RECT
   Dim hBrush As LongPtr
 
   lb.lbColor = Color
 
   'Create a new brush
    hBrush = CreateBrushIndirect(lb)
    With memory_bitmap
       SetRect tRect, 0, 0, .wid, .hgt
    End With
 
    SetBkMode memory_bitmap.hdc, 2 'Opaque
 
    'Paint the mem dc.
    FillRect memory_bitmap.hdc, tRect, hBrush
 
End Sub
 
' Save the memory bitmap into a bitmap file.
Private Sub SaveMemoryBitmap(memory_bitmap As MemoryBitmap, ByVal file_name As String)
 
    Dim bitmap_file_header As BITMAPFILEHEADER
    Dim fnum As Integer
    Dim pixels() As Byte
 
    ' Fill in the BITMAPFILEHEADER.
    With bitmap_file_header
        .bfType = &H4D42   ' "BM"
        .bfOffBits = Len(bitmap_file_header) + _
        Len(memory_bitmap.bitmap_info.bmiHeader)
        .bfSize = .bfOffBits + _
        memory_bitmap.bitmap_info.bmiHeader.biSizeImage
    End With
 
    ' Open the output bitmap file.
    fnum = FreeFile
    Open file_name For Binary As fnum
    ' Write the BITMAPFILEHEADER.
    Put #fnum, , bitmap_file_header
    ' Write the BITMAPINFOHEADER.
    ' (Note that memory_bitmap.bitmap_info.bmiHeader.biHeight
    ' must be positive for this.)
    Put #fnum, , memory_bitmap.bitmap_info
    ' Get the DIB bits.
    ReDim pixels(1 To 4, _
    1 To memory_bitmap.wid, _
    1 To memory_bitmap.hgt)
    GetDIBits memory_bitmap.hdc, memory_bitmap.hbm, _
    0, memory_bitmap.hgt, pixels(1, 1, 1), _
    memory_bitmap.bitmap_info, DIB_RGB_COLORS
    ' Write the DIB bits.
    Put #fnum, , pixels
    ' Close the file.
    Close fnum
 
End Sub
 
' Delete the bitmap and free its resources.
Private Sub DeleteMemoryBitmap(memory_bitmap As MemoryBitmap)
 
    SelectObject memory_bitmap.hdc, memory_bitmap.oldhDC
    DeleteObject memory_bitmap.hbm
    DeleteDC memory_bitmap.hdc
 
End Sub



