VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Class:         cDIB.cls (simplified)
' Author:        Carles P.V.
' Dependencies:  None
' Last revision: 06.21.2003
'================================================

Option Explicit

'-- API:

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type BITMAPINFO_001
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(7)    As Byte
End Type

Private Type BITMAPINFO_004
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(63)   As Byte
End Type

Private Type BITMAPINFO_008
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(1023) As Byte
End Type

Private Type BITMAPINFO_RGB
    bmiHeader       As BITMAPINFOHEADER
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const DIB_RGB_COLORS As Long = &H0
Private Const COLORONCOLOR   As Long = &H3

Private Declare Function CreateDIBSection_001 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFO_001, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection_004 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFO_004, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection_008 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFO_008, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection_RGB Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFO_RGB, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, lpRGBQuad As Any) As Long
Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, lpRGBQuad As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'//

'-- Public Enums.:
Public Enum dibBPPCts
    [01_bpp] = 1
    [04_bpp] = 4
    [08_bpp] = 8
    [16_bpp] = 16
    [24_bpp] = 24
    [32_bpp] = 32
End Enum

'-- Property Variables:
Private m_hDC              As Long
Private m_hDIB             As Long
Private m_lpBits           As Long
Private m_Width            As Long
Private m_Height           As Long
Private m_BPP              As dibBPPCts
Private m_BytesPerScanline As Long
Private m_Size             As Long

'-- Private Variables:
Private m_tBIH             As BITMAPINFOHEADER
Private m_hOldDIB          As Long



'========================================================================================
' Class
'========================================================================================

Private Sub Class_Terminate()
    '-- Destroy current DIB
    Destroy
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function Create(ByVal NewWidth As Long, ByVal NewHeight As Long, ByVal Newbpp As dibBPPCts) As Boolean
    
  Dim BI_001 As BITMAPINFO_001
  Dim BI_004 As BITMAPINFO_004
  Dim BI_008 As BITMAPINFO_008
  Dim BI_RGB As BITMAPINFO_RGB
  
    '-- Destroy previous
    Destroy
    
    '-- Define DIB header
    With m_tBIH
        .biSize = Len(m_tBIH)
        .biPlanes = 1
        .biBitCount = Newbpp
        .biWidth = NewWidth
        .biHeight = -NewHeight
    End With
    Select Case Newbpp
        Case [01_bpp]: BI_001.bmiHeader = m_tBIH
        Case [04_bpp]: BI_004.bmiHeader = m_tBIH
        Case [08_bpp]: BI_008.bmiHeader = m_tBIH
        Case Else:     BI_RGB.bmiHeader = m_tBIH
    End Select
    
    '-- Set private props.
    m_Width = NewWidth
    m_Height = NewHeight
    m_BPP = Newbpp
    m_BytesPerScanline = ((m_Width * Newbpp + 31) \ 32) * 4
    m_Size = m_BytesPerScanline * m_Height
    
    '-- Create DIB and select into a DC
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        Select Case Newbpp
            Case [01_bpp]: m_hDIB = CreateDIBSection_001(m_hDC, BI_001, DIB_RGB_COLORS, m_lpBits, 0, 0)
            Case [04_bpp]: m_hDIB = CreateDIBSection_004(m_hDC, BI_004, DIB_RGB_COLORS, m_lpBits, 0, 0)
            Case [08_bpp]: m_hDIB = CreateDIBSection_008(m_hDC, BI_008, DIB_RGB_COLORS, m_lpBits, 0, 0)
            Case Else:     m_hDIB = CreateDIBSection_RGB(m_hDC, BI_RGB, DIB_RGB_COLORS, m_lpBits, 0, 0)
        End Select
        If (m_hDIB <> 0) Then
            m_hOldDIB = SelectObject(m_hDC, m_hDIB)
          Else
            Destroy
        End If
    End If
    
    '-- Success
    Create = (m_hDIB <> 0)
End Function

Public Sub Destroy()

    '-- Destroy DIB
    If (m_hDC <> 0) Then
        If (m_hDIB <> 0) Then
            SelectObject m_hDC, m_hOldDIB
            DeleteObject m_hDIB
        End If
        DeleteDC m_hDC
    End If
    
    '-- Reset DIB vars.
    m_hDC = 0
    m_hDIB = 0
    m_hOldDIB = 0
    m_lpBits = 0
    
    '-- Reset DIB props.
    m_Width = 0
    m_Height = 0
    m_BPP = 0
    m_BytesPerScanline = 0
    m_Size = 0
End Sub

'//

Public Function LoadBlt(ByVal hSrcDC As Long, Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ByVal nWidth As Long, Optional ByVal nHeight As Long) As Long
    
    If (m_hDIB <> 0) Then
        
        '-- Check bounds
        If (nWidth <= 0) Then nWidth = m_Width
        If (nHeight <= 0) Then nHeight = m_Height
        
        '-- Paint bits
        LoadBlt = BitBlt(m_hDC, x, y, nWidth, nHeight, hSrcDC, x, y, vbSrcCopy)
    End If
End Function
  
Public Function Stretch(ByVal hDstDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, Optional ByVal nSrcWidth As Long, Optional ByVal nSrcHeight As Long, Optional ByVal lROP As RasterOpConstants = vbSrcCopy) As Long
  
  Dim lMode As Long
    
    If (m_hDIB <> 0) Then
        
        '-- Check bounds
        If (nSrcWidth <= 0) Then nSrcWidth = m_Width
        If (nSrcHeight <= 0) Then nSrcHeight = m_Height
        
        '-- BitBlt/StretchBlt
        lMode = GetStretchBltMode(hDstDC)
        SetStretchBltMode hDstDC, COLORONCOLOR
        Stretch = StretchBlt(hDstDC, x, y, nWidth, nHeight, m_hDC, xSrc, ySrc, nSrcWidth, nSrcHeight, lROP)
        SetStretchBltMode hDstDC, lMode
    End If
End Function

Public Sub Cls(Optional ByVal BackColor As Long = &H0)
  
  Dim tRect  As RECT2
  Dim lColor As Long
  Dim hBrush As Long
   
    '-- Clear background (*)
    If (m_hDIB <> 0) Then
        '-- Define rect.
        SetRect tRect, 0, 0, m_Width, m_Height
        '-- Translate color
        OleTranslateColor BackColor, 0, lColor
        '-- Create brush and fill rect.
        hBrush = CreateSolidBrush(lColor)
        FillRect m_hDC, tRect, hBrush
        DeleteObject hBrush
    End If
    
' (*) Paletted formats: Windows will match given color to
'     the closest DIB palette color.
End Sub

'//

Public Sub SetPalette(Palette() As Byte)
    If (m_hDIB <> 0 And m_BPP <= [08_bpp]) Then
        SetDIBColorTable m_hDC, 0, (UBound(Palette) + 1) \ 4, Palette(0)
    End If
End Sub

Public Sub GetPalette(Palette() As Byte)
    If (m_hDIB <> 0 And m_BPP <= [08_bpp]) Then
        ReDim Palette(4 * (2 ^ m_BPP) - 1)
        GetDIBColorTable m_hDC, 0, 2 ^ m_BPP, Palette(0)
    End If
End Sub

'//

'========================================================================================
' Properties
'========================================================================================

Public Property Get hDC() As Long
    hDC = m_hDC
End Property

Public Property Get hDIB() As Long
    hDIB = m_hDIB
End Property

Public Property Get lpBits() As Long
    lpBits = m_lpBits
End Property

Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Get BPP() As dibBPPCts
    BPP = m_BPP
End Property

Public Property Get BytesPerScanline() As Long
    BytesPerScanline = m_BytesPerScanline
End Property

Public Property Get Size() As Long
    Size = m_Size
End Property
