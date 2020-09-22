VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Class:         cGIF.cls (read only)
' Author:        Carles P.V.
' Dependencies:  cDIB.cls
'                mGIFLZWDec.bas
' Last revision: 2003.12.04
'================================================
' Notes:
' - Comments not supported.
'================================================
'
' LOG:
'
' - 2003.08.20: Fixed 'Out of stream' bug.
' - 2003.12.04: Fixed 'Unexpected block skiping'.


Option Explicit

'-- API:

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (lpDst As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)

'//

'-- Private Constants:

Private Const GIF_SIGNATURE                 As String = "GIF"
Private Const GIF_VERSION_87a               As String = "87a"
Private Const GIF_VERSION_89a               As String = "89a"
Private Const GIF_NETSCAPE_ID               As String = "NETSCAPE2.0"
Private Const GIF_TRAILER                   As Byte = &H3B

Private Const GIF_EXTENSION_INTRODUCER      As Byte = &H21
Private Const GIF_APPLICATION_EXTENSION     As Byte = &HFF
Private Const GIF_GRAPHIC_CONTROL_EXTENSION As Byte = &HF9
Private Const GIF_IMAGE_DESCRIPTOR          As Byte = &H2C
Private Const GIF_BLOCK_TERMINATOR          As Byte = &H0


'-- Private Types:

Private Type tGIFScreenDescriptor
    sdLogicalScreenWidth     As Integer
    sdLogicalScreenHeight    As Integer
    sdFlags                  As Byte
    sdBackgroundColorIndex   As Byte
    sdPixelAspectRatio       As Byte
End Type

Private Type tGIFImageDescriptor
    idImageLeftPosition      As Integer
    idImageTopPosition       As Integer
    idImageWidth             As Integer
    idImageHeight            As Integer
    idFlags                  As Byte
End Type

Private Type tGIFGraphicControlExtension
    gceBlockSize             As Byte ' = &H4
    gceFlags                 As Byte
    gceDelayTime             As Integer
    gceTransparentColorIndex As Byte
End Type

Private Type tGIFApplicationExtension
    aeBlockSize              As Byte ' = &HB
    aeApplicationIdentifier  As String * 8
    aeAuthenticationCode     As String * 3
End Type

Private Type tGIFNetscapeApplicationExtension
    nadBlockSize             As Byte ' = &HB
    nadNestcapeID            As String * 11
    nadSubBlockSize          As Byte ' = &H3
    nadFlags                 As Byte ' = &H1
    nadLoops                 As Integer
End Type

'//

Private Type tFrame
    frDIBXOR                 As New cDIB
    frDIBAND                 As New cDIB
    frLeft                   As Integer
    frTop                    As Integer
    frInterlaced             As Boolean
    frUseTransparentColor    As Boolean
    frTransparentColorIndex  As Integer
    frDelay                  As Integer
    frDisposalMethod         As Byte
    frUserInput              As Boolean
    frLocalPaletteUsed       As Boolean
    frLocalPaletteSorted     As Boolean
    frLocalPaletteEntries    As Integer
    frLocalPalette(1023)     As Byte
End Type

'-- Public Enums.:

Public Enum GIFDisposalMethod
    [dmNotSpecified]
    [dmDoNotDispose]
    [dmRestoreToBackground]
    [dmRestoreToPrevious]
End Enum

'-- Private Enums.:

Private Enum GIFflags
    '-- Animation flags
    gfGlobalPaletteExists = &H80
    gfGlobalPaletteBpp = &H70
    gfGlobalPaletteSorted = &H8
    gfGlobalPaletteSize = &H7
    '-- Image descriptor flags
    gfLocalPaletteExists = &H80
    gfLocalPaletteSorted = &H20
    gfLocalPaletteSize = &H7
    gfInterlaced = &H40
    '-- Graphic control flags
    gfDisposalMethod = &H1C
    gfUserInput = &H2
    gfUseTransparentColor = &H1
End Enum

'-- Private Property Variables:

Private m_Signature            As String * 3
Private m_Version              As String * 3
Private m_ScreenDescriptor     As tGIFScreenDescriptor
Private m_Loops                As Integer

Private m_GlobalPaletteExists  As Boolean
Private m_GlobalPaletteSorted  As Boolean
Private m_GlobalPaletteEntries As Integer
Private m_GlobalPalette(1023)  As Byte

Private m_FramesCount          As Integer
Private m_Frame()              As tFrame



'========================================================================================
' Class
'========================================================================================

Private Sub Class_Initialize()
    '-- Initialize mGIFLZWDec ext. module
    mGIFLZWDec.InitPowers
End Sub

Private Sub Class_Terminate()
    '-- Destroy current GIF
    Destroy
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function LoadFromFile(ByVal Filename As String) As Boolean

  Dim hFile         As Long
  Dim aStreamData() As Byte
  
    On Error GoTo ErrH
    
    '-- Open file...
    If (FileLen(Filename)) Then
    
        hFile = FreeFile
        Open Filename For Binary Access Read As #hFile
        '-- Get data...
        ReDim aStreamData(0 To FileLen(Filename) - 1)
        Get #hFile, , aStreamData()
        '-- Close file
        Close #hFile
        
        '-- Load (build) from loaded stream
        LoadFromFile = LoadFromStream(aStreamData())
        Exit Function
    End If
    
ErrH:
    Destroy
End Function

Public Function LoadFromStream(Stream() As Byte) As Boolean
  
  Dim tApplicationExtension    As tGIFApplicationExtension
  Dim tGraphicControlExtension As tGIFGraphicControlExtension
  Dim tFrameDescriptor         As tGIFImageDescriptor
  
  Dim lPos         As Long
  Dim lEnt         As Long
  Dim nImage       As Integer
  Dim aLZWCodeSize As Byte
  Dim aData()      As Byte
  
    '-- Destroy previous GIF
    Destroy

    '-- Check signature:
    CopyMemory ByVal m_Signature, Stream(0), 3
    If (m_Signature <> GIF_SIGNATURE) Then Exit Function
    
    '-- Check version:
    CopyMemory ByVal m_Version, Stream(3), 3
    If (m_Version <> GIF_VERSION_87a And m_Version <> GIF_VERSION_89a) Then Exit Function
    
    On Error Resume Next

    '-- Go to Screen Descriptor pos. and extract it:
    lPos = lPos + 6
    CopyMemory m_ScreenDescriptor, Stream(lPos), Len(m_ScreenDescriptor)
    
    '-- Go to Global Palette pos. and extract it [?]:
    lPos = lPos + Len(m_ScreenDescriptor)
    
    With m_ScreenDescriptor
    
        '-- Does exist [?]
        m_GlobalPaletteExists = -(pvGetFlag(.sdFlags, gfGlobalPaletteExists))
        If (m_GlobalPaletteExists) Then
        
            '-- Get palette flags:
            m_GlobalPaletteSorted = -(pvGetFlag(.sdFlags, gfGlobalPaletteSorted))
            m_GlobalPaletteEntries = 2 ^ (pvGetFlag(.sdFlags, gfGlobalPaletteSize) + 1)
            
            '-- Get palette:
            For lEnt = 0 To m_GlobalPaletteEntries - 1
                m_GlobalPalette(4 * lEnt + 0) = Stream(lPos + 3 * lEnt + 2)
                m_GlobalPalette(4 * lEnt + 1) = Stream(lPos + 3 * lEnt + 1)
                m_GlobalPalette(4 * lEnt + 2) = Stream(lPos + 3 * lEnt + 0)
            Next lEnt
            lPos = lPos + m_GlobalPaletteEntries * 3
        End If
    End With
    
    '-- Start extracting blocks...
    Do: Select Case Stream(lPos)
        
            Case GIF_EXTENSION_INTRODUCER
                
                '-- Move forward one byte (Extension label):
                lPos = lPos + 1
            
                Select Case Stream(lPos)
                
                    Case GIF_APPLICATION_EXTENSION
                 
                        '-- Move forward one byte (Extension label):
                        lPos = lPos + 1
                        
                        '-- Get structure:
                        CopyMemory tApplicationExtension, Stream(lPos), Len(tApplicationExtension)
                        lPos = lPos + Len(tApplicationExtension)
                        
                        '-- Get data sub-block/s:
                        pvGetDataSubBlocks Stream(), lPos, aData()
                        
                        '-- NETSCAPE 2.0 application extension [?]
                        With tApplicationExtension
                            If (.aeApplicationIdentifier & .aeAuthenticationCode = GIF_NETSCAPE_ID) Then
                                m_Loops = aData(2) + aData(3) * 256
                            End If
                        End With
                        
                    Case GIF_GRAPHIC_CONTROL_EXTENSION
                        
                        '-- Move forward one byte (Extension label):
                        lPos = lPos + 1
                        
                        '-- New control extension:
                        ReDim Preserve m_Frame(1 To nImage + 1)
                        
                        '-- Get structure:
                        CopyMemory tGraphicControlExtension, Stream(lPos), Len(tGraphicControlExtension)
                        lPos = lPos + Len(tGraphicControlExtension)
                        
                        '-- Store:
                        With m_Frame(nImage + 1)
                            .frDelay = tGraphicControlExtension.gceDelayTime
                            .frDisposalMethod = pvGetFlag(tGraphicControlExtension.gceFlags, gfDisposalMethod)
                            .frUserInput = pvGetFlag(tGraphicControlExtension.gceFlags, gfUserInput)
                            .frUseTransparentColor = pvGetFlag(tGraphicControlExtension.gceFlags, gfUseTransparentColor)
                            If (.frUseTransparentColor) Then
                                .frTransparentColorIndex = tGraphicControlExtension.gceTransparentColorIndex
                            End If
                        End With
                        
                    Case Else '-- NOT SUPPORTED OR UNEXPECTED extension
                        
                        '-- Move forward one byte (Extension label):
                        lPos = lPos + 1
                        '-- Skip data sub-block/s:
                        pvSkipDataSubBlocks Stream(), lPos
                        
                End Select
                
            Case GIF_IMAGE_DESCRIPTOR
                
                '-- Move forward one byte (Extension label):
                lPos = lPos + 1
                        
                '-- Image counter
                nImage = nImage + 1
                ReDim Preserve m_Frame(1 To nImage)
                
                '-- Get structure:
                CopyMemory tFrameDescriptor, Stream(lPos), Len(tFrameDescriptor)
                lPos = lPos + Len(tFrameDescriptor)
                
                With m_Frame(nImage)
                
                    '-- Image descriptor:
                    .frLeft = tFrameDescriptor.idImageLeftPosition
                    .frTop = tFrameDescriptor.idImageTopPosition
                    .frInterlaced = -(pvGetFlag(tFrameDescriptor.idFlags, gfInterlaced))
                    
                    '-- Local palette [?]:
                    .frLocalPaletteUsed = -(pvGetFlag(tFrameDescriptor.idFlags, gfLocalPaletteExists))
                    If (.frLocalPaletteUsed) Then
                    
                        '-- Get palette flags:
                        .frLocalPaletteSorted = -(pvGetFlag(tFrameDescriptor.idFlags, gfLocalPaletteSorted))
                        .frLocalPaletteEntries = 2 ^ (pvGetFlag(tFrameDescriptor.idFlags, gfLocalPaletteSize) + 1)
                        
                        '-- Get palette:
                        For lEnt = 0 To .frLocalPaletteEntries - 1
                            .frLocalPalette(4 * lEnt + 0) = Stream(lPos + 3 * lEnt + 2)
                            .frLocalPalette(4 * lEnt + 1) = Stream(lPos + 3 * lEnt + 1)
                            .frLocalPalette(4 * lEnt + 2) = Stream(lPos + 3 * lEnt + 0)
                        Next lEnt
                        lPos = lPos + .frLocalPaletteEntries * 3
                      
                      Else
                        '-- Store Global copy:
                        .frLocalPaletteSorted = m_GlobalPaletteSorted
                        .frLocalPaletteEntries = m_GlobalPaletteEntries
                        CopyMemory .frLocalPalette(0), m_GlobalPalette(0), 1024
                    End If
                    
                    '-- Create 8bpp Image (8bpp-XOR) and Mask (1bpp-AND) DIBs
                    .frDIBXOR.Create tFrameDescriptor.idImageWidth, tFrameDescriptor.idImageHeight, [08_bpp]
                    .frDIBAND.Create tFrameDescriptor.idImageWidth, tFrameDescriptor.idImageHeight, [01_bpp]
                End With

                '-- LZW Code size:
                aLZWCodeSize = Stream(lPos)
                lPos = lPos + 1
                
                '-- Get data sub-block/s (Encoded image bytes):
                pvGetDataSubBlocks Stream(), lPos, aData()
                
                '-- Build XOR and AND DIBs:
                pvDecodeFrameImage nImage, aLZWCodeSize, aData()
            
            '-- End of data:
            Case GIF_TRAILER: Exit Do
        
        End Select
        
        '-- Move forward one byte (Next extension label):
        lPos = lPos + 1
        
    Loop Until lPos > UBound(Stream)
        
    '-- Get image/s:
    m_FramesCount = nImage
    
    '-- Success:
    LoadFromStream = -1
    
On Error GoTo 0
End Function


Public Sub Destroy()

  Dim nFrm As Integer
    
    '-- Clear Signature/Version
    m_Signature = vbEmpty
    m_Version = vbEmpty
    
    '-- Clear global variables
    ZeroMemory m_ScreenDescriptor, Len(m_ScreenDescriptor)
    ZeroMemory m_GlobalPalette(0), 1024
    m_GlobalPaletteSorted = 0
    m_GlobalPaletteEntries = 0
    m_Loops = 0
    
    '-- Clear frame/s variables
    For nFrm = 1 To m_FramesCount
        Set m_Frame(nFrm).frDIBXOR = Nothing
        Set m_Frame(nFrm).frDIBAND = Nothing
    Next nFrm
    ReDim m_Frame(1 To 1)
    
    '-- Reset frames counter
    m_FramesCount = 0
End Sub

'//

Public Sub FrameDraw(ByVal hDC As Long, ByVal nFrame As Integer, Optional ByVal Left As Integer = 0, Optional ByVal Top As Integer = 0, Optional ByVal ScaleFactor As Single = 1, Optional ByVal ScreenCrop As Boolean = 0)
  
  Dim nScrW As Integer
  Dim nScrH As Integer
  
    With m_Frame(nFrame)
    
        If (ScreenCrop) Then
            '-- Crop to screen rectangle
            If (.frLeft + .frDIBXOR.Width > m_ScreenDescriptor.sdLogicalScreenWidth) Then
                nScrW = m_ScreenDescriptor.sdLogicalScreenWidth - .frLeft
              Else
                nScrW = .frDIBXOR.Width
            End If
            If (.frTop + .frDIBXOR.Height > m_ScreenDescriptor.sdLogicalScreenHeight) Then
                nScrH = m_ScreenDescriptor.sdLogicalScreenHeight - .frTop
              Else
                nScrH = .frDIBXOR.Height
            End If
          Else
            '-- Whole frame
            nScrW = .frDIBXOR.Width
            nScrH = .frDIBXOR.Height
        End If
        
        '-- Check negatives (avoids flipping)
        If (nScrW < 0) Then nScrW = 0
        If (nScrH < 0) Then nScrH = 0
        
        '-- Paint frame
        If (.frUseTransparentColor) Then
            .frDIBAND.Stretch hDC, Left + ScaleFactor * .frLeft, Top + ScaleFactor * .frTop, ScaleFactor * nScrW, ScaleFactor * nScrH, 0, 0, nScrW, nScrH, vbSrcAnd
            .frDIBXOR.Stretch hDC, Left + ScaleFactor * .frLeft, Top + ScaleFactor * .frTop, ScaleFactor * nScrW, ScaleFactor * nScrH, 0, 0, nScrW, nScrH, vbSrcPaint
          Else
            .frDIBXOR.Stretch hDC, Left + ScaleFactor * .frLeft, Top + ScaleFactor * .frTop, ScaleFactor * nScrW, ScaleFactor * nScrH, 0, 0, nScrW, nScrH
        End If
    End With
End Sub

Public Sub FrameMask(ByVal nFrame As Integer, ByVal TransparentColorIndex As Byte)
  
  Dim aPalXOR() As Byte
  Dim aPalMsk() As Byte
    
    With m_Frame(nFrame)
        
        '-- Temp. palettes
        aPalXOR() = .frLocalPalette()
        aPalMsk() = .frLocalPalette()
        
        '-- Set new transp. entry [?]
        If (.frUseTransparentColor) Then
        
            FillMemory aPalXOR(TransparentColorIndex * 4), 3, &H0
            FillMemory aPalMsk(0), 1024, &H0
            FillMemory aPalMsk(TransparentColorIndex * 4), 3, &HFF

            '-- AND DIB (Use transparent):
            .frDIBXOR.SetPalette aPalMsk()
            .frDIBAND.LoadBlt .frDIBXOR.hDC
            '-- XOR DIB:
            .frDIBXOR.SetPalette aPalXOR()
            
          Else
            '-- AND DIB (Don't use transparent):
            .frDIBAND.Cls &H0
            '-- XOR DIB:
            .frDIBXOR.SetPalette aPalXOR()
        End If
    End With
End Sub

'========================================================================================
' Properties
'========================================================================================

'----------------------------------------------------------------------------------------
' Animation
'----------------------------------------------------------------------------------------

Public Property Get AnimationLoops() As Integer
    AnimationLoops = m_Loops
End Property

'----------------------------------------------------------------------------------------
' Screen
'----------------------------------------------------------------------------------------

Public Property Get ScreenWidth() As Integer
    ScreenWidth = m_ScreenDescriptor.sdLogicalScreenWidth
End Property

Public Property Get ScreenHeight() As Integer
    ScreenHeight = m_ScreenDescriptor.sdLogicalScreenHeight
End Property

Public Property Get ScreenBackgroundColorIndex() As Byte
    ScreenBackgroundColorIndex = m_ScreenDescriptor.sdBackgroundColorIndex
End Property

Public Property Get ScreenPixelAspectRatio() As Single
    ScreenPixelAspectRatio = (m_ScreenDescriptor.sdPixelAspectRatio + 15) / 64
End Property

'----------------------------------------------------------------------------------------
' Global palette
'----------------------------------------------------------------------------------------

Public Property Get lpGlobalPalette() As Long
    lpGlobalPalette = VarPtr(m_GlobalPalette(0))
End Property

Public Property Get GlobalPaletteRGBEntry(ByVal nEntry As Byte) As Long
    If (m_GlobalPaletteExists) Then
        GlobalPaletteRGBEntry = RGB(m_GlobalPalette(nEntry * 4 + 2), m_GlobalPalette(nEntry * 4 + 1), m_GlobalPalette(nEntry * 4))
    End If
End Property

Public Property Get GlobalPaletteExists() As Boolean
    GlobalPaletteExists = m_GlobalPaletteExists
End Property

Public Property Get GlobalPaletteSorted() As Boolean
    GlobalPaletteSorted = m_GlobalPaletteSorted
End Property

Public Property Get GlobalPaletteEntries() As Integer
    GlobalPaletteEntries = m_GlobalPaletteEntries
End Property

'----------------------------------------------------------------------------------------
' Frame
'----------------------------------------------------------------------------------------

Public Property Get FramesCount() As Integer
    FramesCount = m_FramesCount
End Property

'//

Public Property Get FrameDelay(ByVal nFrame As Integer) As Integer
    FrameDelay = m_Frame(nFrame).frDelay
End Property

Public Property Get FrameInterlaced(ByVal nFrame As Integer) As Boolean
    FrameInterlaced = m_Frame(nFrame).frInterlaced
End Property

Public Property Get FrameLeft(ByVal nFrame As Integer) As Integer
    FrameLeft = m_Frame(nFrame).frLeft
End Property

Public Property Get FrameTop(ByVal nFrame As Integer) As Integer
    FrameTop = m_Frame(nFrame).frTop
End Property

Public Property Get FrameWidth(ByVal nFrame As Integer) As Integer
    FrameWidth = m_Frame(nFrame).frDIBXOR.Width
End Property

Public Property Get FrameHeight(ByVal nFrame As Integer) As Integer
    FrameHeight = m_Frame(nFrame).frDIBXOR.Height
End Property

Public Property Get FrameDisposalMethod(ByVal nFrame As Integer) As GIFDisposalMethod
    FrameDisposalMethod = m_Frame(nFrame).frDisposalMethod
End Property

Public Property Get FrameUseTransparentColor(ByVal nFrame As Integer) As Boolean
    FrameUseTransparentColor = m_Frame(nFrame).frUseTransparentColor
End Property

Public Property Get FrameTransparentColorIndex(ByVal nFrame As Integer) As Byte
    FrameTransparentColorIndex = m_Frame(nFrame).frTransparentColorIndex
End Property

Public Property Get FrameUserInput(ByVal nFrame As Integer) As Boolean
    FrameUserInput = m_Frame(nFrame).frUserInput
End Property

'----------------------------------------------------------------------------------------
' Local palette
'----------------------------------------------------------------------------------------

Public Property Get lpLocalPalette(ByVal nFrame As Integer) As Long
    lpLocalPalette = VarPtr(m_Frame(nFrame).frLocalPalette(0))
End Property

Public Property Get LocalPaletteRGBEntry(ByVal nFrame As Integer, ByVal nEntry As Byte) As Long
    With m_Frame(nFrame)
        LocalPaletteRGBEntry = RGB(.frLocalPalette(nEntry * 4 + 2), .frLocalPalette(nEntry * 4 + 1), .frLocalPalette(nEntry * 4))
    End With
End Property

Public Property Get LocalPaletteUsed(ByVal nFrame As Integer) As Boolean
    LocalPaletteUsed = m_Frame(nFrame).frLocalPaletteUsed
End Property

Public Property Get LocalPaletteSorted(ByVal nFrame As Integer) As Boolean
    LocalPaletteSorted = m_Frame(nFrame).frLocalPaletteSorted
End Property

Public Property Get LocalPaletteEntries(ByVal nFrame As Integer) As Integer
    LocalPaletteEntries = m_Frame(nFrame).frLocalPaletteEntries
End Property

'========================================================================================
' Private
'========================================================================================

Private Sub pvGetDataSubBlocks(Stream() As Byte, lPos As Long, Buffer() As Byte)
                    
  ReDim Buffer(0) As Byte
  Dim lBufferMax  As Long
  Dim aBlockSize  As Byte
    
    '-- Get data sub-blocks...
    Do
        aBlockSize = Stream(lPos)
        lBufferMax = UBound(Buffer)
        ReDim Preserve Buffer(lBufferMax + aBlockSize)
        
        lBufferMax = lBufferMax + aBlockSize
        Call CopyMemory(Buffer(lBufferMax - aBlockSize + 1), Stream(lPos + 1), aBlockSize)
        
        lPos = lPos + aBlockSize + 1
    Loop While Stream(lPos) > 0
End Sub

Private Sub pvSkipDataSubBlocks(Stream() As Byte, lPos As Long)
                    
  Dim aBlockSize As Byte
                    
    '-- Skip data sub-blocks...
    Do
        aBlockSize = Stream(lPos + 1)
        lPos = lPos + aBlockSize + 1
    Loop While Stream(lPos) > 0
End Sub

Private Function pvGetFlag(ByVal lFlags As Long, ByVal lMask As GIFflags) As Byte

    '-- Extract packed flag
    If (lMask > 0) Then
        pvGetFlag = (lFlags And lMask)
        Do While (lMask And 1) = 0
            lMask = lMask \ 2
            pvGetFlag = pvGetFlag \ 2
        Loop
    End If
End Function

'//

Private Sub pvDecodeFrameImage(ByVal nFrame As Integer, ByVal LZWCodeSize As Byte, EncodedBytes() As Byte)
  
  Dim aPalXOR() As Byte
  Dim aPalAND() As Byte
  
    With m_Frame(nFrame)
        
        '-- Set XOR palette:
        aPalXOR() = .frLocalPalette()
        If (.frUseTransparentColor) Then
            FillMemory aPalXOR(.frTransparentColorIndex * 4), 3, &H0
        End If
        .frDIBXOR.SetPalette aPalXOR()
        
        '-- Set AND palette:
        ReDim aPalAND(7) As Byte
        FillMemory aPalAND(4), 3, &HFF
        .frDIBAND.SetPalette aPalAND()
        
        '-- Decode XOR bits
        mGIFLZWDec.Decode .frDIBXOR, .frInterlaced, LZWCodeSize, EncodedBytes()
        '-- Mask frame [?]
        If (.frUseTransparentColor) Then
            FrameMask nFrame, .frTransparentColorIndex
        End If
    End With
End Sub
