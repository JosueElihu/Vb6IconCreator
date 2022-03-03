VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIPBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================================================================================
'    Component  : cGDIPBitmap 1.2.2
'    Autor      : J. Elihu
'    Description: Gdip+ Bitmap Class
'    Modified   : 26/02/2022
'=====================================================================================================================

Option Explicit

Private Type COLORMATRIX
  m(0 To 4, 0 To 4) As Single
End Type
Private Type ARGB
  Blue      As Byte
  Green     As Byte
  Red       As Byte
  Alpha     As Byte
End Type
Private Type BMPdata
  Width     As Long
  Height    As Long
  stride    As Long
  PixelFormat As Long
  Scan0     As Long
  Reserved  As Long
End Type
Private Type ICONINFO
  fIcon     As Long
  xHotspot  As Long
  yHotspot  As Long
  hbmMask   As Long
  hbmColor  As Long
End Type

'/* HBitmap */
Private Type BITMAP
  bmType       As Long
  bmWidth      As Long
  bmHeight     As Long
  bmWidthBytes As Long
  bmPlanes     As Integer
  bmBitsPixel  As Integer
  bmBits       As Long
End Type
Private Type BMPIH
  lSize         As Long
  Width         As Long
  Height        As Long
  Planes        As Integer
  BitCount      As Integer
  Compression   As Long
  SizeImage     As Long
  XPelsPerMeter As Long
  YPelsPerMeter As Long
  ClrUsed       As Long
  ClrImportant  As Long
End Type
Private Type BMPI
  bmiHeader    As BMPIH
  bmiPalette   As Long ' bmiColors(3) As Byte
End Type

'/* GDI+  */
Private Declare Function GdiplusStartup Lib "GdiPlus" (ByRef token As Long, ByRef lpInput As Long, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "GdiPlus" (ByVal token As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "GdiPlus" (ByVal hGraphics As Long, ByVal BITMAP As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipSetInterpolationMode Lib "GdiPlus" (ByVal graphics As Long, ByVal Mode As GDIP__InterpolationMode) As Long
Private Declare Function GdipGetInterpolationMode Lib "GdiPlus" (ByVal graphics As Long, Mode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "GdiPlus" (ByVal graphics As Long, ByVal PixelOffSetMode As Long) As Long
Private Declare Function GdipGetPixelOffsetMode Lib "GdiPlus" (ByVal graphics As Long, PixelOffSetMode As Long) As Long
Private Declare Function GdipGetImageDimension Lib "GdiPlus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus" (ByVal Hdc As Long, ByRef graphics As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "GdiPlus" (ByVal mHicon As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus" (ByVal graphics As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus" (ByVal FileName As Long, ByRef Image As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus" (ByVal Image As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "GdiPlus" (ByVal Stream As IUnknown, ByRef Image As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal Format As Long, ByRef Scan0 As Any, ByRef BITMAP As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GdiPlus" (ByVal BITMAP As Long, ByRef hbmReturn As Long, ByVal Background As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "GdiPlus" (ByRef imageattr As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "GdiPlus" (ByVal imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "GdiPlus" (ByVal imageattr As Long, ByVal ColorAdjust As Long, ByVal EnableFlag As Boolean, ByRef MatrixColor As COLORMATRIX, ByRef MatrixGray As COLORMATRIX, ByVal Flags As Long) As Long
Private Declare Function GdipSaveImageToStream Lib "GdiPlus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipSaveImageToFile Lib "GdiPlus" (ByVal hImage As Long, ByVal sFileName As String, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "GdiPlus" (ByVal BITMAP As Long, hbmReturn As Long) As Long
Private Declare Function GdipDrawImageRect Lib "GdiPlus" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "GdiPlus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "GdiPlus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "GdiPlus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "GdiPlus" (ByVal graphics As Long, ByVal dx As Single, ByVal dy As Single, ByVal Order As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "GdiPlus" (ByVal hImage As Long, PixelFormat As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "GdiPlus" (ByVal BITMAP As Long, ByRef RECT As Long, ByVal Flags As Long, ByVal PixelFormat As Long, ByRef LockedBitmapData As BMPdata) As Long
Private Declare Function GdipBitmapUnlockBits Lib "GdiPlus" (ByVal BITMAP As Long, ByRef LockedBitmapData As BMPdata) As Long


'/* GDI */
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal Hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal Hdc As Long, pBitmapInfo As BMPI, ByVal un As Long, ByRef lplpVoid As Any, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal Hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal Hdc As Long) As Long

'/* User32 */
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal Hdc As Long) As Long
Private Declare Function LoadImageA Lib "user32" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Any, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal Hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'/* Ole32 */
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long

'/* Kernel32 */
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal length As Long, ByVal Fill As Byte)

'/* Shlwapi */
Private Declare Function PathIsURL Lib "shlwapi" Alias "PathIsURLA" (ByVal pszPath As String) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

'/* Others */
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (lpPictDesc As Any, riid As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

Private Const QM_INVALID     As Long = -1&
Private Const QM_DEFAULT     As Long = 0&
Private Const QM_LOW         As Long = 1&
Private Const QM_HIGH        As Long = 2&

Public Enum GDIP__InterpolationMode
  ImInvalid_ = QM_INVALID
  ImDefault_ = QM_DEFAULT
  ImLowQuality_ = QM_LOW
  ImHighQuality_ = QM_HIGH
  ImModeBilinear_ = QM_HIGH + 1
  ImBicubic_ = QM_HIGH + 2
  ImNearestNeighbOr_ = QM_HIGH + 3
  ImHighQualityBilinear_ = QM_HIGH + 4
  ImHighQualityBicubic_ = QM_HIGH + 5
End Enum
Public Enum GDIP__PixelOffsetMode
  PomDefault_ = 0
  PomHighSpeed_ = 1
  PomHighQuality_ = 2
  PomModeNone_ = 3
  PomModeHalf_ = 4
End Enum
Private gdip_       As Long
Private dpi_        As Single

Private m_Bmp       As Long
Private m_Width     As Single
Private m_Height    As Single

'/* PreservePropCase */
#If False Then
   Private hIcon, hBitmap
#End If

Private Sub Class_Initialize()
Dim gdipSI(3) As Long

    gdipSI(0) = 1&
    Call GdiplusStartup(gdip_, gdipSI(0), ByVal 0)
    
    dpi_ = mvWindowsDPI
End Sub
Private Sub Class_Terminate()
    Call Dispose
    If gdip_ <> 0 Then Call GdiplusShutdown(gdip_)
End Sub

Property Get Width() As Integer: Width = m_Width: End Property
Property Get Height() As Integer: Height = m_Height: End Property
Property Get hImage() As Long: hImage = m_Bmp: End Property


Public Sub Dispose()
    If m_Bmp Then
        Call GdipDisposeImage(m_Bmp)
        m_Bmp = 0
        m_Width = 0
        m_Height = 0
    End If
End Sub
Public Function LoadImage(Source As Variant) As Boolean
On Error GoTo e
Dim IStream As IUnknown
Dim Out()   As Byte

    Call Dispose
    Select Case VarType(Source)
        Case vbString
        
            If PathIsURL(Source) Then
                
            ElseIf PathFileExists(Source) Then
               
                Dim lBmp     As Long
                Dim Ext      As String
                
                Const IMAGE_BITMAP = 0
                Const IMAGE_ICON = 1
                ' Const IMAGE_CURSOR = 2
                ' Const IMAGE_ENHMETAFILE = 3
                
                Ext = LCase(Right(Source, 3))
                Select Case Ext
                    Case "bmp", "ico"
                        lBmp = LoadImageA(App.hInstance, Source, IIf(Ext = "bmp", IMAGE_BITMAP, IMAGE_ICON), 0, 0, &H10)
                        If lBmp = 0 Then Exit Function
                        LoadImage = LoadImage(lBmp)
                        If Ext = "ico" Then DestroyIcon lBmp Else DeleteObject lBmp
                        Exit Function
                    Case Else
                        Call GdipLoadImageFromFile(StrPtr(Source), m_Bmp)
                End Select
               
            End If
            
        Case (vbArray + vbByte): Out() = Source
        Case vbLong
            If GetObjectType(Source) = 7 Then ' /* OBJ_BITMAP */
                If mvIsAlphaBmp(Source) Then Out = mvStreamFromH(Source)
                If Not mvIsArrayDim(VarPtrArray(Out)) Then Call GdipCreateBitmapFromHBITMAP(Source, 0, m_Bmp)
            Else
                If mvIsAlphaBmp(Source, True) Then Out = mvStreamFromH(Source, True)
                If Not mvIsArrayDim(VarPtrArray(Out)) Then Call GdipCreateBitmapFromHICON(Source, m_Bmp)
            End If
        Case vbDataObject
            Call GdipLoadImageFromStream(Source, m_Bmp)
    End Select
    
    If m_Bmp = 0 And mvIsArrayDim(VarPtrArray(Out)) Then
        Set IStream = pvStreamFromArray(VarPtr(Out(LBound(Out))), UBound(Out) - LBound(Out) + 1&)
        If Not IStream Is Nothing Then Call GdipLoadImageFromStream(IStream, m_Bmp)
        Set IStream = Nothing
    End If
    
    If m_Bmp <> 0 Then
        GdipGetImageDimension m_Bmp, m_Width, m_Height
        LoadImage = True
    End If
e:
End Function
Public Function Resize(ByVal lNewWidth As Long, ByVal lNewHeight As Long, Optional ByVal dpiAware As Boolean) As Boolean
Dim lBmp As Long

    If m_Bmp = 0 Then Exit Function
    If dpiAware Then lNewWidth = lNewWidth * dpi_: lNewHeight = lNewHeight * dpi_
    If lNewWidth = 0 And lNewHeight = 0 Then Exit Function
   
    If mvGetResizedBmp(lNewWidth, lNewHeight, lBmp) = False Then Exit Function
    Call GdipDisposeImage(m_Bmp)
    m_Bmp = lBmp
    GdipGetImageDimension m_Bmp, m_Width, m_Height
    Resize = True
    
End Function
Public Function Save(ByVal FileName As String) As Boolean
Dim lGuid(3) As Long
Dim lBmp     As Long
Dim oStream  As IUnknown

    If m_Bmp = 0 Then Exit Function

    '- ImageCodecBMP = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
    '- ImageCodecJPG = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
    '- ImageCodecGIF = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
    '- ImageCodecTIF = "{557CF405-1A04-11D3-9A73-0000F81EF32E}"
    '- ImageCodecPNG = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
    '- ImageCodecICO = "{557CF407-1A04-11D3-9A73-0000F81EF32E}"
    
    Select Case UCase(Right(FileName, Len(FileName) - InStrRev(FileName, ".")))
        Case "PNG":  CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), lGuid(0)
        Case "JPG":  CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), lGuid(0)
        Case "JEPG": CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), lGuid(0)
        Case "BMP":  CLSIDFromString StrPtr("{557CF400-1A04-11D3-9A73-0000F81EF32E}"), lGuid(0)
        Case "ICO":  CLSIDFromString StrPtr("{557CF407-1A04-11D3-9A73-0000F81EF32E}"), lGuid(0)
        Case Else:   CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), lGuid(0)
    End Select
   Save = GdipSaveImageToFile(m_Bmp, StrConv(FileName, vbUnicode), lGuid(0), ByVal 0&) = 0&
   
End Function
Public Sub Render(lHdc As Long, Optional x As Long, Optional y As Long, Optional ByVal Width As Long, Optional ByVal Height As Long, _
                  Optional ByVal SrcX As Long, Optional ByVal SrcY As Long, Optional ByVal SrcWidth As Long, Optional ByVal SrcHeight As Long, _
                  Optional ByVal Alpha As Long = 100, Optional ByVal lColor As Long = -1, Optional InterpolationMode As GDIP__InterpolationMode = ImInvalid_, Optional PixelOffSetMode As GDIP__PixelOffsetMode)

Dim hGraphics   As Long
Dim mtColor     As COLORMATRIX
Dim mtGray      As COLORMATRIX
Dim hAttributes As Long

    
    If Width = 0 Then Width = m_Width
    If Height = 0 Then Height = m_Height
    
    If SrcWidth = 0 Then SrcWidth = m_Width
    If SrcHeight = 0 Then SrcHeight = m_Height
    
    If GdipCreateFromHDC(lHdc, hGraphics) <> 0 Then Exit Sub
    If InterpolationMode = ImInvalid_ Then
        If (Height > m_Height) And (Width > m_Width) Then
           Call GdipSetInterpolationMode(hGraphics, 5&)    '/* NearestNeighbor     */
        ElseIf (Height < m_Height) Or (Width < m_Width) Then
           Call GdipSetInterpolationMode(hGraphics, 7&)    '/* HighQualityBicubic  */
        End If
    Else
        Call GdipSetInterpolationMode(hGraphics, InterpolationMode)
    End If
    If PixelOffSetMode Then GdipSetPixelOffsetMode hGraphics, PixelOffSetMode

    mvSetupMatrixColor mtColor, lColor, Alpha
    If GdipCreateImageAttributes(hAttributes) <> 0 Then Exit Sub
    If GdipSetImageAttributesColorMatrix(hAttributes, 0, True, mtColor, mtGray, 0) <> 0 Then Exit Sub
    GdipDrawImageRectRectI hGraphics, m_Bmp, x, y, Width, Height, SrcX, SrcY, SrcWidth, SrcHeight, &H2, hAttributes
    GdipDisposeImageAttributes hAttributes
    GdipDeleteGraphics hGraphics
    
End Sub

Public Sub Render2(Graphic As Long, Optional x As Long, Optional y As Long, Optional ByVal Width As Long, Optional ByVal Height As Long, _
                  Optional ByVal SrcX As Long, Optional ByVal SrcY As Long, Optional ByVal SrcWidth As Long, Optional ByVal SrcHeight As Long, _
                  Optional ByVal Alpha As Long = 100, Optional ByVal lColor As Long = -1, Optional InterpolationMode As GDIP__InterpolationMode = ImInvalid_, Optional PixelOffSetMode As GDIP__PixelOffsetMode)

Dim mtColor     As COLORMATRIX
Dim mtGray      As COLORMATRIX
Dim hAttributes As Long

Dim Im_     As Long
Dim Pom_    As Long

    If Graphic = 0 Then Exit Sub
    If Width = 0 Then Width = m_Width
    If Height = 0 Then Height = m_Height
    If SrcWidth = 0 Then SrcWidth = m_Width
    If SrcHeight = 0 Then SrcHeight = m_Height
    
    '/* Save values */
    Im_ = ImInvalid_
    GdipGetInterpolationMode Graphic, Im_
    GdipGetPixelOffsetMode Graphic, Pom_
    
    If InterpolationMode = ImInvalid_ Then
        If (Height > m_Height) And (Width > m_Width) Then
           Call GdipSetInterpolationMode(Graphic, 5&)    '/* NearestNeighbor     */
        ElseIf (Height < m_Height) Or (Width < m_Width) Then
           Call GdipSetInterpolationMode(Graphic, 7&)    '/* HighQualityBicubic  */
        End If
    Else
        Call GdipSetInterpolationMode(Graphic, InterpolationMode)
    End If
    If PixelOffSetMode <> Pom_ Then GdipSetPixelOffsetMode Graphic, PixelOffSetMode
    
    mvSetupMatrixColor mtColor, lColor, Alpha
    If GdipCreateImageAttributes(hAttributes) <> 0 Then Exit Sub
    If GdipSetImageAttributesColorMatrix(hAttributes, 0, True, mtColor, mtGray, 0) <> 0 Then Exit Sub

    GdipDrawImageRectRectI Graphic, m_Bmp, x, y, Width, Height, SrcX, SrcY, SrcWidth, SrcHeight, &H2, hAttributes
    GdipDisposeImageAttributes hAttributes
    
    If Im_ <> ImInvalid_ Then GdipSetInterpolationMode Graphic, Im_
    If Pom_ <> PixelOffSetMode Then GdipSetPixelOffsetMode Graphic, Pom_
    
End Sub

Public Function Blur(ByVal BlurDepth As Long, Optional ByVal Color As OLE_COLOR)
On Error GoTo e
Dim Bmpd1   As BMPdata
Dim Bmpd2   As BMPdata
Dim lRct(3) As Long
Dim x       As Long
Dim y       As Long
Dim BmpShdw As Long

Dim t2xBlur     As Long
Dim Alpha       As Byte
Dim lSrcAlpha   As Long
Dim lDestAlpha  As Long
Dim Out()       As Byte
Dim Srcb()      As Byte
Dim vTally()    As Long
Dim tAlpha      As Long
Dim tColumn     As Long
Dim tAvg        As Long
Dim initY       As Long
Dim initYstop   As Long
Dim initYstart  As Long
Dim initX       As Long
Dim initXstop   As Long

Dim R As Long
Dim G As Long
Dim b As Long

    If m_Bmp = 0& Then Exit Function
    t2xBlur = BlurDepth * 2
 
    SetRect lRct(0), 0, 0, m_Width, m_Height
    ReDim Srcb(lRct(2) * 4 - 1&, lRct(3) - 1&)
    Bmpd1.Scan0 = VarPtr(Srcb(0&, 0&))
    Bmpd1.stride = 4& * lRct(2)
    Call GdipBitmapLockBits(m_Bmp, lRct(0), &H4 Or &H1, &HE200B, Bmpd1)
    
    SetRect lRct(0), 0, 0, m_Width + t2xBlur, m_Height + t2xBlur
    Call GdipCreateBitmapFromScan0(lRct(2), lRct(3), 0&, &HE200B, ByVal 0&, BmpShdw)

    ReDim Out(lRct(2) * 4 - 1&, lRct(3) - 1&)
    Bmpd2.Scan0 = VarPtr(Out(0&, 0&))
    Bmpd2.stride = 4& * lRct(2)

    Call GdipBitmapLockBits(BmpShdw, lRct(0), &H4 Or &H1 Or &H2, &HE200B, Bmpd2)
    
    R = Color And &HFF
    G = (Color \ &H100&) And &HFF
    b = (Color \ &H10000) And &HFF
    
    
    tAvg = (t2xBlur + 1) * (t2xBlur + 1)    ' how many pixels are being blurred
    ReDim vTally(0 To t2xBlur)              ' number of blur columns per pixel
    
    For y = 0 To m_Height + t2xBlur - 1     ' loop thru shadow dib
    
        FillMemory vTally(0), (t2xBlur + 1) * 4, 0  ' reset column totals
        
        If y < t2xBlur Then         ' y does not exist in source
            initYstart = 0          ' use 1st row
        Else
            initYstart = y - t2xBlur ' start n blur rows above y
        End If
        
        ' how may source rows can we use for blurring?
        If y < m_Height Then initYstop = y Else initYstop = m_Height - 1
        
        tAlpha = 0  ' reset alpha sum
        tColumn = 0    ' reset column counter
        
        ' the first n columns will all be zero
        ' only the far right blur column has values; tally them
        For initY = initYstart To initYstop
            tAlpha = tAlpha + Srcb(3, initY)
        Next
        ' assign the right column value
        vTally(t2xBlur) = tAlpha
        
        For x = 3 To (m_Width - 2) * 4 - 1 Step 4
            ' loop thru each source pixel's alpha
            
            ' set shadow alpha using blur average
            Out(x, y) = tAlpha \ tAvg
            ' and set shadow color
            Select Case Out(x, y)
            Case 255
                Out(x - 1, y) = R
                Out(x - 2, y) = G
                Out(x - 3, y) = b
            Case 0
            Case Else
                Out(x - 1, y) = R * Out(x, y) \ 255
                Out(x - 2, y) = G * Out(x, y) \ 255
                Out(x - 3, y) = b * Out(x, y) \ 255
            End Select
            ' remove the furthest left column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' count the next column of alphas
            vTally(tColumn) = 0&
            For initY = initYstart To initYstop
                vTally(tColumn) = vTally(tColumn) + Srcb(x + 4, initY)
            Next
            ' add the new column's sum to the overall sum
            tAlpha = tAlpha + vTally(tColumn)
            ' set the next column to be recalculated
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
        
        ' now to finish blurring from right edge of source
        For x = x To (m_Width + t2xBlur - 1) * 4 - 1 Step 4
            Out(x, y) = tAlpha \ tAvg
            Select Case Out(x, y)
            Case 255
                Out(x - 1, y) = R
                Out(x - 2, y) = G
                Out(x - 3, y) = b
            Case 0
            Case Else
                Out(x - 1, y) = R * Out(x, y) \ 255
                Out(x - 2, y) = G * Out(x, y) \ 255
                Out(x - 3, y) = b * Out(x, y) \ 255
            End Select
            ' remove this column's alpha sum
            tAlpha = tAlpha - vTally(tColumn)
            ' set next column to be removed
            tColumn = (tColumn + 1) Mod (t2xBlur + 1)
        Next
    Next
    
    Call GdipBitmapUnlockBits(m_Bmp, Bmpd1)
    Call GdipBitmapUnlockBits(BmpShdw, Bmpd2)
    
    Call GdipDisposeImage(m_Bmp)
    m_Bmp = BmpShdw
    GdipGetImageDimension m_Bmp, m_Width, m_Height
e:
End Function



Property Get Stream(Optional ByVal W As Long, Optional ByVal H As Long, Optional ByVal Encode As String, Optional ByVal dpiAware As Boolean = True, Optional SS) As Byte()
Dim lGuid(3) As Long
Dim lBmp     As Long
Dim oStream  As IUnknown

    If m_Bmp = 0 Then Exit Property
    If dpiAware Then W = W * dpi_: H = H * dpi_
    
    If W = 0 Then W = m_Width
    If H = 0 Then H = m_Height
    If W = m_Width And H = m_Height Then GoTo Build_
    If mvGetResizedBmp(W, H, lBmp) = False Then Exit Property
    
Build_:
    
    '- ImageCodecBMP = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
    '- ImageCodecJPG = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
    '- ImageCodecGIF = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
    '- ImageCodecTIF = "{557CF405-1A04-11D3-9A73-0000F81EF32E}"
    '- ImageCodecPNG = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
    '- ImageCodecICO = "{557CF407-1A04-11D3-9A73-0000F81EF32E}"
    
    Select Case UCase(Encode)
        Case "PNG":  CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), lGuid(0)
        Case "JPG":  CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), lGuid(0)
        Case "JEPG": CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), lGuid(0)
        Case "BMP":  CLSIDFromString StrPtr("{557CF400-1A04-11D3-9A73-0000F81EF32E}"), lGuid(0)
        Case Else:   CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), lGuid(0)
    End Select
    
    Set oStream = pvStreamFromArray(0&, 0&)
    If Not oStream Is Nothing Then
        If GdipSaveImageToStream(IIf(lBmp, lBmp, m_Bmp), oStream, lGuid(0), ByVal 0&) = 0& Then
            Call pvStreamToArray(ObjPtr(oStream), Stream)
        End If
    End If
    
    If lBmp Then GdipDisposeImage lBmp
   
End Property

Property Get hBitmap(Optional ByVal W As Long, Optional ByVal H As Long, Optional ByVal BackColor As Long = -1, Optional ByVal dpiAware As Boolean = True) As Long
Dim lBmp    As Long

    If m_Bmp = 0 Then Exit Property
    If dpiAware Then W = W * dpi_: H = H * dpi_
    
    If W = 0 Then W = m_Width
    If H = 0 Then H = m_Height
    
    If W = m_Width And H = m_Height Then GoTo Build_
    If mvGetResizedBmp(W, H, lBmp) = False Then Exit Property
    
Build_:

    If BackColor = -1 Then BackColor = 0 Else BackColor = RGBtoARGB(BackColor, 100)
    GdipCreateHBITMAPFromBitmap IIf(lBmp, lBmp, m_Bmp), hBitmap, BackColor
    
    If lBmp Then GdipDisposeImage lBmp

End Property
Property Get Picture(Optional ByVal W As Long, Optional ByVal H As Long, Optional ByVal BackColor As Long = -1, Optional ByVal PicType As PictureTypeConstants = vbPicTypeBitmap, Optional ByVal dpiAware As Boolean = True) As StdPicture
On Error GoTo e
Dim lGuid(3) As Long
Dim lPic(4)  As Long
    
    If m_Bmp = 0 Then Exit Property
    
    lPic(0) = 20
    lPic(1) = PicType
    lPic(2) = IIf(PicType = vbPicTypeIcon, hIcon(W, H, dpiAware), hBitmap(W, H, BackColor, dpiAware))
    
    ' - IPicture: {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    lGuid(0) = &H7BF80980: lGuid(1) = &H101ABF32: lGuid(2) = &HAA00BB8B: lGuid(3) = &HAB0C3000
    Call OleCreatePictureIndirect(lPic(0), lGuid(0), True, Picture)
e:
End Property

Property Get hIcon(Optional ByVal W As Long, Optional ByVal H As Long, Optional ByVal dpiAware As Boolean = True) As Long
Dim lBmp    As Long

    If m_Bmp = 0 Then Exit Property
    If dpiAware Then W = W * dpi_: H = H * dpi_
    
    If W = 0 Then W = m_Width
    If H = 0 Then H = m_Height
    
    If W = m_Width And H = m_Height Then GoTo Build_
    If mvGetResizedBmp(W, H, lBmp) = False Then Exit Property
    
Build_:
    GdipCreateHICONFromBitmap IIf(lBmp, lBmp, m_Bmp), hIcon
    If lBmp Then GdipDisposeImage lBmp
    
    'Dim Bbb() As Byte
    'Bbb() = Stream(W, H, dpiAware)
    'hIcon = CreateIconFromResourceEx(Bbb(0), UBound(Bbb) + 1&, 1&, &H30000, 0&, 0&, 0&)
End Property




'TODO: Private Subs
'=====================================================================================================================

Private Function mvStreamFromH(ByVal hBmp As Long, Optional ByVal IsIcon As Boolean) As Byte()
Dim tBmp    As BITMAP
Dim tBmpI   As BMPI

Dim lHdc    As Long
Dim lPtr    As Long
Dim hDIB    As Long
Dim OldBmp  As Long

Dim sWidth  As Single
Dim sHeight As Single

Dim lGuid(3) As Long
Dim IStrm    As IUnknown
Dim lBmp     As Long
    
    '/* Load 32bpp Bitmaps  */
    If Not IsIcon Then
        If GetGDIObject(hBmp, Len(tBmp), tBmp) = 0& Then Exit Function
    Else
        Call GdipCreateBitmapFromHICON(hBmp, lBmp)
        GdipGetImageDimension lBmp, sWidth, sHeight
        GdipDisposeImage lBmp
        lBmp = 0
        tBmp.bmWidth = sWidth
        tBmp.bmHeight = sHeight
    End If
   
    '/*Init DIB */
    Dim Bmp_ As BMPI
    With Bmp_.bmiHeader
        .lSize = Len(Bmp_.bmiHeader)
        .BitCount = 32
        .Height = tBmp.bmWidth
        .Width = tBmp.bmHeight
        .Planes = 1
        .SizeImage = .Width * .Height * 4
    End With
    
    lHdc = CreateCompatibleDC(0)
    hDIB = CreateDIBSection(lHdc, Bmp_, 0&, lPtr, 0&, 0&)
    OldBmp = SelectObject(lHdc, hDIB)
    
    '/* Process Bmp */
    If IsIcon Then Call DrawIconEx(lHdc, 0, 0, hBmp, 0, 0, 0, 0, &H1 Or &H2): GoTo Scan_
    If tBmp.bmBits = 0& Then
    
        With tBmpI.bmiHeader
            .lSize = Len(tBmpI.bmiHeader)
            .BitCount = 32
            .Height = tBmp.bmHeight
            .Width = tBmp.bmWidth
            .Planes = 1
            .SizeImage = .Width * .Height * 4
        End With
        
        '/* Flag to allow BitBlt when this failed       */
        If GetDIBits(lHdc, hBmp, 0, tBmpI.bmiHeader.Height, ByVal lPtr, tBmpI, 0) = 0 Then tBmp.bmBitsPixel = 0&
    Else
        '/* If we have a bits pointer, Simply copy data */
        CopyMemory ByVal lPtr, ByVal tBmp.bmBits, tBmp.bmWidthBytes * tBmp.bmHeight
    End If
    
Scan_:
    
    '/* Convert to Gdi+ Image */
    If lPtr Then
        Call GdipCreateBitmapFromScan0(tBmp.bmWidth, tBmp.bmHeight, tBmp.bmWidth * 4&, &HE200B, ByVal lPtr, lBmp)
        If lBmp = 0& Then GoTo out_
        GdipImageRotateFlip lBmp, 6&
    
        Set IStrm = pvStreamFromArray(0&, 0&)
        If Not IStrm Is Nothing Then
            lGuid(0) = &H557CF406: lGuid(1) = &H11D31A04: lGuid(2) = &H739A: lGuid(3) = &H2EF31EF8
            If GdipSaveImageToStream(lBmp, IStrm, lGuid(0), ByVal 0&) = 0& Then Call pvStreamToArray(ObjPtr(IStrm), mvStreamFromH)
        End If
        GdipDisposeImage lBmp
    End If
out_:

    '/*Destroy DIB */
    If hDIB = 0 Then Exit Function
    Call DeleteObject(SelectObject(lHdc, OldBmp)): hDIB = 0
    Call DeleteDC(lHdc): lHdc = 0
    
End Function

Private Function mvIsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    mvIsArrayDim = Not (lAddress = 0)
End Function
Private Sub mvSetupMatrixColor(mtColor As COLORMATRIX, mxColor As Long, mxAlpha As Long)
    With mtColor
        If mxColor <> -1 Then
            Dim R As Byte, G As Byte, b As Byte
            b = ((mxColor \ &H10000) And &HFF)
            G = ((mxColor \ &H100) And &HFF)
            R = (mxColor And &HFF)
            .m(0, 0) = R / 255
            .m(1, 0) = G / 255
            .m(2, 0) = b / 255
            .m(0, 4) = R / 255
            .m(1, 4) = G / 255
            .m(2, 4) = b / 255
        Else
            .m(0, 0) = 1
            .m(1, 1) = 1
            .m(2, 2) = 1
            .m(4, 4) = 1
        End If
        .m(3, 3) = mvParseAlpha(mxAlpha)
    End With
End Sub
Private Function mvParseAlpha(ByVal lAlpha As Long) As Single
    If lAlpha < 0 Then lAlpha = 0
    If lAlpha > 100 Then lAlpha = 100
    If lAlpha > 0 Then mvParseAlpha = lAlpha / 100
End Function
Private Function mvWindowsDPI() As Double
Dim Hdc  As Long
Dim lPx  As Double
    Hdc = GetDC(0)
    lPx = CDbl(GetDeviceCaps(Hdc, 88))
    ReleaseDC 0, Hdc
    If (lPx = 0) Then mvWindowsDPI = 1# Else mvWindowsDPI = lPx / 96#
End Function
Private Function mvIsAlphaBmp(ByVal hBmp As Long, Optional ByVal IsIcon As Boolean) As Boolean
Dim ARGB()      As ARGB
Dim lRct(3)     As Long
Dim BMPdata     As BMPdata
Dim lPixelFormat As Long
Dim lx          As Long
Dim ly          As Long
Dim sWidth      As Single
Dim sHeight     As Single
Dim Bmp         As Long
Dim II_         As ICONINFO

    If hBmp = 0 Then Exit Function
    
    If IsIcon Then
        If GetIconInfo(hBmp, II_) = 0 Then Exit Function
        If II_.hbmColor = 0 Then GoTo e
    End If
    
    If GdipCreateBitmapFromHBITMAP(IIf(IsIcon, II_.hbmColor, hBmp), 0&, Bmp) <> 0 Then GoTo e
    If GdipGetImagePixelFormat(Bmp, lPixelFormat) <> 0 Then GoTo e
    If lPixelFormat <> &H22009 Then GoTo e '/* PixelFormat32bppRGB */
    If GdipGetImageDimension(Bmp, sWidth, sHeight) <> 0 Then GoTo e
    
    lRct(2) = CLng(sWidth)
    lRct(3) = CLng(sHeight)
    
    ReDim ARGB(lRct(2) - 1&, lRct(3) - 1&)
    BMPdata.Scan0 = VarPtr(ARGB(0&, 0&))
    BMPdata.stride = 4& * lRct(2)
    If GdipBitmapLockBits(Bmp, lRct(0), &H1 Or &H4, lPixelFormat, BMPdata) <> 0 Then GoTo e
    
    For ly = 0 To BMPdata.Height - 1
        For lx = 0 To BMPdata.Width - 1
            If ARGB(lx, ly).Alpha > 0 Then
                If ARGB(lx, ly).Alpha < 255 Then mvIsAlphaBmp = True: Exit For
            End If
        Next
        If mvIsAlphaBmp Then Exit For
    Next
    Call GdipBitmapUnlockBits(Bmp, BMPdata)
e:
    If Bmp Then Call GdipDisposeImage(Bmp)
    If II_.hbmColor Then Call DeleteObject(II_.hbmColor)
    If II_.hbmMask Then Call DeleteObject(II_.hbmMask)
    
End Function

Private Function RGBtoARGB(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
    'By LaVople
    ' GDI+ color conversion routines. Most GDI+ functions require ARGB format vs standard RGB format
    ' This routine will return the passed RGBcolor to RGBA format
    ' Passing VB system color constants is allowed, i.e., vbButtonFace
    ' Pass Opacity as a value from 0 to 255

    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    RGBtoARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
    
End Function

Private Function mvGetResizedBmp(W As Long, H As Long, OutBmp As Long) As Boolean
Dim hGrphc   As Long

    If GdipCreateBitmapFromScan0(W, H, 0&, &HE200B, ByVal 0&, OutBmp) = 0 Then
        If GdipGetImageGraphicsContext(OutBmp, hGrphc) = 0 Then
           Call GdipSetInterpolationMode(hGrphc, 7&)     '/* HighQualityBicubic */
           Call GdipSetPixelOffsetMode(hGrphc, 4&)       '/* HALF               */
           Call GdipSetSmoothingMode(hGrphc, 4&)         '/* AntiAlias          */
           
           GdipDrawImageRect hGrphc, m_Bmp, 0, 0, W, H
           mvGetResizedBmp = True
        End If
    End If
    If Not mvGetResizedBmp Then
        If OutBmp Then Call GdipDisposeImage(OutBmp): OutBmp = 0
    End If
    
    If hGrphc Then GdipDeleteGraphics hGrphc
    
End Function

Private Function pvStreamFromArray(ArrayPtr As Long, length As Long) As stdole.IUnknown
On Error GoTo e
Dim o_hMem As Long
Dim o_lpMem  As Long
     
    If ArrayPtr = 0& Then
        CreateStreamOnHGlobal 0&, 1&, pvStreamFromArray
    ElseIf length <> 0& Then
        o_hMem = GlobalAlloc(&H2&, length)
        If o_hMem <> 0 Then
            o_lpMem = GlobalLock(o_hMem)
            If o_lpMem <> 0 Then
                CopyMemory ByVal o_lpMem, ByVal ArrayPtr, length
                Call GlobalUnlock(o_hMem)
                Call CreateStreamOnHGlobal(o_hMem, 1&, pvStreamFromArray)
            End If
        End If
    End If
e:
End Function
Private Function pvStreamToArray(hStream As Long, arrayBytes() As Byte) As Boolean
Dim o_hMem        As Long
Dim o_lpMem       As Long
Dim o_lByteCount  As Long
    
    If hStream Then
        If GetHGlobalFromStream(ByVal hStream, o_hMem) = 0 Then
            o_lByteCount = GlobalSize(o_hMem)
            If o_lByteCount > 0 Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem <> 0 Then
                    ReDim arrayBytes(0 To o_lByteCount - 1)
                    CopyMemory arrayBytes(0), ByVal o_lpMem, o_lByteCount
                    GlobalUnlock o_hMem
                    pvStreamToArray = True
                End If
            End If
        End If
        
    End If
End Function
