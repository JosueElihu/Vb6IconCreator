VERSION 5.00
Begin VB.UserControl JImageListEx 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   ClipBehavior    =   0  'None
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   Picture         =   "JImageListEx.ctx":0000
   PropertyPages   =   "JImageListEx.ctx":099A
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   30
   ToolboxBitmap   =   "JImageListEx.ctx":09B1
End
Attribute VB_Name = "JImageListEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================================================================================
'    Component  : JImageListEx
'    Autor      : J. Elihu
'    Modified   : 1.9.6 - 03/03/2022
'=====================================================================================================================

Option Explicit

Private Type GUID
  Data1   As Long
  Data2   As Integer
  Data3   As Integer
  Data4(7) As Byte
End Type

Private Type PICTDESC
  Size     As Long
  Type     As Long
  hBmp   As Long
  Data1    As Long
  Data2    As Long
End Type

'/* GDI+    */
Private Declare Function GdipDrawImageRectRectI Lib "GdiPlus" (ByVal Graphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipSetInterpolationMode Lib "GdiPlus" (ByVal Graphics As Long, ByVal InterpolationMode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "GdiPlus" (ByVal Graphics As Long, ByVal PixelOffsetMode As Long) As Long
Private Declare Function GdipGetImageDimension Lib "GdiPlus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus" (ByVal hDC As Long, ByRef Graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus" (ByVal Graphics As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus" (ByVal FileName As Long, ByRef Image As Long) As Long
Private Declare Function GdiplusStartup Lib "GdiPlus" (ByRef token As Long, ByRef lpInput As Long, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "GdiPlus" (ByVal token As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus" (ByVal Image As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "GdiPlus" (ByVal imageattr As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "GdiPlus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "GdiPlus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "GdiPlus" (ByVal imageattr As Long, ByVal ColorAdjust As Long, ByVal EnableFlag As Boolean, ByRef MatrixColor As Matrix_, ByRef MatrixGray As Matrix_, ByVal Flags As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "GdiPlus" (ByVal BITMAP As Long, hbmReturn As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal Format As Long, ByRef Scan0 As Any, ByRef BITMAP As Long) As Long
Private Declare Function GdipSaveImageToStream Lib "GdiPlus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "GdiPlus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipDrawImageRect Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GdiPlus" (ByVal BITMAP As Long, ByRef hbmReturn As Long, ByVal Background As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "GdiPlus" (ByVal Graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "GdiPlus" (ByVal Graphics As Long, ByVal dx As Single, ByVal dy As Single, ByVal Order As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus" (ByVal Graphics As Long, ByVal SmoothingMd As Long) As Long

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Any, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)

'/* DPI */
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function OleCreatePictureIndirect Lib "olepro32" (lpPictDesc As PICTDESC, riid As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long


Private Type Matrix_
  m(0 To 4, 0 To 4) As Single
End Type

Private Type ImlxData
  Strm()     As Byte
End Type

Private gdip_       As Long
Private dpi_        As Single
Private m_Item()    As ImlxData

'/* PreservePropCase */
#If False Then
   Private HIcon, HBitmap
#End If

Private Sub UserControl_Initialize()
Dim gdipSI(3) As Long

    gdipSI(0) = 1&
    Call GdiplusStartup(gdip_, gdipSI(0), ByVal 0)
    dpi_ = mvWindowsDPI
    
End Sub
Private Sub UserControl_Terminate()
    If gdip_ <> 0 Then Call GdiplusShutdown(gdip_): gdip_ = 0
    Erase m_Item
End Sub
Private Sub UserControl_Paint()
    With UserControl
        .Width = .ScaleX(30, vbPixels, vbTwips)
        .Height = .ScaleY(26, vbPixels, vbTwips)
    End With
End Sub

Property Get ImageCount() As Integer
On Error GoTo e
    ImageCount = UBound(m_Item) + 1
e:
End Property
Property Get DipScale() As Single: DipScale = dpi_: End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
Dim i           As Integer
Dim j           As Integer

    With PropBag
        j = .ReadProperty("Count", 0)
        If j > 0 Then
            j = j - 1
            ReDim m_Item(j)
            For i = 0 To j
               m_Item(i).Strm = .ReadProperty("Data_" & i)
            Next
        End If
    End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
Dim i As Integer

    With PropBag
        .WriteProperty "Count", ImageCount
        For i = 0 To ImageCount - 1
            .WriteProperty "Data_" & i, m_Item(i).Strm
        Next
    End With
    
End Sub

'?Property Page
Friend Function ppgGetData() As PropertyBag
Dim ppBag As PropertyBag
Dim i As Integer

    Set ppBag = New PropertyBag
    With ppBag
        
        .WriteProperty "Count", ImageCount
        For i = 0 To ImageCount - 1
            .WriteProperty "Data_" & i, m_Item(i).Strm
        Next
        
       Set ppgGetData = ppBag
    End With
End Function
Friend Sub ppgSetData(ppBag As PropertyBag)
Dim i As Integer
Dim j As Integer

    Erase m_Item
    With ppBag
        j = .ReadProperty("Count", 0)
        If Not j > 0 Then GoTo e
        
        ReDim m_Item(j - 1)
        For i = 0 To j - 1
            m_Item(i).Strm = .ReadProperty("Data_" & i)
        Next
    End With
e:
    PropertyChanged ""
End Sub


Property Get RawStream(ByVal Index As Long) As Byte()
On Error GoTo e
    RawStream = m_Item(Index).Strm
e:
End Property

Property Get Stream(ByVal Index As Long, Optional ByVal W As Long, Optional ByVal H As Long, Optional lColor As Long = -1, Optional ByVal Alpha As Long = 100, Optional ByVal dpiAware As Boolean = True) As Byte()
On Error GoTo e
Dim lBmpSrc As Long
Dim lBmp    As Long
Dim oStream As IUnknown
Dim lW      As Single
Dim lH      As Single
Dim eGuid   As GUID
Dim mxColor As Matrix_

    If Index < 0 Or Index > ImageCount - 1 Then Exit Property
    If Not mvLoadImage(lBmpSrc, Index, lW, lH) Then Exit Property

    If dpiAware Then W = W * dpi_: H = H * dpi_
    If W = 0 Then W = lW
    If H = 0 Then H = lH

    mvSetupMatrixColor mxColor, lColor, Alpha
    If Not mvGetResizedBmp(W, H, lBmpSrc, lW, lH, lBmp, mxColor) Then Exit Property
    Set oStream = pvStreamFromArray(0&, 0&)
    If Not oStream Is Nothing Then
        CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), eGuid
        If GdipSaveImageToStream(lBmp, oStream, eGuid, ByVal 0&) = 0& Then
            Call pvStreamToArray(ObjPtr(oStream), Stream)
        End If
    End If
    
    If lBmpSrc Then GdipDisposeImage lBmpSrc
    If lBmp Then GdipDisposeImage lBmp
e:
End Property
Property Get HBitmap(ByVal Index As Long, Optional ByVal W As Long, Optional ByVal H As Long, Optional lColor As Long = -1, Optional ByVal Alpha As Long = 100, _
                    Optional ByVal BackColor As Long = -1, Optional ByVal dpiAware As Boolean = True) As Long
On Error GoTo e
Dim eGuid   As GUID
Dim lBmpSrc As Long
Dim lBmp    As Long
Dim lW      As Single
Dim lH      As Single
Dim mxColor  As Matrix_

    If Index < 0 Or Index > ImageCount - 1 Then Exit Property
    If Not mvLoadImage(lBmpSrc, Index, lW, lH) Then Exit Property

    If dpiAware Then W = W * dpi_: H = H * dpi_
    If W = 0 Then W = lW
    If H = 0 Then H = lH

    mvSetupMatrixColor mxColor, lColor, Alpha
    If Not mvGetResizedBmp(W, H, lBmpSrc, lW, lH, lBmp, mxColor) Then Exit Property
    
    If BackColor = -1 Then BackColor = 0 Else BackColor = RGBtoARGB(BackColor, 100)
    GdipCreateHBITMAPFromBitmap lBmp, HBitmap, BackColor
    
    If lBmpSrc Then GdipDisposeImage lBmpSrc
    If lBmp Then GdipDisposeImage lBmp
e:
End Property

Property Get Picture(ByVal Index As Long, Optional ByVal W As Long, Optional ByVal H As Long, Optional lColor As Long = -1, Optional ByVal Alpha As Long = 100, _
                     Optional ByVal BackColor As Long = -1, Optional ByVal PicType As PictureTypeConstants = vbPicTypeBitmap, Optional ByVal dpiAware As Boolean = True) As StdPicture
On Error GoTo e
Dim ePic    As PICTDESC
Dim eGuid   As GUID
    
    With ePic
        .Size = Len(ePic)
        .Type = PicType
        .hBmp = IIf(PicType = vbPicTypeIcon, HIcon(Index, W, H, lColor, Alpha, dpiAware), HBitmap(Index, W, H, lColor, Alpha, BackColor, dpiAware))
    End With
    With eGuid
       .Data1 = &H20400
       .Data4(0) = &HC0
       .Data4(7) = &H46
    End With
    OleCreatePictureIndirect ePic, eGuid, True, Picture
e:
End Property

Property Get HIcon(ByVal Index As Long, Optional ByVal W As Long, Optional ByVal H As Long, Optional lColor As Long = -1, Optional ByVal Alpha As Long = 100, Optional ByVal dpiAware As Boolean = True) As Long
Dim lBmpSrc As Long
Dim lBmp    As Long
Dim lW      As Single
Dim lH      As Single
Dim mxColor As Matrix_

    If Index < 0 Or Index > ImageCount - 1 Then Exit Property
    'GoTo ResourceEx_
    
    If Not mvLoadImage(lBmpSrc, Index, lW, lH) Then Exit Property
    If dpiAware Then W = W * dpi_: H = H * dpi_

    If W = 0 Then W = lW
    If H = 0 Then H = lH

    mvSetupMatrixColor mxColor, lColor, Alpha
    If Not mvGetResizedBmp(W, H, lBmpSrc, lW, lH, lBmp, mxColor) Then Exit Property

    Call GdipCreateHICONFromBitmap(lBmp, HIcon)
    If lBmpSrc Then GdipDisposeImage lBmpSrc
    If lBmp Then GdipDisposeImage lBmp
    
    Exit Property
ResourceEx_:
    Dim Out() As Byte
    Out = Me.Stream(Index, W, H, lColor, Alpha, dpiAware)
    HIcon = CreateIconFromResourceEx(Out(0), UBound(Out) + 1&, 1&, &H30000, 0&, 0&, 0&)
End Property
Public Sub Clear(): Erase m_Item: End Sub

'Public Sub Draw(lHdc As Long, ByVal Index As Long, Left As Integer, Top As Integer)
'Dim Bmp     As Long
'Dim hGraphic   As Long
'Dim lW          As Single
'Dim lH          As Single
'
'    if mvLoadImage( Bmp, Index, lW, lH) =false then exit sub
'    If GdipCreateFromHDC(lHdc, hGraphic) = 0 Then
'        Call GdipSetInterpolationMode(hGraphic, 5&)
'        Call GdipSetPixelOffsetMode(hGraphic, 4&)
'
'        GdipDrawImageRectRectI hGraphic, Bmp, Left, Top, lW, lH, 0, 0, lW, lH, &H2, 0&, 0&, 0&
'        Call GdipDeleteGraphics(hGraphic)
'    End If
'    Call GdipDisposeImage(Bmp)
'End Sub

Public Sub Render( _
       lHdc As Long, _
       ByVal ImageIndex As Long, _
       ByVal DstX As Long, _
       ByVal DstY As Long, _
       Optional ByVal DstWidth As Long, _
       Optional ByVal DstHeight As Long, _
       Optional ByVal SrcX As Long, _
       Optional ByVal SrcY As Long, _
       Optional ByVal SrcWidth As Long, _
       Optional ByVal SrcHeight As Long, _
       Optional ByVal Alpha As Long = 100, _
       Optional ByVal lColor As Long = -1, Optional ByVal Angle As Long, Optional ByVal dpiAware As Boolean)
       
Dim Bmp         As Long
Dim hGraphic    As Long
Dim hAttributes As Long
Dim tmColor     As Matrix_
Dim tmGray      As Matrix_
Dim lW          As Single
Dim lH          As Single

    If mvLoadImage(Bmp, ImageIndex, lW, lH) = False Then Exit Sub
    If Not GdipCreateFromHDC(lHdc, hGraphic) = 0 Then Exit Sub
    If dpiAware Then DstWidth = DstWidth * dpi_: DstHeight = DstHeight * dpi_

    '/* Sizes   */
    If DstWidth = 0 Then DstWidth = lW
    If DstHeight = 0 Then DstHeight = lH
    If SrcWidth = 0 Then SrcWidth = lW
    If SrcHeight = 0 Then SrcHeight = lH
    
    Call GdipSetPixelOffsetMode(hGraphic, 4&)
    If DstWidth <> lW Or DstHeight <> lH Then
        Call GdipSetInterpolationMode(hGraphic, 7&)    '/* HighQualityBicubic  */
        Call GdipSetPixelOffsetMode(hGraphic, 4&)      '/* HALF                */
    Else
        Call GdipSetInterpolationMode(hGraphic, 5&)    '/* NearestNeighbor     */
    End If
    
    mvSetupMatrixColor tmColor, lColor, Alpha
    If Not GdipCreateImageAttributes(hAttributes) = 0 Then Exit Sub
    If Not GdipSetImageAttributesColorMatrix(hAttributes, 0, True, tmColor, tmGray, 0) = 0 Then Exit Sub
    
    If Angle = 0 Then
        GdipDrawImageRectRectI hGraphic, Bmp, DstX, DstY, DstWidth, DstHeight, SrcX, SrcY, SrcWidth, SrcHeight, &H2, hAttributes, 0&, 0&
    Else
        If GdipRotateWorldTransform(hGraphic, Angle + 180, 0) = 0 Then
            Call GdipTranslateWorldTransform(hGraphic, DstX + (DstWidth \ 2), DstY + (DstHeight \ 2), 1)
        End If
        GdipDrawImageRectRectI hGraphic, Bmp, DstWidth \ 2, DstHeight \ 2, -DstWidth, -DstHeight, SrcX, SrcY, SrcWidth, SrcHeight, &H2, hAttributes, 0&, 0&
    End If
    
    Call GdipDisposeImageAttributes(hAttributes)
    Call GdipDeleteGraphics(hGraphic)
    Call GdipDisposeImage(Bmp)
    
End Sub
Public Sub Render2( _
       Graphic As Long, _
       ByVal ImageIndex As Long, _
       ByVal DstX As Long, _
       ByVal DstY As Long, _
       Optional ByVal DstWidth As Long, _
       Optional ByVal DstHeight As Long, _
       Optional ByVal SrcX As Long, _
       Optional ByVal SrcY As Long, _
       Optional ByVal SrcWidth As Long, _
       Optional ByVal SrcHeight As Long, _
       Optional ByVal Alpha As Long = 100, _
       Optional ByVal lColor As Long = -1, Optional ByVal Angle As Long, Optional ByVal dpiAware As Boolean)
       
Dim Bmp         As Long
Dim hAttributes As Long
Dim tmColor     As Matrix_
Dim tmGray      As Matrix_
Dim lW          As Single
Dim lH          As Single

    If Graphic = 0 Then Exit Sub
    If mvLoadImage(Bmp, ImageIndex, lW, lH) = False Then Exit Sub
    If dpiAware Then DstWidth = DstWidth * dpi_: DstHeight = DstHeight * dpi_
    
    '/* Sizes   */
    If DstWidth = 0 Then DstWidth = lW
    If DstHeight = 0 Then DstHeight = lH
    If SrcWidth = 0 Then SrcWidth = lW
    If SrcHeight = 0 Then SrcHeight = lH
    
    Call GdipSetPixelOffsetMode(Graphic, 4&)
    If DstWidth <> lW Or DstHeight <> lH Then
        Call GdipSetInterpolationMode(Graphic, 7&)    '/* HighQualityBicubic  */
        Call GdipSetPixelOffsetMode(Graphic, 4&)      '/* HALF                */
    Else
        Call GdipSetInterpolationMode(Graphic, 5&)    '/* NearestNeighbor     */
    End If
    
    mvSetupMatrixColor tmColor, lColor, Alpha
    If Not GdipCreateImageAttributes(hAttributes) = 0 Then Exit Sub
    If Not GdipSetImageAttributesColorMatrix(hAttributes, 0, True, tmColor, tmGray, 0) = 0 Then Exit Sub
    
    If Angle = 0 Then
        GdipDrawImageRectRectI Graphic, Bmp, DstX, DstY, DstWidth, DstHeight, SrcX, SrcY, SrcWidth, SrcHeight, &H2, hAttributes, 0&, 0&
    Else
        If GdipRotateWorldTransform(Graphic, Angle + 180, 0) = 0 Then
            Call GdipTranslateWorldTransform(Graphic, DstX + (DstWidth \ 2), DstY + (DstHeight \ 2), 1)
        End If
        GdipDrawImageRectRectI Graphic, Bmp, DstWidth \ 2, DstHeight \ 2, -DstWidth, -DstHeight, SrcX, SrcY, SrcWidth, SrcHeight, &H2, hAttributes, 0&, 0&
    End If
    
    Call GdipDisposeImageAttributes(hAttributes)
    Call GdipDisposeImage(Bmp)
    
End Sub



'TODO: Private Subs
'=====================================================================================================================
Private Function mvLoadImage(HBitmap As Long, Index As Long, lW As Single, lH As Single) As Boolean
On Error GoTo e
Dim IStream   As IUnknown

    Set IStream = pvStreamFromArray(VarPtr(m_Item(Index).Strm(0)), UBound(m_Item(Index).Strm) + 1&)
    If Not IStream Is Nothing Then
        mvLoadImage = (GdipLoadImageFromStream(IStream, HBitmap) = 0)
    End If
    Set IStream = Nothing
    If mvLoadImage Then GdipGetImageDimension HBitmap, lW, lH
e:
End Function
Private Function mvCheckSizes(lSrcW As Single, lSrcH As Single, lNewW As Long, lNewH As Long, dpiAware As Boolean) As Boolean
Dim lW      As Single
Dim lH      As Single

    If dpiAware Then lNewW = lNewW * dpi_: lNewH = lNewH * dpi_
    If lNewW = 0 Then lNewW = lSrcW
    If lNewH = 0 Then lNewH = lSrcH
    
End Function
Private Function mvGetResizedBmp(W As Long, H As Long, SrcBmp As Long, ByVal lW As Long, ByVal lH As Long, OutBmp As Long, mxColor As Matrix_) As Boolean
Dim mxGray   As Matrix_
Dim hGrphc   As Long
Dim hAtrb    As Long

    If GdipCreateBitmapFromScan0(W, H, 0&, &HE200B, ByVal 0&, OutBmp) = 0 Then
    
        If GdipGetImageGraphicsContext(OutBmp, hGrphc) = 0 Then

            Call GdipSetInterpolationMode(hGrphc, 7&)     '/* HighQualityBicubic */
            Call GdipSetPixelOffsetMode(hGrphc, 4&)       '/* HALF               */
            'Call GdipSetSmoothingMode(hGrphc, 4&)         '/* AntiAlias          */

            If GdipCreateImageAttributes(hAtrb) = 0 Then
            
                Call GdipSetImageAttributesColorMatrix(hAtrb, 0&, True, mxColor, mxGray, 0&)
                Call GdipDrawImageRectRectI(hGrphc, SrcBmp, 0, 0, W, H, 0, 0, lW, lH, &H2, hAtrb)
                Call GdipDisposeImageAttributes(hAtrb)
                
                mvGetResizedBmp = True
            End If
            
        End If
    End If
    
    If hGrphc Then Call GdipDeleteGraphics(hGrphc): hGrphc = 0
    If Not mvGetResizedBmp Then
        If OutBmp Then Call GdipDisposeImage(OutBmp): OutBmp = 0
    End If
    
End Function

Private Sub mvSetupMatrixColor(mxColor As Matrix_, lColor As Long, mxAlpha As Long)
    With mxColor
        If lColor <> -1 Then
            Dim R As Byte, G As Byte, b As Byte
            b = ((lColor \ &H10000) And &HFF)
            G = ((lColor \ &H100) And &HFF)
            R = (lColor And &HFF)
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
        '.m(3, 3) = mxAlpha
        .m(3, 3) = mvParseAlpha(mxAlpha)
    End With
End Sub
Private Function mvWindowsDPI() As Double
Dim hDC  As Long
Dim lPx  As Double
Const LOGPIXELSX As Long = 88

    hDC = GetDC(0)
    lPx = CDbl(GetDeviceCaps(hDC, LOGPIXELSX))
    ReleaseDC 0, hDC
    If (lPx = 0) Then mvWindowsDPI = 1# Else mvWindowsDPI = lPx / 96#
    
End Function
Private Function mvParseAlpha(ByVal lAlpha As Long) As Single
    If lAlpha < 0 Then lAlpha = 0
    If lAlpha > 100 Then lAlpha = 100
    If lAlpha > 0 Then mvParseAlpha = lAlpha / 100
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

