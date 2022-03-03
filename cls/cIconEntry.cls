VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIconEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------------------------------
'    Component  : cIconEntry
'    Autor      : J. Elhiu
'    Description: RW Icons File
'    Modified   : 24/02/2021
'---------------------------------------------------------------------------------------------------------------------
' Thanks to Leandro Ascierto: AlphaIcon && ClsSnapWebSite

Option Explicit

Private Type IconHeader
  Reserved      As Integer
  Type          As Integer
  Count         As Integer
End Type

Private Type IconEntry
  Width         As Byte
  Height        As Byte
  ColorCount    As Byte
  Reserved      As Byte
  Planes        As Integer
  BitCount      As Integer
  BytesInRes    As Long
  ImageOffset   As Long
End Type

Private Type MemIconEntry
  Width        As Byte
  Height       As Byte
  ColorCount   As Byte
  Reserved     As Byte
  Planes       As Integer
  BitCount     As Integer
  BytesInRes   As Long
  ID           As Integer
End Type

Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDst As Any, ByRef pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub FillMem Lib "kernel32" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal length As Long, ByVal Fill As Byte)

Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, lpName As Any, lpType As Any) As Long
Private Declare Function EnumResourceNamesByNum Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FindResourceByNum Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long


Private Declare Function StrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function StrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Private Declare Function DrawIconEx Lib "user32" (ByVal Hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Type IconBytes
   Strm() As Byte
End Type

Private m_IHeader   As IconHeader
Private m_IEntry()  As IconEntry
Private m_IBytes()  As IconBytes

'/* Enum Icons in Resource */
Private ArrResource() As String
Private z_CbMem     As Long                 'Callback allocated memory address
Private z_Cb()      As Long                 'Callback thunk array


Private Sub Class_Initialize()
'
End Sub
Private Sub Class_Terminate()
    Erase m_IEntry
    Erase m_IBytes
    
    Call zTerminate
End Sub

Property Get Icon(ByVal Index As Long) As Long
On Error GoTo e
    Icon = CreateIconFromResourceEx(m_IBytes(Index).Strm(0), m_IEntry(Index).BytesInRes, 1, &H30000, m_IEntry(Index).Width, m_IEntry(Index).Height, &H0)
e:
End Property
Property Get IconCount() As Long: IconCount = m_IHeader.Count: End Property
Property Get IconStream(ByVal Index As Long) As Byte()
On Error GoTo e
    IconStream = m_IBytes(Index).Strm()
e:
End Property
Property Get IconWidth(ByVal Index As Long) As Long
On Error GoTo e
    IconWidth = m_IEntry(Index).Width
e:
End Property
Property Get IconHeight(ByVal Index As Long) As Long
On Error GoTo e
    IconHeight = m_IEntry(Index).Height
e:
End Property

Property Get IconBitCount(ByVal Index As Long) As Long
On Error GoTo e
    IconBitCount = m_IEntry(Index).BitCount
e:
End Property
Property Get ColorDepth(ByVal Index As Long) As Long
On Error GoTo e
    If m_IEntry(Index).Planes = 1 Then
        ' 16bit, 24bit, 32bit
        ColorDepth = m_IEntry(Index).BitCount
    Else
        ' 1bit, 4bit, 8bit
        Select Case m_IEntry(Index).ColorCount
            Case 2:     ColorDepth = 1
            Case 16:    ColorDepth = 4
            Case Else:  ColorDepth = 8
        End Select
    End If
e:
End Property


Property Get Stream() As Byte()
On Error GoTo e
Dim FF      As Long
Dim Out()   As Byte
Dim i       As Long
Dim IOffset As Long
    

    m_IHeader.Type = 1 'IMAGE_ICON
    ReDim Out(6 + (m_IHeader.Count * 16) - 1)
    MemCopy Out(0), m_IHeader, 6
    
    For i = 0 To UBound(m_IEntry)
        m_IEntry(i).ImageOffset = 6 + (m_IHeader.Count * 16) + IOffset
        MemCopy Out(6 + (i * 16)), m_IEntry(i), 16
        IOffset = IOffset + UBound(m_IBytes(i).Strm) + 1
    Next
    
    ReDim Preserve Out(UBound(Out) + IOffset)
    For i = 0 To UBound(m_IBytes)
        MemCopy Out(m_IEntry(i).ImageOffset), m_IBytes(i).Strm(0), UBound(m_IBytes(i).Strm) + 1
    Next
    Stream = Out
e:
End Property


Public Function OpenIconFile(ByVal sIconFile As String) As Boolean
On Error GoTo e
Dim FF   As Long
Dim ab() As Byte
Dim i    As Long

    Erase m_IEntry
    Erase m_IBytes
        
    FF = FreeFile
    Open sIconFile For Binary Access Read As #FF
    ReDim ab(LOF(FF) - 1)
    Get #FF, , ab()
    Close #FF
    
    FillMem m_IHeader, Len(m_IHeader), 0
    Call MemCopy(m_IHeader, ab(0), Len(m_IHeader))
    
    If m_IHeader.Type = 1 And m_IHeader.Count > 0 Then
        ReDim m_IEntry(m_IHeader.Count - 1)
        ReDim m_IBytes(m_IHeader.Count - 1)
        Call MemCopy(m_IEntry(0), ab(Len(m_IHeader)), Len(m_IEntry(0)) * m_IHeader.Count)
        For i = 0 To m_IHeader.Count - 1
            ReDim m_IBytes(i).Strm(m_IEntry(i).BytesInRes - 1)
            Call MemCopy(m_IBytes(i).Strm(0), ab(m_IEntry(i).ImageOffset), m_IEntry(i).BytesInRes)
        Next
    End If
    OpenIconFile = True
    Exit Function
e:
    On Error Resume Next
    If FF Then Close #FF
End Function
Public Function SaveIconFile(ByVal sIconFile As String) As Boolean
On Error GoTo e
Dim FF      As Long
    FF = FreeFile
    Open sIconFile For Binary As FF
    Put FF, , Stream
    Close FF
    SaveIconFile = True
e:
End Function

Public Function OpenIconFromPE(ByVal sFile As String, Optional ByVal lpID As Long = 0, Optional ByVal lpName As String = "") As Boolean
Dim hLibrary    As Long
Dim hRsrc       As Long
Dim hGlobal     As Long
Dim lPtr        As Long
Dim MemIE       As MemIconEntry
Dim ID()       As Integer
Dim iBaseOffset As Long
Dim lSize       As Long
Dim bFail       As Boolean
Dim i           As Long
Dim ab()        As Byte

Const RT_ICON = 3
Const DIFFERENCE = 11
Const RT_GROUP_ICON = RT_ICON + DIFFERENCE

    
    Erase m_IEntry
    Erase m_IBytes
    ReDim ArrResource(0)
    FillMem m_IHeader, Len(m_IHeader), 0
    
    hLibrary = LoadLibraryEx(sFile, ByVal 0&, &H2&)
    If hLibrary = 0 Then Exit Function
    
    Call EnumResourceNamesByNum(hLibrary, RT_GROUP_ICON, zb_AddressOf(1, 4, 1), 0)
    
    lpName = ArrResource(lpID)
    hRsrc = FindResourceByNum(hLibrary, ByVal lpName, ByVal RT_GROUP_ICON)
    If (hRsrc = 0) Then GoTo e
    
    hGlobal = LoadResource(hLibrary, hRsrc)
    If (hGlobal = 0) Then GoTo e
    
    lPtr = LockResource(hGlobal)
    If (lPtr = 0) Then GoTo e
    
    ' Get the IconHeader:
    MemCopy m_IHeader, ByVal lPtr, Len(m_IHeader)
    If (m_IHeader.Count = 0) Then GoTo e
    
    ReDim m_IEntry(m_IHeader.Count - 1)
    ReDim ID(0 To m_IHeader.Count - 1)
    
    ' Get all the directory information into a byte array (to avoid
    ' problems with WORD alignment of structures):
    ReDim ab(Len(m_IHeader) + Len(MemIE) * m_IHeader.Count - 1) As Byte
    MemCopy ab(0), ByVal lPtr, Len(m_IHeader) + Len(MemIE) * m_IHeader.Count

    ' Loop through the entries, getting the IDs and creating a standard
    ' ICONDIRENTRY structure:
    For i = 0 To m_IHeader.Count - 1
        ' Get the MEMICONENTRY structure:
        MemCopy MemIE, ab(Len(m_IHeader) + i * Len(MemIE)), Len(MemIE)
        ' Store the icon's resource id:
        ID(i) = MemIE.ID
        ' Copy data into standard ICONDIRENTRY structure.  Note the .dwImageOffset
        ' member will be wrong at this stage:
        MemCopy m_IEntry(i), MemIE, Len(MemIE)
    Next

    
    ' Now correct the ICONENTRY byte offsets:
    iBaseOffset = Len(m_IHeader) + Len(m_IEntry(0)) * m_IHeader.Count
    m_IEntry(0).ImageOffset = iBaseOffset
    For i = 1 To m_IHeader.Count - 1
        m_IEntry(i).ImageOffset = m_IEntry(i - 1).ImageOffset + m_IEntry(i - 1).BytesInRes
    Next
    
    ' Now we have the ICONDIRENTRY structures, get the actual bits of the icons:
    ReDim m_IBytes(m_IHeader.Count - 1)
    For i = 0 To m_IHeader.Count - 1
    
        ' Load the icon with the specified resource ID:
        lpName = "#" & ID(i)
        hRsrc = FindResource(hLibrary, ByVal lpName, ByVal RT_ICON)
    
        If (hRsrc = 0) Then bFail = True: Exit For
        
        ' Load the resource:
        hGlobal = LoadResource(hLibrary, hRsrc)
        If (hGlobal = 0) Then bFail = True: Exit For

        ' Determine the size of the resource:
        lSize = SizeofResource(hLibrary, hRsrc)
    
        ' If the size is valid:
        If (lSize > 0) And (lSize = m_IEntry(i).BytesInRes) Then
        
            ' Lock the resource and get a pointer to the memory:
            lPtr = LockResource(hGlobal)
            If (lPtr = 0) Then bFail = True: Exit For

            ' Store this memory in the bitmap bits array:
            ReDim Preserve m_IBytes(i).Strm(lSize - 1)
            MemCopy m_IBytes(i).Strm(0), ByVal lPtr, lSize
        Else
            bFail = True
        End If
    
    
    Next
    
    ' Did we succeed?
    If (bFail) Then
        ' ensure clear:
        Erase m_IEntry
        Erase m_IBytes
        FillMem m_IHeader, Len(m_IHeader), 0
    End If
    OpenIconFromPE = Not (bFail)
e:
    FreeLibrary hLibrary
End Function

Public Function Add(Image As cGDIPBitmap, Optional ByVal TruncateBitCount As Boolean = True) As Boolean
On Error GoTo e
Dim mc      As Long

    mc = m_IHeader.Count
    ReDim Preserve m_IEntry(mc)
    ReDim Preserve m_IBytes(mc)
    m_IHeader.Count = mc + 1
    
    
    '/*     ATENCION    */
    '----------------------------------------------------------------------------------------
    ' TruncateBitCount deja la profundidad de bits de los iconos en cero
    '----------------------------------------------------------------------------------------
    If Not TruncateBitCount Then
        m_IEntry(mc).BitCount = 32
        m_IEntry(mc).Planes = 1
    End If

    If Image.Width < 256 Or Image.Height < 256 Then
        m_IEntry(mc).Width = Image.Width
        m_IEntry(mc).Height = Image.Height
        Call StreamImageToIconEntry(Image.Stream(Encode:="bmp"), m_IEntry(mc), m_IBytes(mc).Strm)
    Else
        m_IBytes(mc).Strm = Image.Stream
        m_IEntry(mc).BytesInRes = UBound(m_IBytes(mc).Strm) + 1
    End If
    Add = True
e:
End Function


Public Sub Truncate()
On Error GoTo e
Dim i As Long

    ' Establecer profundidad de Bits de en 0
    '-------------------------------------------
    For i = 0 To m_IHeader.Count - 1
        m_IEntry(i).BitCount = 0
        m_IEntry(i).Planes = 0
    Next
e:
End Sub

Public Function GdipBitmap(ByVal Index As Long) As cGDIPBitmap
On Error GoTo e
Dim hIcon As Long

    hIcon = CreateIconFromResourceEx(m_IBytes(Index).Strm(0), m_IEntry(Index).BytesInRes, 1, &H30000, m_IEntry(Index).Width, m_IEntry(Index).Height, &H0)
    If hIcon Then
        Set GdipBitmap = New cGDIPBitmap
        GdipBitmap.LoadImage hIcon
        DestroyIcon hIcon
    End If
e:
End Function

Public Sub DrawIcon(lHdc As Long, ByVal Index As Long, ByVal x As Long, ByVal y As Long, Optional ByVal lW As Long, Optional ByVal lH As Long)
On Error GoTo e
Dim hIcon As Long
    hIcon = CreateIconFromResourceEx(m_IBytes(Index).Strm(0), m_IEntry(Index).BytesInRes, 1, &H30000, m_IEntry(Index).Width, m_IEntry(Index).Height, &H0)
    If hIcon Then
        DrawIconEx lHdc, x, y, hIcon, lW, lH, 0, 0, &H1 Or &H2
        DestroyIcon hIcon
    End If
e:
End Sub




'TODO: Private Subs
'-----------------------------------------------------------------------------------------------------------------------
Private Function StreamImageToIconEntry(Src() As Byte, IEntry As IconEntry, Out() As Byte) As Boolean
Dim lSize       As Long

        lSize = UBound(Src) - 14                         ' 14 = len(BITMAPFILEHEADER)
        ReDim Out(lSize + AlignScan(IEntry.Width, 32))   ' + Size Mask (Empty)
        MemCopy Out(0), Src(14), lSize                   ' Remove BITMAPFILEHEADER
        MemCopy Out(8), CLng(IEntry.Height * 2), 4       ' BITMAPINFOHEADER.biHeight * 2
        
        IEntry.BytesInRes = UBound(Out) + 1
e:
End Function

Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8&
End Function


'TODO: Enum Resource icons
'-----------------------------------------------------------------------------------------------------------------------
Private Function zb_AddressOf(ByVal nOrdinal As Long, _
       ByVal nParamCount As Long, _
       Optional ByVal nThunkNo As Long = 0, _
       Optional ByVal oCallback As Object = Nothing, _
       Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
       
    '*************************************************************************************************
    '* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
    '* nParamCount  - The number of parameters that will callback
    '* nThunkNo     - Optional, allows multiple simultaneous callbacks by referencing different thunks... adjust the MAX_THUNKS Const if you need to use more than two thunks simultaneously
    '* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety   - Optional, set to false to disable IDE protection.
    '*************************************************************************************************
    
    Const MAX_FUNKS   As Long = 2                                               'Number of simultaneous thunks, adjust to taste
    Const FUNK_LONGS  As Long = 22                                              'Number of Longs in the thunk
    Const FUNK_LEN    As Long = FUNK_LONGS * 4                                  'Bytes in a thunk
    Const MEM_LEN     As Long = MAX_FUNKS * FUNK_LEN                            'Memory bytes required for the callback thunk
    Const PAGE_RWX    As Long = &H40&                                           'Allocate executable memory
    Const MEM_COMMIT  As Long = &H1000&                                         'Commit allocated memory
    Dim nAddr         As Long
  
    If nThunkNo < 0 Or nThunkNo > (MAX_FUNKS - 1) Then
        MsgBox "nThunkNo doesn't exist.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
        Exit Function
    End If
  
    If oCallback Is Nothing Then                                                'If the user hasn't specified the callback owner
        Set oCallback = Me                                                      'Then it is me
    End If
  
    nAddr = zAddressOf(oCallback, nOrdinal)                                     'Get the callback address of the specified ordinal
    If nAddr = 0 Then
        MsgBox "Callback address not found.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
        Exit Function
    End If
  
    If z_CbMem = 0 Then                                                         'If memory hasn't been allocated
        ReDim z_Cb(0 To FUNK_LONGS - 1, 0 To MAX_FUNKS - 1) As Long             'Create the machine-code array
        z_CbMem = VirtualAlloc(z_CbMem, MEM_LEN, MEM_COMMIT, PAGE_RWX)          'Allocate executable memory
    End If
  
    If z_Cb(0, nThunkNo) = 0 Then                                               'If this ThunkNo hasn't been initialized...
        z_Cb(3, nThunkNo) = _
           GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
        z_Cb(4, nThunkNo) = &HBB60E089
        z_Cb(5, nThunkNo) = VarPtr(z_Cb(0, nThunkNo))                           'Set the data address
        z_Cb(6, nThunkNo) = &H73FFC589: z_Cb(7, nThunkNo) = &HC53FF04: z_Cb(8, nThunkNo) = &H7B831F75: z_Cb(9, nThunkNo) = &H20750008: z_Cb(10, nThunkNo) = &HE883E889: z_Cb(11, nThunkNo) = &HB9905004: z_Cb(13, nThunkNo) = &H74FF06E3: z_Cb(14, nThunkNo) = &HFAE2008D: z_Cb(15, nThunkNo) = &H53FF33FF: z_Cb(16, nThunkNo) = &HC2906104: z_Cb(18, nThunkNo) = &H830853FF: z_Cb(19, nThunkNo) = &HD87401F8: z_Cb(20, nThunkNo) = &H4589C031: z_Cb(21, nThunkNo) = &HEAEBFC
    End If
  
    z_Cb(0, nThunkNo) = ObjPtr(oCallback)                                       'Set the Owner
    z_Cb(1, nThunkNo) = nAddr                                                   'Set the callback address
  
    If bIdeSafety Then                                                          'If the user wants IDE protection
        z_Cb(2, nThunkNo) = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")  'EbMode Address
    End If
    
    z_Cb(12, nThunkNo) = nParamCount                                            'Set the parameter count
    z_Cb(17, nThunkNo) = nParamCount * 4                                        'Set the number of stck bytes to release on thunk return
  
    nAddr = z_CbMem + (nThunkNo * FUNK_LEN)                                     'Calculate where in the allocated memory to copy the thunk
    RtlMoveMemory nAddr, VarPtr(z_Cb(0, nThunkNo)), FUNK_LEN                    'Copy thunk code to executable memory
    zb_AddressOf = nAddr + 16                                                   'Thunk code start address
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    Dim bSub  As Byte                                                           'Value we expect to find pointed at by a vTable method entry
    Dim bVal  As Byte
    Dim nAddr As Long                                                           'Address of the vTable
    Dim i     As Long                                                           'Loop index
    Dim j     As Long                                                           'Loop limit
  
    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                           'Get the address of the callback object's instance
    If Not zProbe(nAddr + &H1C, i, bSub) Then                                   'Probe for a Class method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
            If Not zProbe(nAddr + &H7A4, i, bSub) Then                          'Probe for a UserControl method
                Exit Function                                                   'Bail...
            End If
        End If
    End If
  
    i = i + 4                                                                   'Bump to the next entry
    j = i + 1024                                                                'Set a reasonable limit, scan 256 vTable entries
    Do While i < j
        RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
        If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4             'Return the specified vTable entry address
            Exit Do                                                             'Bad method signature, quit loop
        End If

        RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
        If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4             'Return the specified vTable entry address
            Exit Do                                                             'Bad method signature, quit loop
        End If
    
        i = i + 4                                                               'Next vTable entry
    Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
    Dim bVal    As Byte
    Dim nAddr   As Long
    Dim nLimit  As Long
    Dim nEntry  As Long
  
    nAddr = nStart                                                              'Start address
    nLimit = nAddr + 32                                                         'Probe eight entries
    Do While nAddr < nLimit                                                     'While we've not reached our probe depth
        RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
        If nEntry <> 0 Then                                                     'If not an implemented interface
            RtlMoveMemory VarPtr(bVal), nEntry, 1                               'Get the value pointed at by the vTable entry
            If bVal = &H33 Or bVal = &HE9 Then                                  'Check for a native or pcode method signature
                nMethod = nAddr                                                 'Store the vTable entry
                bSub = bVal                                                     'Store the found method signature
                zProbe = True                                                   'Indicate success
                Exit Function                                                   'Return
            End If
        End If
    
        nAddr = nAddr + 4                                                       'Next vTable entry
    Loop
End Function

Private Sub zTerminate()
    Const MEM_RELEASE As Long = &H8000&                                         'Release allocated memory flag
    If Not zb_AddressOf(1, 4, 1) = 0 Then                                       'If memory allocated
        If Not VirtualFree(z_CbMem, 0, MEM_RELEASE) = 0 Then
            z_CbMem = 0  'Release; Indicate memory released
            Erase z_Cb()
        End If
    End If
End Sub


Private Function PtrToVBString(ByVal lpszBuffer As Long) As String
Dim Buffer As String, LenBuffer As Long

    LenBuffer = StrLen(lpszBuffer)
    Buffer = String(LenBuffer + 1, 0)
    StrCpy Buffer, lpszBuffer
    PtrToVBString = Left(Buffer, LenBuffer)
End Function


Private Function ResNamesCallBack(ByVal hMod As Long, ByVal ResType As Long, ByVal ResId As Long, ByVal lParam As Long) As Long

    If (ResId > &HFFFF&) Or (ResId < 0) Then
        ArrResource(UBound(ArrResource)) = PtrToVBString(ResId)
    Else
        ArrResource(UBound(ArrResource)) = "#" & ResId
    End If

    ReDim Preserve ArrResource(UBound(ArrResource) + 1)
    ResNamesCallBack = True
    
End Function
