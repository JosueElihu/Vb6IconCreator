VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : cDialog
'    Autor      : J. Elihu
'    Description: Reemplazo al CommonDialog
'--------------------------------------------------------------------------------
Option Explicit

Private Type OSVERSIONINFO
  dwSize            As Long
  dwMajorVersion    As Long
  dwMinorVersion    As Long
  dwBuildNumber     As Long
  dwPlatformId      As Long
  szCSDVersion      As String * 128
End Type

Private Type OpenFile
  StructSize      As Long
  hWnd            As Long
  hInstance       As Long
  Filter          As String
  CustomFilter    As String
  MaxCustFilter   As Long
  FilterIndex     As Long
  file            As String
  MaxFile         As Long
  FileTitle       As String
  MaxTitle        As Long
  InitialDir      As String
  Title           As String
  Flags           As Long
  FileOffset      As Integer
  FileExtension   As Integer
  DefaultExt      As String
  CustomData      As Long
  Hook            As Long
  TemplateName    As String
End Type
Private Type BrowseFolder
  hwndOwner       As Long
  pidlRoot        As Long
  pszDisplayName  As String
  lpszTitle       As String
  ulFlags         As Long
  lpfnCallback    As Long
  lParam          As Long
  iImage          As Long
End Type

Private Type ChooseColor
  lStructSize    As Long
  hwndOwner      As Long
  hInstance      As Long
  RGBResult      As Long
  lpCustColors   As String
  Flags          As Long
  lCustData      As Long
  lpfnHook       As Long
  lpTemplateName As String
End Type

'/Open File
Private Declare Function GetOpenFileName Lib "COMDLG32" Alias "GetOpenFileNameA" (pOpenfilename As OpenFile) As Long
Private Declare Function PickIconDlg Lib "Shell32" Alias "#62" (ByVal hwndOwner As Long, ByVal lpstrFile As String, ByVal nMaxFile As Long, lpdwiIndex As Long) As Long
Private Declare Function ExtractIcon Lib "Shell32" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal niIndex As Long) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" Alias "GetSaveFileNameA" (pOpenfilename As OpenFile) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Private Const OFN_EXPLORER              As Long = &H80000
Private Const OFN_LONGNAMES             As Long = &H200000
Private Const OFN_CREATEPROMPT          As Long = &H2000
Private Const OFN_ALLOWMULTISELECT      As Long = &H200
Private Const OFN_NOCHANGEDIR           As Long = &H8
Private Const OFN_OVERWRITEPROMPT       As Long = &H2
Private Const OFS_FILE_OPEN_FLAGS       As Long = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT ' Or OFN_NOCHANGEDIR
Private Const OFN_HIDEREADONLY          As Long = &H4


'/ Browser Folder
Private Declare Function SHBrowseForFolder Lib "Shell32" Alias "SHBrowseForFolderA" (lpBROWSEINFOTYPE As BrowseFolder) As Long
'Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseFolder) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pIDL As Long, ByVal pszPath As String) As Long

Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function GetVersionExA Lib "kernel32" (ByRef osv As OSVERSIONINFO) As Long

Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)


'/Choose Color
Private Declare Function ChooseColor Lib "COMDLG32" Alias "ChooseColorA" (Color As ChooseColor) As Long
Private Declare Function OleTranslateColor2 Lib "olepro32" Alias "OleTranslateColor" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long


'-----------------------------------

Private m_File()        As String
Private m_Folder        As String
Private m_FilterIndex   As Long
Private m_MultiSelect   As Boolean
Private m_Filter        As String
Private m_Title         As String
Private m_DefExt        As String
Private m_OverWrite     As Boolean
Private m_NewFolder     As Boolean
Private m_Color         As OLE_COLOR

Private Sub Class_Initialize()
Dim osv As OSVERSIONINFO

    osv.dwSize = Len(osv)
    GetVersionExA osv
    
    ReDim m_File(0)
    m_Filter = "Todos los archivos (*.*)|*.*"
    m_NewFolder = True
End Sub
Private Sub Class_Terminate()
    Erase m_File
End Sub

Property Get Filter() As String: Filter = m_Filter: End Property
Property Let Filter(ByVal Value As String): m_Filter = Value: End Property
Property Get MultiSelect() As Boolean: MultiSelect = m_MultiSelect: End Property
Property Let MultiSelect(ByVal Value As Boolean): m_MultiSelect = Value: End Property
Property Get Folder() As String: Folder = m_Folder: End Property
Property Let Folder(ByVal Value As String): m_Folder = Value: End Property
Property Get Title() As String: Title = m_Title: End Property
Property Let Title(ByVal Value As String): m_Title = Value: End Property
Property Get FilterIndex() As Long: FilterIndex = m_FilterIndex: End Property
Property Let FilterIndex(ByVal Value As Long): m_FilterIndex = Value: End Property
Property Get DefExtension() As String: DefExtension = m_DefExt: End Property
Property Let DefExtension(ByVal Value As String)
    m_DefExt = Value
    If Left$(m_DefExt, 1) <> "." And Trim$(m_DefExt) <> "" Then m_DefExt = "." & m_DefExt
End Property
Property Let OverWritePrompt(Value As Boolean): m_OverWrite = Value: End Property
Property Get OverWritePrompt() As Boolean: OverWritePrompt = m_OverWrite: End Property
Property Get NewFolderButton() As Boolean: NewFolderButton = m_NewFolder: End Property
Property Let NewFolderButton(ByVal Value As Boolean): m_NewFolder = Value: End Property
Property Get Color() As OLE_COLOR: Color = m_Color: End Property
Property Let Color(ByVal Value As OLE_COLOR): m_Color = Value: End Property


Property Get FileCount() As Long
On Error GoTo e
    FileCount = UBound(m_File) + 1
e:
End Property

Property Get FileName(Optional Index As Long) As String
On Error GoTo e
    FileName = m_File(Index)
e:
End Property

Property Let FileName(Optional Index As Long, ByVal Value As String)
On Error GoTo e
    m_File(Index) = Value
e:
End Property

Public Function ShowOpen(Optional hWnd As Long) As Boolean
Dim OF          As OpenFile
Dim tmp         As String
Dim Items()     As String
Dim i           As Long


        With OF
            .StructSize = Len(OF)
            .hWnd = hWnd
            .hInstance = App.hInstance
            .Title = m_Title
            .InitialDir = m_Folder
            .file = m_File(0) & Space$(10240) & vbNullChar & vbNullChar
            .MaxFile = Len(.file)
            .FileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
            .MaxTitle = Len(OF.FileTitle)
            .Filter = ppFilter
            .FilterIndex = m_FilterIndex
            .Flags = OFS_FILE_OPEN_FLAGS Or IIf(m_MultiSelect, OFN_ALLOWMULTISELECT, 0)
        End With
        
        
        If GetOpenFileName(OF) Then
        
            DoEvents
            If m_MultiSelect Then
            
                Items = Split(OF.file, Chr(0))
                
                If UBound(Items) > 4 Then '
                    
                    ReDim m_File(UBound(Items) - 5)
                    
                    If Right(Items(0), 1) <> "\" Then Items(0) = Items(0) & "\"
                    
                    For i = 1 To UBound(Items) - 4
                        m_File(i - 1) = Items(0) & Items(i)
                        Debug.Print m_File(i - 1)
                    Next
                    
                    m_Folder = Items(0)
                    m_Folder = GetFolder(m_File(0))
                    
                Else
                    ReDim m_File(0)
                    m_File(0) = StripNull(OF.file)
                    m_Folder = GetFolder(m_File(0))
                End If
                
            Else
                ReDim m_File(0)
                m_File(0) = StripNull(OF.file)
                m_Folder = GetFolder(m_File(0))
            End If
            
            ShowOpen = True
            
        End If
    
End Function

Public Function ShowSave(hWnd As Long) As Boolean
Dim OF     As OpenFile

    With OF
        .StructSize = Len(OF)
        .hWnd = hWnd
        .Title = m_Title
        .hInstance = App.hInstance
        .InitialDir = m_Folder
        
        If m_OverWrite Then .Flags = &H2
        
        .file = m_File(0) & String$(260 - Len(m_File(0)), 0)
        .MaxFile = 255
        .Filter = ppFilter
        .FilterIndex = m_FilterIndex
        .DefaultExt = m_DefExt
        
    End With
    
    '/ShowDialog
     If GetSaveFileName(OF) Then
        ReDim m_File(0)
        
        m_File(0) = StripNull(OF.file)
        m_Folder = GetFolder(m_File(0))
        
        ShowSave = True
     End If

End Function

Public Function ShowFolder(hWnd As Long) As Boolean
On Error Resume Next
Dim eBF     As BrowseFolder
Dim tmp     As String '* 260
Dim lPtr    As Long
Dim lPtrl   As Long

Const ZPTR = (&H0 Or &H40)

    tmp = m_Folder
    If Len(tmp) Then
        If Not Right$(tmp, 1) <> "\" Then tmp = Left$(tmp, Len(tmp) - 1)
    End If
    
    With eBF
        .hwndOwner = hWnd
        .lpszTitle = m_Title
        .ulFlags = &H1 Or IIf(m_NewFolder, &H40, 0)
        .lpfnCallback = 0&
        lPtr = LocalAlloc(ZPTR, Len(tmp) + 1)
        Call CopyMemory(ByVal lPtr, ByVal tmp, Len(tmp) + 1)
        .lParam = lPtr
    End With
    
    lPtrl = SHBrowseForFolder(eBF)
    If lPtrl Then
        tmp = String$(256, Chr(0))
        If SHGetPathFromIDList(lPtrl, tmp) Then
            m_Folder = StripNull(tmp)
            If Len(m_Folder) Then
                If Right(m_Folder, 1) <> "\" Then m_Folder = m_Folder & "\"
            End If
        End If
        
        Call CoTaskMemFree(lPtrl)
        ShowFolder = True
    Else
        ShowFolder = False
    End If
    Call LocalFree(lPtr)
    
End Function

Function ShowColor(hWnd As Long) As Boolean
Dim eCC     As ChooseColor
Dim lPtrc   As Long
Dim lbcc()  As Byte
Dim zpcc(16) As Long

    With eCC
        .lStructSize = Len(eCC)
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .lpCustColors = StrConv(lbcc, vbUnicode)
        .Flags = &H1 Or &H2 '-> RGB , Complete Dialog
        
       OleTranslateColor2 m_Color, 0, .RGBResult    '   -> If System Color?
    End With
    
    If ChooseColor(eCC) Then
        m_Color = eCC.RGBResult
        ShowColor = True
        lbcc = StrConv(eCC.lpCustColors, vbFromUnicode)
    End If
    
End Function

Private Function GetFolder(FilePath As String) As String
On Error Resume Next
    GetFolder = Left$(FilePath, InStrRev(FilePath, "\"))
End Function


Private Function ppFilter() As String
    ppFilter = Replace(m_Filter, "|", vbNullChar)
    ppFilter = Replace(ppFilter, ":", vbNullChar)
    ppFilter = ppFilter & vbNullChar & vbNullChar
End Function

Private Function StripNull(Data As String) As String
    StripNull = Left$(Data, lstrlen(Data))
End Function


