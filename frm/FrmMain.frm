VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IconCreator"
   ClientHeight    =   4845
   ClientLeft      =   8865
   ClientTop       =   4725
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   516
   Begin VB.PictureBox PicPrev 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   1800
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   5775
   End
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin IconCreator.JImageListEx imlMnu 
      Left            =   480
      Top             =   4080
      _ExtentX        =   794
      _ExtentY        =   688
      Count           =   10
      Data_0          =   "FrmMain.frx":26438
      Data_1          =   "FrmMain.frx":26826
      Data_2          =   "FrmMain.frx":27AC1
      Data_3          =   "FrmMain.frx":28333
      Data_4          =   "FrmMain.frx":28B6F
      Data_5          =   "FrmMain.frx":29EB2
      Data_6          =   "FrmMain.frx":2AF6C
      Data_7          =   "FrmMain.frx":2CFDF
      Data_8          =   "FrmMain.frx":2DA20
      Data_9          =   "FrmMain.frx":2F783
   End
   Begin VB.Menu mnu 
      Caption         =   "Archivo"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "Nuevo"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Abrir"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Guardar"
         Index           =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Guardar como"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Crear desde imagen"
         Index           =   5
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "Bitmaps"
      Index           =   1
      Begin VB.Menu mnuBmp 
         Caption         =   "Añadir"
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuBmp 
         Caption         =   "Eliminar"
         Index           =   1
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuBmp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuBmp 
         Caption         =   "Copiar"
         Index           =   3
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuBmp 
         Caption         =   "Pegar"
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuBmp 
         Caption         =   "Duplicar"
         Index           =   5
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuBmp 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuBmp 
         Caption         =   "Redimencionar"
         Index           =   7
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuBmp 
         Caption         =   "Duplicar y redimencionar"
         Index           =   8
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuBmp 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuBmp 
         Caption         =   "Añadir desde recursos"
         Index           =   10
      End
      Begin VB.Menu mnuBmp 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuBmp 
         Caption         =   "Exportar bitmap"
         Index           =   12
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "Ayuda"
      Index           =   3
      Begin VB.Menu mnuHelp 
         Caption         =   "Acerca de..."
         Index           =   0
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function IconDialog Lib "Shell32" Alias "#62" (ByVal hWndOwner As Long, ByVal lpstrFile As String, ByVal nMaxFile As Long, lpdwiIconNum As Long) As Long

Private Const IMAGE_FILTER = "Imagenes|*.bmp;*.ico;*.png;*.gif;*.jpg;*.jpeg"
Private Const EXPORT_FILTER = "Imagen PNG|*.png|Bitmap|*.bmp|Icono|*.ico|JPG|*.jpg|Todo los archivos|*.*"
Private Const ICO_FILTER = "Iconos de windows|*.ico|Todos los archivos|*.*"

Private Bitmaps     As Collection
Private FileName    As String
Private Changed     As Boolean
Private MnuIcons    As Collection


'/* Import From PE */
Private IconLib   As String
Private IconNum   As Long

Private Sub Form_Load()
Dim i As Long

    Call mvDrawBack

   
    Set MnuIcons = New Collection
    For i = 0 To imlMnu.ImageCount - 1
        MnuIcons.Add imlMnu.Picture(i, 16, 16)
    Next
    
    PutIconToVBMenu Me.hWnd, MnuIcons(1), 0, 0
    PutIconToVBMenu Me.hWnd, MnuIcons(2), 1, 0
    PutIconToVBMenu Me.hWnd, MnuIcons(3), 2, 0

    PutIconToVBMenu Me.hWnd, MnuIcons(5), 0, 1
    PutIconToVBMenu Me.hWnd, MnuIcons(6), 1, 1
    PutIconToVBMenu Me.hWnd, MnuIcons(7), 3, 1
    PutIconToVBMenu Me.hWnd, MnuIcons(8), 4, 1
    PutIconToVBMenu Me.hWnd, MnuIcons(9), 12, 1
    
    PutIconToVBMenu Me.hWnd, MnuIcons(10), 0, 2
    imlMnu.Clear
    
    If Command <> vbNullString Then
        FileName = Replace$(Command, Chr(34), vbNullString)
        Call ImportFromIconFile(FileName, Bitmaps)
        Call LoadBitmapList
    End If
    
End Sub

Private Sub mnuFile_Click(Index As Integer)
Dim Cmmdlg     As cDialog

    Select Case Index
        Case 0: '/* New */
            
            If Bitmaps Is Nothing Then Exit Sub
            
            If Changed Then
                Dim lRet As VbMsgBoxResult
                lRet = MsgBox("¡Desea guargar los cambios?", vbQuestion + vbYesNoCancel, "IconCreator")
                If lRet = vbCancel Then Exit Sub
                If lRet = vbYes Then Call mnuFile_Click(2)
            End If
            
            FileName = vbNullString
            Changed = False
            Set Bitmaps = Nothing
            lst.Clear
            
        Case 1 '/* Open */
        
            If Changed Then
                lRet = MsgBox("¡Desea guargar los cambios?", vbQuestion + vbYesNoCancel, "IconCreator")
                If lRet = vbCancel Then Exit Sub
                If lRet = vbYes Then Call mnuFile_Click(2)
            End If
            
            Set Cmmdlg = New cDialog
            Cmmdlg.Title = "Abrir icono"
            Cmmdlg.Filter = ICO_FILTER
            If Not Cmmdlg.ShowOpen(Me.hWnd) Then Exit Sub
            
            FileName = Cmmdlg.FileName
            Call ImportFromIconFile(Cmmdlg.FileName, Bitmaps)
            Call LoadBitmapList
            Changed = False
            
        Case 2, 3 '/* Save */
        
            If Bitmaps Is Nothing Then Exit Sub
            If FileName = vbNullString Or Index = 3 Then
                Set Cmmdlg = New cDialog
                Cmmdlg.Title = "Guardar icono"
                Cmmdlg.Filter = ICO_FILTER
                Cmmdlg.DefExtension = "ico"
                Cmmdlg.OverWritePrompt = True
                
                If Not Cmmdlg.ShowSave(Me.hWnd) Then Exit Sub
                FileName = Cmmdlg.FileName
            End If
            Call ExportToIconFile(FileName, Bitmaps)
            Changed = False
            
        Case 5 '/* Create From Image */
        
            If Changed Then
                lRet = MsgBox("¡Desea guargar los cambios?", vbQuestion + vbYesNoCancel, "IconCreator")
                If lRet = vbCancel Then Exit Sub
                If lRet = vbYes Then Call mnuFile_Click(2)
            End If
            
            Set Cmmdlg = New cDialog
            Cmmdlg.Title = "Crear paquete de iconos"
            Cmmdlg.Filter = IMAGE_FILTER
            If Not Cmmdlg.ShowOpen(Me.hWnd) Then Exit Sub
            
            If LCase(GetFileExt(Cmmdlg.FileName)) = "ico" Then
                FileName = Cmmdlg.FileName
                Call ImportFromIconFile(Cmmdlg.FileName, Bitmaps)
                Call LoadBitmapList
                Changed = False
                Exit Sub
            End If
            
            Dim CBmp    As cGDIPBitmap
            Dim cBmpItm As cGDIPBitmap
            Dim mFrm    As FrmCreate
            Dim cColl   As Collection
            Dim vElmnt  As Variant
        
            Set CBmp = New cGDIPBitmap
            
            If Not CBmp.LoadImage(Cmmdlg.FileName) Then
                MsgBox "¡No se ha podido cargar la imágen!", vbCritical
                Exit Sub
            End If
        
            Set mFrm = New FrmCreate
            mFrm.lblSize = CBmp.Width & "x" & CBmp.Height
            Set cColl = mFrm.GetSizes(Me)
            Unload mFrm
            
            If cColl Is Nothing Then Exit Sub
            
            lst.Clear
            PicPrev.Cls
            FileName = vbNullString
            Set Bitmaps = New Collection
            
            For Each vElmnt In cColl
                Set cBmpItm = New cGDIPBitmap
                If cBmpItm.LoadImage(CBmp.Stream(vElmnt(0), vElmnt(1), , False)) Then
                    Bitmaps.Add cBmpItm
                End If
            Next
            Call LoadBitmapList
            If lst.ListCount Then lst.ListIndex = 0
            Changed = True
            
    End Select
End Sub
Private Sub mnuBmp_Click(Index As Integer)


    Select Case Index
        Case 0: '/* Add */
        
            Dim Cmmdlg  As cDialog
            Dim CBmp    As cGDIPBitmap
            
            Set Cmmdlg = New cDialog
            Cmmdlg.Title = "Añadir imagen"
            Cmmdlg.Filter = IMAGE_FILTER
            If Not Cmmdlg.ShowOpen(Me.hWnd) Then Exit Sub
            If Bitmaps Is Nothing Then Set Bitmaps = New Collection
            
            If LCase(GetFileExt(Cmmdlg.FileName)) = "ico" Then
                Dim mFrm As FrmImport
                Set mFrm = New FrmImport
                Set CBmp = mFrm.ImportFrom(Cmmdlg.FileName, Me)
                Unload mFrm
                If CBmp Is Nothing Then Exit Sub
            Else
                Set CBmp = New cGDIPBitmap
                If Not CBmp.LoadImage(Cmmdlg.FileName) Then Exit Sub
            End If
        
            Bitmaps.Add CBmp
            
            lst.AddItem CBmp.Width & "x" & CBmp.Height
            lst.ListIndex = lst.ListCount - 1
            Changed = True
            
        Case 1 '/* Remove */
        
            If Bitmaps Is Nothing Then Exit Sub
            If lst.ListIndex = -1 Then Exit Sub
            If MsgBox("¿Desea eliminar la imagen?", vbQuestion + vbYesNo, "IconCreator") = vbNo Then Exit Sub
            Bitmaps.Remove lst.ListIndex + 1
            lst.RemoveItem lst.ListIndex
            PicPrev.Cls
            Changed = True
            
        Case 3 '/* Copy */
        
            If Bitmaps Is Nothing Then Exit Sub
            If lst.ListIndex = -1 Then Exit Sub
            
            Dim oPic As StdPicture
            Set oPic = Bitmaps(lst.ListIndex + 1).Picture
            Clipboard.Clear
            Clipboard.SetData oPic, vbCFDIB
        
        Case 4 '/* Paste */
        
            If Bitmaps Is Nothing Then Set Bitmaps = New Collection
            If Clipboard.GetFormat(vbCFBitmap) Then
                Set CBmp = New cGDIPBitmap
                If Not CBmp.LoadImage(Clipboard.GetData(vbCFBitmap)) Then Exit Sub
                Bitmaps.Add CBmp
                
                lst.AddItem CBmp.Width & "x" & CBmp.Height
                lst.ListIndex = lst.ListCount - 1
                Changed = True
                
            Else
                MsgBox "No Bitmap in Clipboard", vbCritical
            End If
        
        Case 5 '/* Duplicate */
        
            If Bitmaps Is Nothing Then Exit Sub
            If lst.ListIndex = -1 Then Exit Sub
            
            Dim ab()    As Byte
            
            ab = Bitmaps(lst.ListIndex + 1).Stream
            Set CBmp = New cGDIPBitmap
            If Not CBmp.LoadImage(ab) Then Exit Sub
            Bitmaps.Add CBmp
            
            lst.AddItem CBmp.Width & "x" & CBmp.Height
            lst.ListIndex = lst.ListCount - 1
            Changed = True
            
        Case 7 '/* Resize */
        
            If Bitmaps Is Nothing Then Exit Sub
            If lst.ListIndex = -1 Then Exit Sub
            
            Dim tmp     As String
            Dim lW      As Long
            Dim lH      As Long
            
            tmp = Bitmaps(lst.ListIndex + 1).Width & "x" & Bitmaps(lst.ListIndex + 1).Height
            tmp = InputBox("Ingrese el tamaño de la imagen " & vbNewLine & "Ancho x Alto [40x40] ", "Redimencionar", tmp)
            If tmp = vbNullString Then Exit Sub
            Call GetSize(tmp, lW, lH)
            
            If Bitmaps(lst.ListIndex + 1).Resize(lW, lH) Then
                lst.List(lst.ListIndex) = lW & "x" & lH
                lst_Click '/* Redraw */
                Changed = True
            End If

        Case 8 '/* Duplicate && Resize */
        
            If Bitmaps Is Nothing Then Exit Sub
            If lst.ListIndex = -1 Then Exit Sub
            
            tmp = Bitmaps(lst.ListIndex + 1).Width & "x" & Bitmaps(lst.ListIndex + 1).Height
            tmp = InputBox("Ingrese el tamaño de la imagen " & vbNewLine & "Ancho x Alto [40x40] ", "Redimencionar", tmp)
            
            If tmp = vbNullString Then Exit Sub
            Call GetSize(tmp, lW, lH)
            
            ab = Bitmaps(lst.ListIndex + 1).Stream(lW, lH, , False)
            Set CBmp = New cGDIPBitmap
            If Not CBmp.LoadImage(ab) Then Exit Sub
            Bitmaps.Add CBmp
            
            lst.AddItem CBmp.Width & "x" & CBmp.Height
            lst.ListIndex = lst.ListCount - 1
            Changed = True
        
        Case 10 '/* Add From Resource */
            
            If Not mvIconDialog Then Exit Sub
            If Bitmaps Is Nothing Then Set Bitmaps = New Collection
            
            Set mFrm = New FrmImport
            Set CBmp = mFrm.ImportFrom(IconLib, Me, True, IconNum)
            Unload mFrm
            If CBmp Is Nothing Then Exit Sub
            
            Bitmaps.Add CBmp
            
            lst.AddItem CBmp.Width & "x" & CBmp.Height
            lst.ListIndex = lst.ListCount - 1
            Changed = True
            
        Case 12 '/* Export */
        
            If Bitmaps Is Nothing Then Exit Sub
            If lst.ListIndex = -1 Then Exit Sub
    
            Set Cmmdlg = New cDialog
            Cmmdlg.Title = "Exportar"
            Cmmdlg.Filter = EXPORT_FILTER
            Cmmdlg.DefExtension = "png"
            Cmmdlg.OverWritePrompt = True
            If Not Cmmdlg.ShowSave(Me.hWnd) Then Exit Sub
            
            If LCase(GetFileExt(Cmmdlg.FileName) = "ico") Then
                If Not ExportBitmapToIcon(Bitmaps(lst.ListIndex + 1), Cmmdlg.FileName) Then MsgBox "¡Ocurrió un error al exportar!", vbCritical
            Else
                If Not Bitmaps(lst.ListIndex + 1).Save(Cmmdlg.FileName) Then MsgBox "¡Ocurrió un error al exportar!", vbCritical
            End If
            
    End Select
    
    
End Sub
Private Sub mnuHelp_Click(Index As Integer)
    FrmAbout.Show 1, Me
End Sub
Private Sub lst_Click()
On Error GoTo e
Dim ly      As Long
Dim lx      As Long
Dim lW      As Long
Dim lH      As Long
Dim Bmp     As cGDIPBitmap

    PicPrev.Cls
    Set Bmp = Bitmaps(lst.ListIndex + 1)
    If Bmp Is Nothing Then Debug.Print "Nothing Bmp for draw ": Exit Sub
    mvScalePicture Bmp.Width, Bmp.Height, PicPrev.ScaleWidth, PicPrev.ScaleHeight, lW, lH, lx, ly
    Bmp.Render PicPrev.Hdc, lx, ly, lW, lH
e:
End Sub
Private Sub lst_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If Button = 2 And lst.ListIndex <> -1 Then Me.PopupMenu mnu(1), , lst.Left + (x \ Screen.TwipsPerPixelX), lst.Top + (y \ Screen.TwipsPerPixelY), mnuBmp(0)
    If Button = 2 Then Me.PopupMenu mnu(1), , lst.Left + (x \ Screen.TwipsPerPixelX), lst.Top + (y \ Screen.TwipsPerPixelY), mnuBmp(0)
End Sub

Private Sub PicPrev_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If Button = 2 And lst.ListIndex <> -1 Then Me.PopupMenu mnu(1), , PicPrev.Left + x, PicPrev.Top + y, mnuBmp(0)
    If Button = 2 Then Me.PopupMenu mnu(1), , PicPrev.Left + x, PicPrev.Top + y, mnuBmp(0)
End Sub

Private Sub LoadBitmapList()
On Error GoTo e
Dim Bmp As Object

    lst.Clear
    For Each Bmp In Bitmaps
        lst.AddItem Bmp.Width & "x" & Bmp.Height
    Next
    lst.ListIndex = 0
    
    Exit Sub
e:
    Err.Clear
    Resume Next
End Sub
Private Function mvIconDialog() As Boolean
On Error Resume Next
    
    'If IconLib = vbNullString Then IconLib = "shell32.dll" & vbNullChar & Space(256)
    'IconLib = IconLib & vbNullChar & Space(256)
    
    IconLib = "shell32.dll" & vbNullChar & Space(256)
    IconLib = StrConv(IconLib, vbUnicode)
    
    If IconDialog(Me.hWnd, IconLib, Len(IconLib), IconNum) Then
        IconLib = StrConv(IconLib, vbFromUnicode)
        IconLib = Left(IconLib, InStr(IconLib, vbNullChar) - 1)
        mvIconDialog = True
    End If
    
End Function

Private Sub mvDrawBack()
Dim i As Long
Dim j As Long
Dim x As Long
    For j = -1 To PicPrev.ScaleHeight Step 5
        x = IIf(x = -1, 4, -1)
        For i = x To PicPrev.ScaleWidth Step 10
            PicPrev.Line (i, j)-(i + 4, j + 4), &HF2F2F2, BF
        Next
    Next
    PicPrev.Line (0, 0)-(PicPrev.ScaleWidth - 1, PicPrev.ScaleHeight - 1), vbButtonShadow, B
    PicPrev = PicPrev.Image
End Sub
Private Function GetSize(tmp As String, lW As Long, lH As Long) As Boolean
On Error Resume Next
Dim lSep As String
    
    If InStr(tmp, "*") Then lSep = "*"
    If InStr(LCase(tmp), "x") Then lSep = "x"
    lW = Val(Split(tmp, lSep)(0))
    lH = Val(Split(tmp, lSep)(1))
End Function
Private Function mvScalePicture( _
       ByVal lSrcWidth As Long, _
       ByVal lSrcHeight As Long, _
       ByVal lDstWidth As Long, _
       ByVal lDstHeight As Long, _
       ByRef lNewWidth As Long, _
       ByRef lNewHeight As Long, _
       ByRef lNewLeft As Long, _
       ByRef lNewTop As Long)

    Dim dHRatio As Double
    Dim dVRatio As Double
    Dim dRatio  As Double
    
    If lSrcWidth < lDstWidth And lSrcHeight < lDstHeight Then
        lNewWidth = lSrcWidth
        lNewHeight = lSrcHeight
        lNewLeft = (lDstWidth - lNewWidth) \ 2
        lNewTop = (lDstHeight - lNewHeight) \ 2
        Exit Function
    End If
    
    dHRatio = lSrcWidth / lDstWidth
    dVRatio = lSrcHeight / lDstHeight
     

    If dHRatio > dVRatio Then
        dRatio = dHRatio
    Else
        dRatio = dVRatio
    End If

    If Not dRatio = 0 Then
        lNewWidth = lSrcWidth / dRatio
        lNewHeight = lSrcHeight / dRatio
    End If
    
    lNewLeft = (lDstWidth - lNewWidth) / 2
    lNewTop = (lDstHeight - lNewHeight) / 2
End Function


