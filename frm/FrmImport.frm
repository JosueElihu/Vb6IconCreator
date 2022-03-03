VERSION 5.00
Begin VB.Form FrmImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar icono"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtnMain 
      Caption         =   "Importar"
      Height          =   495
      Index           =   0
      Left            =   4680
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
   End
   Begin VB.PictureBox PicPrev 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   2400
      ScaleHeight     =   260
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   3900
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
      Height          =   3840
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FrmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Import  As cIconEntry
Private Bmp     As cGDIPBitmap

Private Sub Form_Load()
    Call mvDrawBack
End Sub
Private Sub btnMain_Click(Index As Integer)
    Select Case Index
        Case 0
            If lst.ListCount = 0 Then Exit Sub
            If lst.ListIndex = -1 Then MsgBox "¡No ha seleccionado ningun item!", vbInformation: Exit Sub
            
            Dim hIcon As Long
            hIcon = Import.Icon(lst.ListIndex)
            If hIcon Then
                Set Bmp = New cGDIPBitmap
                If Not Bmp.LoadImage(hIcon) Then
                    MsgBox "¡No se ha podido importar el item!", vbCritical
                    Set Bmp = Nothing
                Else
                    Me.Visible = False
                End If
                DestroyIcon hIcon
            End If
End Select
End Sub
Private Sub lst_Click()
Dim lx  As Long
Dim ly  As Long
Dim lW  As Long
Dim lH  As Long
Dim i   As Long

    i = lst.ListIndex
    If i = -1 Then Exit Sub
    
    If Import.IconWidth(i) = 0 And Import.IconHeight(i) = 0 Then
        lW = 256
        lH = 256
    Else
        lW = Import.IconWidth(i)
        lH = Import.IconHeight(i)
    End If
    
    With PicPrev
        .Cls
        lx = (.ScaleWidth / 2) - (lW / 2)
        ly = (.ScaleHeight / 2) - (lH / 2)
        Import.DrawIcon .Hdc, i, lx, ly, lW, lH
        .Refresh
    End With
        
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '
End Sub


Public Function ImportFrom(ByVal sFileName As String, Frm As Form, Optional ByVal IsPE As Boolean, Optional ByVal lPEID As Long) As cGDIPBitmap

    Set Import = New cIconEntry
    If IsPE And LCase$(GetFileExt(sFileName)) = "ico" Then IsPE = False
    If Not IsPE Then Import.OpenIconFile sFileName Else Import.OpenIconFromPE sFileName, lPEID
    If Import.IconCount = 0 Then MsgBox "¡No se ha sido posible importar los iconos!", vbCritical: Exit Function
    LoadIconEntries
    If lst.ListCount Then lst.ListIndex = 0
    
    Me.Show 1, Frm
    Set ImportFrom = Bmp
    
End Function

Private Sub LoadIconEntries()
Dim i As Long
    For i = 0 To Import.IconCount - 1
        lst.AddItem Import.IconWidth(i) & "x" & Import.IconHeight(i) & " " & Import.ColorDepth(i) & "Bit"
    Next
End Sub


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

