VERSION 5.00
Begin VB.Form FrmCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crear paquete de iconos"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCreate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   297
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtnMain 
      Caption         =   "Crear"
      Height          =   495
      Index           =   2
      Left            =   2280
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton BtnMain 
      Caption         =   "Quitar tamaño"
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton BtnMain 
      Caption         =   "Añadir tamaño"
      Height          =   495
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
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
      Height          =   2790
      ItemData        =   "FrmCreate.frx":6012
      Left            =   120
      List            =   "FrmCreate.frx":6031
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0x0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2280
      TabIndex        =   6
      Top             =   600
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño de Imagen"
      Height          =   195
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   1380
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   152
      X2              =   280
      Y1              =   152
      Y2              =   152
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaños"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "FrmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Coll  As Collection

Private Sub BtnMain_Click(Index As Integer)

    Select Case Index
        Case 0:
                Dim tmp As String
                tmp = InputBox("Ingrese el tamaño del icono " & vbNewLine & "Ancho x Alto [40x40] ", "Redimencionar", vbNullString)
                If tmp = vbNullString Then Exit Sub
                If InStr(tmp, "*") = 0 And InStr(tmp, "x") = 0 And InStr(tmp, "X") = 0 Then MsgBox "Tamaño inválido", vbCritical
                tmp = Replace$(tmp, " ", vbNullString)
                tmp = Replace$(tmp, "*", "x")
                tmp = Replace$(tmp, "X", "x")
                
                If ExistSize(tmp) Then
                    If MsgBox("El tamaño " & tmp & " ya existe" & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
                End If
                lst.AddItem tmp
                
        Case 1:
            
            If lst.ListIndex = -1 Then Exit Sub
            lst.RemoveItem lst.ListIndex
        
        Case 2
            If lst.ListCount = 0 Then MsgBox "¡No ha ingresado ningun tamaño!", vbInformation: Exit Sub
            Me.Visible = False
            
            Dim sElmnt() As String
            Dim i As Long
            
            Set m_Coll = New Collection
            For i = 0 To lst.ListCount - 1
                sElmnt = Split(lst.List(i), "x")
                m_Coll.Add sElmnt
            Next
    End Select
End Sub

Private Function ExistSize(StrSize As String) As Boolean
Dim i As Long

    For i = 0 To lst.ListCount - 1
        If lst.List(i) = StrSize Then ExistSize = True: Exit Function
    Next
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
End Sub

Public Function GetSizes(Frm As Form) As Collection
    
    Me.Show 1, Frm
    Set GetSizes = m_Coll
    'If GetSizes Is Nothing Then Set GetSizes = New Collection
    
End Function
