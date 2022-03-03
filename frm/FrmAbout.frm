VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de..."
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4395
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   293
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton BtnMain 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   240
      MousePointer    =   99  'Custom
      Picture         =   "FrmAbout.frx":11D8A
      ScaleHeight     =   36
      ScaleMode       =   0  'User
      ScaleWidth      =   36
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   540
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   16
      X2              =   272
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Simple Icon Creator "
      Height          =   795
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   3465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   16
      X2              =   272
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IconCreator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   1170
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Form_Load()
    With Picture1
    .Picture = GdipBitmap_(.Picture).Picture(.ScaleWidth, .ScaleHeight, .BackColor, dpiAware:=False)
    End With

    lblTitle = "IconCreator " & App.Major & "." & App.Minor & "." & App.Revision
    Label1 = "Simple icon maker built in VB6, to create windows icon pack easily" & vbNewLine & vbNewLine & "By: J. Elihu - © 2022"
    
End Sub

Private Sub BtnMain_Click()
    Unload Me
End Sub

