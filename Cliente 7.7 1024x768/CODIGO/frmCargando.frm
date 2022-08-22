VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7500
   ClientLeft      =   1005
   ClientTop       =   810
   ClientWidth     =   9975
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmCargando.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Status 
      Height          =   2340
      Left            =   3240
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   2640
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   4128
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmCargando.frx":0CCA
      MouseIcon       =   "frmCargando.frx":0D4E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image LOGO 
      Height          =   7500
      Left            =   0
      Picture         =   "frmCargando.frx":0D6A
      Top             =   0
      Width           =   9990
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim result As Long
    result = SetWindowLong(Status.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    LOGO.Picture = LoadPicture(DirGraficos & "mar.jpg")

End Sub

'IRON AO: Auto Update
Private Sub Inet1_StateChanged(ByVal State As Integer)

End Sub

