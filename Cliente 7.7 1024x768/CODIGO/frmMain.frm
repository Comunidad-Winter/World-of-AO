VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Server Aodrag"
   ClientHeight    =   10815
   ClientLeft      =   1725
   ClientTop       =   1200
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":0CCA
   MousePointer    =   99  'Custom
   Picture         =   "frmMain.frx":1994
   ScaleHeight     =   721
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6360
      Top             =   2640
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.CommandButton PGM 
      Caption         =   "Panel GM"
      Height          =   375
      Left            =   10080
      TabIndex        =   57
      Top             =   9600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Torneo 
      Caption         =   "Torneo"
      Height          =   375
      Left            =   10200
      MaskColor       =   &H000000FF&
      TabIndex        =   56
      Top             =   9960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Contador 
      Interval        =   1000
      Left            =   4560
      Top             =   4080
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   1680
      Top             =   960
   End
   Begin RichTextLib.RichTextBox RecTxt2 
      Height          =   1185
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   240
      Visible         =   0   'False
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2090
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":731FA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox PicMontura 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   9465
      MouseIcon       =   "frmMain.frx":73277
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   49
      Top             =   9930
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.ListBox Chats 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      ItemData        =   "frmMain.frx":73F41
      Left            =   8280
      List            =   "frmMain.frx":73F43
      MouseIcon       =   "frmMain.frx":73F45
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   10440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer CuentaSeg 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6480
      Top             =   6240
   End
   Begin VB.Timer mensajes 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7440
      Top             =   6720
   End
   Begin VB.Timer mensajes1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7440
      Top             =   5880
   End
   Begin VB.Timer Wpetimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1680
      Top             =   1440
   End
   Begin VB.Timer CheatTimer2 
      Interval        =   60000
      Left            =   2160
      Top             =   1440
   End
   Begin VB.Timer Cheat 
      Interval        =   2000
      Left            =   2640
      Top             =   960
   End
   Begin MSWinsockLib.Winsock ws_cliente 
      Left            =   7560
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7200
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer TimerLabel 
      Interval        =   8000
      Left            =   2640
      Top             =   1440
   End
   Begin VB.Timer smstimer 
      Interval        =   1500
      Left            =   1200
      Top             =   960
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1500
      Visible         =   0   'False
      Width           =   10950
   End
   Begin VB.Timer Sh 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   720
      Top             =   960
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2160
      Top             =   960
   End
   Begin RichTextLib.RichTextBox RecTxt3 
      Height          =   1185
      Left            =   90
      TabIndex        =   44
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   240
      Visible         =   0   'False
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2090
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":74C0F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1185
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   240
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2090
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":74C8C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   12240
      MouseIcon       =   "frmMain.frx":74D09
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":759D3
      ScaleHeight     =   192
      ScaleMode       =   0  'User
      ScaleWidth      =   192
      TabIndex        =   8
      Top             =   3120
      Width           =   2880
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   12120
      MouseIcon       =   "frmMain.frx":99A17
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   3135
   End
   Begin RichTextLib.RichTextBox RecTxt4 
      Height          =   1185
      Left            =   90
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   240
      Visible         =   0   'False
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   2090
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":9A6E1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image8 
      Height          =   495
      Left            =   1920
      Top             =   9720
      Width           =   1575
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   3840
      Top             =   10320
      Width           =   375
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   720
      Top             =   9840
      Width           =   405
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "99999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   14220
      TabIndex        =   55
      Top             =   2085
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "99999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   14220
      TabIndex        =   54
      Top             =   1755
      Width           =   975
   End
   Begin VB.Image InfoX 
      Height          =   540
      Left            =   14235
      Picture         =   "frmMain.frx":9A75E
      Top             =   5580
      Width           =   945
   End
   Begin VB.Image LanzarX 
      Height          =   555
      Left            =   12135
      Picture         =   "frmMain.frx":9FAEF
      Top             =   5580
      Width           =   1995
   End
   Begin VB.Image Mascotax 
      Height          =   375
      Left            =   1440
      Top             =   9840
      Width           =   375
   End
   Begin VB.Label FPSVIEW 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   5760
      TabIndex        =   53
      Top             =   10200
      Width           =   375
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3840
      TabIndex        =   52
      Top             =   10440
      Width           =   1215
   End
   Begin VB.Label Label55 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   51
      Top             =   10440
      Width           =   1215
   End
   Begin VB.Image Donaciones 
      Height          =   255
      Left            =   13800
      Top             =   2040
      Width           =   375
   End
   Begin VB.Image Drops 
      Height          =   495
      Left            =   8520
      Top             =   10320
      Width           =   375
   End
   Begin VB.Image InfoQuest 
      Height          =   375
      Left            =   1140
      Top             =   9840
      Width           =   255
   End
   Begin VB.Image Revivir 
      Height          =   345
      Left            =   12840
      Top             =   9720
      Width           =   375
   End
   Begin VB.Image Canjes 
      Height          =   255
      Left            =   13860
      Top             =   1680
      Width           =   315
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00000000&
      Height          =   7725
      Left            =   0
      Top             =   1845
      Width           =   12015
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   14880
      Top             =   120
      Width           =   375
   End
   Begin VB.Label cmdMascotas 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   8880
      MouseIcon       =   "frmMain.frx":A4C42
      MousePointer    =   99  'Custom
      TabIndex        =   39
      ToolTipText     =   "Mascotas"
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1. General"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10725
      TabIndex        =   50
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Image IrCastillo 
      Height          =   255
      Index           =   4
      Left            =   11550
      MousePointer    =   99  'Custom
      Top             =   9960
      Width           =   255
   End
   Begin VB.Image IrCastillo 
      Height          =   255
      Index           =   3
      Left            =   11640
      MousePointer    =   99  'Custom
      Top             =   10440
      Width           =   255
   End
   Begin VB.Image IrCastillo 
      Height          =   255
      Index           =   2
      Left            =   11040
      MousePointer    =   99  'Custom
      Top             =   9960
      Width           =   375
   End
   Begin VB.Image IrCastillo 
      Height          =   255
      Index           =   1
      Left            =   11640
      MousePointer    =   99  'Custom
      Top             =   9600
      Width           =   255
   End
   Begin VB.Image IrCastillo 
      Height          =   255
      Index           =   0
      Left            =   12000
      MousePointer    =   99  'Custom
      Top             =   9960
      Width           =   255
   End
   Begin VB.Image Map 
      Height          =   375
      Left            =   240
      MousePointer    =   99  'Custom
      ToolTipText     =   "Mapa del mundo."
      Top             =   9780
      Width           =   495
   End
   Begin VB.Image ImageMensaje 
      Height          =   375
      Left            =   1440
      MouseIcon       =   "frmMain.frx":A590C
      MousePointer    =   99  'Custom
      Top             =   10320
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   300
      Left            =   -120
      MouseIcon       =   "frmMain.frx":A65D6
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":A72A0
      ToolTipText     =   "Borrar todos los mensajes almacenados."
      Top             =   600
      Width           =   300
   End
   Begin VB.Image Fortaleza 
      Height          =   225
      Left            =   11610
      MouseIcon       =   "frmMain.frx":A7792
      Picture         =   "frmMain.frx":A845C
      ToolTipText     =   "Fortaleza Atacada"
      Top             =   9960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Norte 
      Height          =   225
      Left            =   11640
      MouseIcon       =   "frmMain.frx":A86F6
      Picture         =   "frmMain.frx":A93C0
      ToolTipText     =   "Castillo Norte atacado."
      Top             =   9600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image emoticono 
      Height          =   375
      Left            =   14280
      MouseIcon       =   "frmMain.frx":A965A
      MousePointer    =   99  'Custom
      ToolTipText     =   "Emoticonos"
      Top             =   9720
      Width           =   375
   End
   Begin VB.Image foto 
      Height          =   300
      Left            =   14760
      MouseIcon       =   "frmMain.frx":AA324
      MousePointer    =   99  'Custom
      ToolTipText     =   "Hacer Foto"
      Top             =   9720
      Width           =   285
   End
   Begin VB.Image sonido 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   13920
      MouseIcon       =   "frmMain.frx":AAFEE
      MousePointer    =   99  'Custom
      ToolTipText     =   "Ajustes de Sonido."
      Top             =   9720
      Width           =   330
   End
   Begin VB.Image DesInv 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   15000
      MouseIcon       =   "frmMain.frx":ABCB8
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":AC982
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image DesInv 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   15000
      MouseIcon       =   "frmMain.frx":ACD2A
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":AD9F4
      Stretch         =   -1  'True
      Top             =   5280
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Mochila 
      Height          =   615
      Left            =   12120
      MouseIcon       =   "frmMain.frx":ADD90
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      ToolTipText     =   "Inventario"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Image HechizosImg 
      Height          =   615
      Left            =   13680
      MouseIcon       =   "frmMain.frx":AEA5A
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      ToolTipText     =   "Hechizos"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1080
      MouseIcon       =   "frmMain.frx":AF724
      MousePointer    =   99  'Custom
      TabIndex        =   45
      ToolTipText     =   "Logros"
      Top             =   10320
      Width           =   270
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Misiones"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   8400
      MouseIcon       =   "frmMain.frx":B03EE
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Top             =   10440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Minimapa 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   17400
      MouseIcon       =   "frmMain.frx":B10B8
      MousePointer    =   99  'Custom
      TabIndex        =   42
      ToolTipText     =   "Activar el MiniMapa."
      Top             =   10440
      Width           =   1455
   End
   Begin VB.Label cmdFoto 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   14760
      MouseIcon       =   "frmMain.frx":B1D82
      MousePointer    =   99  'Custom
      TabIndex        =   41
      ToolTipText     =   "Pulsa aqu� para hacer una foto."
      Top             =   9720
      Width           =   255
   End
   Begin VB.Label cmdEmoticon 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   14280
      MouseIcon       =   "frmMain.frx":B2A4C
      MousePointer    =   99  'Custom
      TabIndex        =   40
      ToolTipText     =   "Pulsa aqu� para colocar Emoticonos sobre tu personaje."
      Top             =   9720
      Width           =   375
   End
   Begin VB.Label cmdEstadisticas 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   720
      MouseIcon       =   "frmMain.frx":B3716
      MousePointer    =   99  'Custom
      TabIndex        =   38
      ToolTipText     =   "Estadisticas"
      Top             =   10320
      Width           =   330
   End
   Begin VB.Label cmdHabilidades 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   12720
      MouseIcon       =   "frmMain.frx":B43E0
      MousePointer    =   99  'Custom
      TabIndex        =   37
      ToolTipText     =   "Habilidades"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label cmdMejores 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmMain.frx":B50AA
      MousePointer    =   99  'Custom
      TabIndex        =   36
      ToolTipText     =   "Mejores"
      Top             =   10200
      Width           =   375
   End
   Begin VB.Label cmdClanes 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3360
      MouseIcon       =   "frmMain.frx":B5D74
      MousePointer    =   99  'Custom
      TabIndex        =   35
      ToolTipText     =   "Clanes"
      Top             =   10320
      Width           =   255
   End
   Begin VB.Label cmdCastillos 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1920
      MouseIcon       =   "frmMain.frx":B6A3E
      MousePointer    =   99  'Custom
      TabIndex        =   34
      ToolTipText     =   "Castillos"
      Top             =   10320
      Width           =   375
   End
   Begin VB.Label cmdAyuda 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3000
      MouseIcon       =   "frmMain.frx":B7708
      MousePointer    =   99  'Custom
      TabIndex        =   33
      ToolTipText     =   "Ayuda"
      Top             =   10320
      Width           =   255
   End
   Begin VB.Label cmdLanzar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   12120
      MouseIcon       =   "frmMain.frx":B83D2
      MousePointer    =   99  'Custom
      TabIndex        =   32
      ToolTipText     =   "Lanzar un Hechizo"
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label cmdInfo 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   14280
      MouseIcon       =   "frmMain.frx":B909C
      MousePointer    =   99  'Custom
      TabIndex        =   31
      ToolTipText     =   "Informaci�n sobre los hechizos"
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Mexp 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   10200
      LinkTimeout     =   0
      TabIndex        =   30
      Top             =   10080
      Width           =   615
   End
   Begin VB.Label Mnivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   10080
      TabIndex        =   29
      Top             =   9840
      Width           =   615
   End
   Begin VB.Label Mnombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   9240
      TabIndex        =   28
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Label Arma 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8640
      TabIndex        =   27
      Top             =   10440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Escudo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   8640
      TabIndex        =   26
      Top             =   10440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Botas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   8640
      TabIndex        =   25
      Top             =   10560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Tronco 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   12480
      TabIndex        =   24
      Top             =   10560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Cabeza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   12480
      TabIndex        =   23
      Top             =   10440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label FamaLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   12840
      TabIndex        =   22
      ToolTipText     =   "Popularidad"
      Top             =   1755
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   12360
      MouseIcon       =   "frmMain.frx":B9D66
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   1680
      Width           =   495
   End
   Begin VB.Image Sur 
      Height          =   225
      Left            =   11640
      MouseIcon       =   "frmMain.frx":BAA30
      Picture         =   "frmMain.frx":BB6FA
      ToolTipText     =   "Castillo Sur atacado."
      Top             =   10440
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Este 
      Height          =   225
      Left            =   12000
      MouseIcon       =   "frmMain.frx":BB994
      Picture         =   "frmMain.frx":BC65E
      ToolTipText     =   "Castillo Este atacado."
      Top             =   9960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image CandadoO 
      Height          =   375
      Left            =   13560
      MouseIcon       =   "frmMain.frx":BC8F8
      MousePointer    =   99  'Custom
      ToolTipText     =   "Seguro de Objetos"
      Top             =   9720
      Width           =   375
   End
   Begin VB.Image CandadoA 
      Height          =   375
      Left            =   13320
      MouseIcon       =   "frmMain.frx":BD5C2
      MousePointer    =   99  'Custom
      ToolTipText     =   "Seguro de armas."
      Top             =   9720
      Width           =   255
   End
   Begin VB.Label lblparty 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2400
      MouseIcon       =   "frmMain.frx":BE28C
      MousePointer    =   99  'Custom
      TabIndex        =   20
      ToolTipText     =   "Partys"
      Top             =   10320
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "67"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   14520
      TabIndex        =   19
      Top             =   10200
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "34"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   14040
      TabIndex        =   18
      Top             =   10200
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Jugadores: 110"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10200
      MouseIcon       =   "frmMain.frx":BEF56
      TabIndex        =   17
      Top             =   10320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   11160
      MouseIcon       =   "frmMain.frx":BFC20
      MousePointer    =   99  'Custom
      ToolTipText     =   "Pulsa para ver el n�mero de Jugadores Online."
      Top             =   10320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "32000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   12840
      TabIndex        =   16
      Top             =   2010
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   12000
      MouseIcon       =   "frmMain.frx":C08EA
      MousePointer    =   99  'Custom
      ToolTipText     =   "Muertes Conseguidas."
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "110"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   13680
      MouseIcon       =   "frmMain.frx":C15B4
      MousePointer    =   99  'Custom
      TabIndex        =   15
      ToolTipText     =   "Pulsa para ver el Mapa del Mundo AodraG"
      Top             =   10200
      Width           =   375
   End
   Begin VB.Label Porexp 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6780
      TabIndex        =   14
      Top             =   10200
      Width           =   555
   End
   Begin VB.Label Labelagua 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13080
      TabIndex        =   13
      Top             =   9375
      Width           =   1095
   End
   Begin VB.Image AGUAsp 
      Height          =   390
      Left            =   12195
      Picture         =   "frmMain.frx":C227E
      Top             =   9300
      Width           =   2940
   End
   Begin VB.Label Labelenergia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13080
      TabIndex        =   11
      Top             =   6660
      Width           =   1095
   End
   Begin VB.Image STAShp 
      Height          =   390
      Left            =   12195
      Picture         =   "frmMain.frx":C74AE
      Top             =   6570
      Width           =   2940
   End
   Begin VB.Label Labelmana 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9999/9999"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13080
      TabIndex        =   10
      Top             =   7995
      Width           =   1095
   End
   Begin VB.Image MANShp 
      Height          =   390
      Left            =   12195
      Picture         =   "frmMain.frx":CC4F1
      Top             =   7920
      Width           =   2940
   End
   Begin VB.Label Labelvida 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   13080
      TabIndex        =   9
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Image Hpshp 
      Height          =   390
      Left            =   12195
      Picture         =   "frmMain.frx":D31B7
      Top             =   7245
      Width           =   2940
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   13080
      TabIndex        =   7
      ToolTipText     =   "Nivel del Personaje."
      Top             =   840
      Width           =   165
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "aodragbot"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   13800
      TabIndex        =   6
      ToolTipText     =   "Nombre del Personaje."
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "999999999999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   165
      Left            =   14220
      TabIndex        =   5
      Top             =   1425
      Width           =   900
   End
   Begin VB.Image Image3 
      Height          =   315
      Index           =   0
      Left            =   13800
      MouseIcon       =   "frmMain.frx":D84F8
      MousePointer    =   99  'Custom
      ToolTipText     =   "Soltar Oro."
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label peso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100/100 KG"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   165
      Left            =   8280
      TabIndex        =   4
      ToolTipText     =   "Peso"
      Top             =   10560
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Labelcomida 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13080
      TabIndex        =   12
      Top             =   8685
      Width           =   1095
   End
   Begin VB.Image COMIDAsp 
      Height          =   390
      Left            =   12195
      Picture         =   "frmMain.frx":D91C2
      Top             =   8610
      Width           =   2940
   End
   Begin VB.Image Oeste 
      Height          =   225
      Left            =   11160
      MouseIcon       =   "frmMain.frx":DE341
      Picture         =   "frmMain.frx":DF00B
      ToolTipText     =   "Castillo Oeste atacado."
      Top             =   9960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image ImageMensaje1 
      Height          =   480
      Left            =   8520
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":DF2A5
      Top             =   10320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999999/999999"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   46
      Top             =   9855
      Width           =   2625
   End
   Begin VB.Image Estimulo 
      Height          =   435
      Left            =   12300
      Picture         =   "frmMain.frx":DFBE7
      ToolTipText     =   "Indica si el personaje tiene bonus en sus atributos."
      Top             =   10095
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image luzaviso 
      Height          =   750
      Left            =   8400
      MouseIcon       =   "frmMain.frx":E36DB
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":E43A5
      ToolTipText     =   "Indicador de la dificultad del Mapa en el que te encuentras."
      Top             =   10200
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image ExpSP 
      Height          =   255
      Left            =   4230
      Picture         =   "frmMain.frx":E493E
      Top             =   9840
      Width           =   4410
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public str_contenido_archivo As String, str_nombre_archivo As String, str_ruta_remota As String
Attribute str_nombre_archivo.VB_VarUserMemId = 1073938432
Attribute str_ruta_remota.VB_VarUserMemId = 1073938432

Dim lng_tama�o_archivo As Long
Attribute lng_tama�o_archivo.VB_VarUserMemId = 1073938435
'-----------------

Dim Text1 As String
Attribute Text1.VB_VarUserMemId = 1073938436
Dim Text2 As String
Attribute Text2.VB_VarUserMemId = 1073938437
Dim Text3 As String
Attribute Text3.VB_VarUserMemId = 1073938438

'----------------------------
'Public ActualSecond As Long
'Public LastSecond As Long
Public tX As Integer
Attribute tX.VB_VarUserMemId = 1073938439
Public tY As Integer
Attribute tY.VB_VarUserMemId = 1073938440
Public MouseX As Long
Attribute MouseX.VB_VarUserMemId = 1073938441
Public MouseY As Long
Attribute MouseY.VB_VarUserMemId = 1073938442
Dim HechiClk As Integer
Attribute HechiClk.VB_VarUserMemId = 1073938443
Dim gDSB As DirectSoundBuffer
Attribute gDSB.VB_VarUserMemId = 1073938444
Dim gD As DSBUFFERDESC
Attribute gD.VB_VarUserMemId = 1073938445
Dim gW As WAVEFORMATEX
Attribute gW.VB_VarUserMemId = 1073938446
Dim gFileName As String
Attribute gFileName.VB_VarUserMemId = 1073938447
Dim dsE As DirectSoundEnum
Attribute dsE.VB_VarUserMemId = 1073938448
Dim pos(0) As DSBPOSITIONNOTIFY
Attribute pos.VB_VarUserMemId = 1073938449
Public IsPlaying As Byte
Attribute IsPlaying.VB_VarUserMemId = 1073938450
Dim variable As String
Attribute variable.VB_VarUserMemId = 1073938451
Dim ie As Object
Attribute ie.VB_VarUserMemId = 1073938452
Dim endEvent As Long
Attribute endEvent.VB_VarUserMemId = 1073938453
Implements DirectXEvent
Dim ShiftDown As Boolean
Attribute ShiftDown.VB_VarUserMemId = 1073938454

Private Declare Sub ZeroMemory _
                     Lib "kernel32.dll" _
                         Alias "RtlZeroMemory" (destination As Any, _
                                                ByVal length As Long)

Private Type RECT

    Left As Long
    Top As Long
    Right As Long
    Bottom As Long

End Type

Private Declare Function GetWindowRect _
                          Lib "user32" (ByVal hwnd As Long, _
                                        lpRect As RECT) As Long
Private Declare Function SetWindowPos _
                          Lib "user32" (ByVal hwnd As Long, _
                                        ByVal hWndInsertAfter As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal cx As Long, _
                                        ByVal cy As Long, _
                                        ByVal wFlags As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Sub CandadoA_Click()

'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus
'pluto:6.0A
    If SeguroCrimi = True Then
        'frmMain.CandadoA.Picture = LoadPicture(App.Path & "\graficos\c1a.jpg")
        SeguroCrimi = False
    Else
        'frmMain.CandadoA.Picture = LoadPicture(App.Path & "\graficos\c1c.jpg")
        SeguroCrimi = True

    End If

    Call SendData("SEG")

End Sub

Private Sub CandadoO_Click()
'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus

    If seguroobjetos = True Then
        'frmMain.CandadoO.Picture = LoadPicture(App.Path & "\graficos\c2a.jpg")
        Call AddtoRichTextBox(frmMain.RecTxt, "Seguro De Objetos Desactivado", 0, 191, 128, True, False, False)

        seguroobjetos = False
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, "Seguro De Objetos Activado", 0, 191, 128, True, False, False)
        seguroobjetos = True
        'frmMain.CandadoO.Picture = LoadPicture(App.Path & "\graficos\c2c.jpg")

    End If

End Sub


Private Sub Canjes_Click()
    SendData ("CCANJE")
End Sub

Private Sub Chats_Click()

    If Chats.ListIndex = 0 Then
        frmMain.RecTxt2.Visible = True
        frmMain.RecTxt.Visible = False
        frmMain.SendTxt.Visible = True
        frmMain.RecTxt3.Visible = False
        frmMain.RecTxt4.Visible = False
        frmMain.SendTxt.SetFocus
        ChatElegido = 1

    End If

    If Chats.ListIndex = 1 Then
        ChatElegido = 2
        frmMain.RecTxt.Visible = True
        frmMain.RecTxt2.Visible = False
        frmMain.SendTxt.Visible = True
        frmMain.RecTxt4.Visible = False
        frmMain.SendTxt.SetFocus
        frmMain.RecTxt3.Visible = False

    End If

    If Chats.ListIndex = 2 Then
        frmMain.RecTxt.Visible = False
        frmMain.RecTxt2.Visible = False
        frmMain.RecTxt3.Visible = True
        frmMain.RecTxt4.Visible = False
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
        ChatElegido = 0

    End If

    If Chats.ListIndex = 3 Then
        frmMain.RecTxt.Visible = False
        frmMain.RecTxt2.Visible = False
        frmMain.RecTxt3.Visible = False
        frmMain.RecTxt4.Visible = True
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
        ChatElegido = 0

    End If

    If Chats.ListIndex = 4 Then
        frmMain.RecTxt.Visible = True
        frmMain.RecTxt2.Visible = False
        frmMain.RecTxt3.Visible = False
        frmMain.RecTxt4.Visible = False
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
        ChatElegido = 4

    End If

End Sub

Private Sub cmdAyuda_Click()
    Frmayuda.Show , frmMain
    'If cmdManual.Visible = False Then
    'cmdManual.Visible = True
    'cmdComanditos.Visible = True
    'cmdTeclas.Visible = True
    'cmdChat.Visible = True
    'cmdForo.Visible = True
    'cmdWeb.Visible = True
    'decorado.Visible = False
    '-------------------------
    'cmdBoveda.Visible = False
    'cmdIngresar.Visible = False
    'cmdRetirar.Visible = False
    'cmdDescansar.Visible = False
    'cmdMeditar.Visible = False
    'cmdComerciar.Visible = False
    'cmdCambiarClave.Visible = False
    'cmdComanditos.Visible = False
    '---------------------------
    'cmdObjetos.Visible = False

    '----------------------------
    'cmdEstadisticas.Visible = False
    'cmdHabilidades.Visible = False
    'cmdMejores.Visible = False
    'cmdClanes.Visible = False
    'cmdCastillos.Visible = False
    'cmdMascotas.Visible = False
    '----------------------------
    'Else
    'cmdManual.Visible = False

    'cmdTeclas.Visible = False
    'cmdChat.Visible = False
    'cmdForo.Visible = False
    'cmdWeb.Visible = False
    'decorado.Visible = True
    'End If
End Sub

'Private Sub cmdBoveda_Click()
'Call SendData("/BOVEDA")
'End Sub

Private Sub cmdCambiarClave_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, _
                          "Para cambiar tu clave escribe el comando /PASSWD (deja un espacio) a continuaci�n escribe tu nueva clave", _
                          187, 87, 87, 0, 0)

End Sub

Private Sub cmdAyuda_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    frmMain.cmdEstadisticas.ForeColor = &H808080
    frmMain.cmdAyuda.ForeColor = &HC0C0C0
    frmMain.cmdCastillos.ForeColor = &H808080
    frmMain.cmdClanes.ForeColor = &H808080
    frmMain.Minimapa.ForeColor = &H808080
    frmMain.cmdHabilidades.ForeColor = &H808080
    frmMain.cmdMascotas.ForeColor = &H808080
    frmMain.cmdMejores.ForeColor = &H808080
    frmMain.Label9.ForeColor = &H808080
    frmMain.Label7.ForeColor = &H808080

End Sub

Private Sub cmdCastillos_Click()
    Call SendData("CT")

End Sub

Private Sub cmdChat_Click()
    variable = "http://www.irc-hispano.org/index.php?seccion=canal&sec=&can=aodrag"

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate variable
    Call AddtoRichTextBox(frmMain.RecTxt, _
                          "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 0, 0, 0, _
                          True, False, False)

End Sub

Private Sub cmdCastillos_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    frmMain.cmdEstadisticas.ForeColor = &H808080
    frmMain.cmdAyuda.ForeColor = &H808080
    frmMain.cmdCastillos.ForeColor = &HC0C0C0
    frmMain.cmdClanes.ForeColor = &H808080
    frmMain.Minimapa.ForeColor = &H808080
    frmMain.cmdHabilidades.ForeColor = &H808080
    frmMain.cmdMascotas.ForeColor = &H808080
    frmMain.cmdMejores.ForeColor = &H808080
    frmMain.Label9.ForeColor = &H808080
    frmMain.Label7.ForeColor = &H808080

End Sub

Private Sub cmdClanes_Click()

    If Not frmGuildLeader.Visible Then
        Call SendData("GLINFO")

    End If

End Sub

Private Sub cmdComanditos_Click()
    frmComandos.Show vbModal

End Sub

Private Sub cmdComerciar_Click()
    Call SendData("/COMERCIAR")

End Sub

Private Sub cmdDescansar_Click()
    Call SendData("/DESCANSAR")

End Sub

Private Sub cmdClanes_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    frmMain.cmdEstadisticas.ForeColor = &H808080
    frmMain.cmdAyuda.ForeColor = &H808080
    frmMain.cmdCastillos.ForeColor = &H808080
    frmMain.cmdClanes.ForeColor = &HC0C0C0
    frmMain.Minimapa.ForeColor = &H808080
    frmMain.cmdHabilidades.ForeColor = &H808080
    frmMain.cmdMascotas.ForeColor = &H808080
    frmMain.cmdMejores.ForeColor = &H808080
    frmMain.Label9.ForeColor = &H808080
    frmMain.Label7.ForeColor = &H808080

End Sub

Private Sub cmdEstadisticas_Click()
    LlegaronAtrib = False
    'LlegaronSkills = False
    LlegoFama = False
    LLegoEsta = False
    SendData "ATRI"
    'SendData "ESKI"
    SendData "FAMA"
    SendData "ESTA"

    Do While Not LLegoEsta Or Not LlegaronAtrib Or Not LlegoFama
        DoEvents    'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    'If FrmHechizos.Visible = True Then FrmHechizos.SetFocus
    'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus
    ' frmEstadisticas.Iniciar_Labels
    'frmEstadisticas.Show
    'frmEstadisticas.Visible = True
    LlegaronAtrib = False
    'LlegaronSkills = False
    LlegoFama = False
    LLegoEsta = False

End Sub

Private Sub cmdForo_Click()
    variable = "http://juegosdrag.es/foros/"

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate variable
    Call AddtoRichTextBox(frmMain.RecTxt, _
                          "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 0, 0, 0, _
                          True, False, False)

End Sub

Private Sub cmdEstadisticas_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    frmMain.cmdEstadisticas.ForeColor = &HC0C0C0
    frmMain.cmdAyuda.ForeColor = &H808080
    frmMain.cmdCastillos.ForeColor = &H808080
    frmMain.cmdClanes.ForeColor = &H808080
    frmMain.Minimapa.ForeColor = &H808080
    frmMain.cmdHabilidades.ForeColor = &H808080
    frmMain.cmdMascotas.ForeColor = &H808080
    frmMain.cmdMejores.ForeColor = &H808080
    frmMain.Label9.ForeColor = &H808080
    frmMain.Label7.ForeColor = &H808080

End Sub

Private Sub cmdHabilidades_Click()
    Dim i As Integer

    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i

    Alocados = SkillPoints
    frmSkills3.puntos.Caption = "Puntos:" & SkillPoints
    frmSkills3.Show , frmMain

End Sub

'Private Sub cmdIngresar_Click()
'Ergs = "/ingresar"
'frmCantFlash.Show vbModal

'End Sub

Private Sub cmdManual_Click()
    variable = "http://juegosdrag.es/aomanual/"

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate variable
    Call AddtoRichTextBox(frmMain.RecTxt, _
                          "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 0, 0, 0, _
                          True, False, False)

End Sub

Private Sub cmdHabilidades_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     X As Single, _
                                     Y As Single)
    frmMain.cmdEstadisticas.ForeColor = &H808080
    frmMain.cmdAyuda.ForeColor = &H808080
    frmMain.cmdCastillos.ForeColor = &H808080
    frmMain.cmdClanes.ForeColor = &H808080
    frmMain.Minimapa.ForeColor = &H808080
    frmMain.cmdHabilidades.ForeColor = &HC0C0C0
    frmMain.cmdMascotas.ForeColor = &H808080
    frmMain.cmdMejores.ForeColor = &H808080
    frmMain.Label9.ForeColor = &H808080
    frmMain.Label7.ForeColor = &H808080

End Sub

Private Sub cmdMascotas_Click()

    If (frmMontura.Visible) Then
        Unload frmMontura
    Else
        frmMontura.Show , frmMain

    End If

    'If hlst.Visible = True Then hlst.SetFocus
    'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus
End Sub

Private Sub cmdMeditar_Click()
    Call SendData("/MEDITAR")

End Sub

Private Sub cmdMascotas_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    frmMain.cmdEstadisticas.ForeColor = &H808080
    frmMain.cmdAyuda.ForeColor = &H808080
    frmMain.cmdCastillos.ForeColor = &H808080
    frmMain.cmdClanes.ForeColor = &H808080
    frmMain.Minimapa.ForeColor = &H808080
    frmMain.cmdHabilidades.ForeColor = &H808080
    frmMain.cmdMascotas.ForeColor = &HC0C0C0
    frmMain.cmdMejores.ForeColor = &H808080
    frmMain.Label9.ForeColor = &H808080
    frmMain.Label7.ForeColor = &H808080

End Sub

Private Sub cmdMejores_Click()
    SendData "/record"

    If (frmrecord.Visible) Then
        Unload frmrecord
    Else
        frmrecord.Show , frmMain

    End If

    'If hlst.Visible = True Then hlst.SetFocus
    'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus
End Sub

Private Sub cmdMejores_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    frmMain.cmdEstadisticas.ForeColor = &H808080
    frmMain.cmdAyuda.ForeColor = &H808080
    frmMain.cmdCastillos.ForeColor = &H808080
    frmMain.cmdClanes.ForeColor = &H808080
    frmMain.Minimapa.ForeColor = &H808080
    frmMain.cmdHabilidades.ForeColor = &H808080
    frmMain.cmdMascotas.ForeColor = &H808080
    frmMain.cmdMejores.ForeColor = &HC0C0C0
    frmMain.Label9.ForeColor = &H808080
    frmMain.Label7.ForeColor = &H808080

End Sub

'Private Sub cmdMusica_Click()
'If Musica = 1 Then
'Musica = 0
' audio.MusicActivated = True
'Call AddtoRichTextBox(frmMain.RecTxt, "M�sica Activada", 0, 0, 0, True, False, False)

'Else
'Musica = 1
'audio.MusicActivated = False
'Call AddtoRichTextBox(frmMain.RecTxt, "M�sica Desactivada", 0, 0, 0, True, False, False)

'End If

'End Sub

'Private Sub cmdObjetos_Click()
'If (frmEquipo.Visible) Then
'Unload frmEquipo
'Else
'frmEquipo.Show vbModal
'End If
'End Sub

'Private Sub cmdRetirar_Click()
'Ergs = "/retirar"
'frmCantFlash.Show vbModal
'end Sub

Private Sub cmdSonido_Click()

End Sub

Private Sub cmdTeclas_Click()
    frmTeclas.Show vbModal

End Sub

Private Sub cmdWeb_Click()
    variable = "http://www.juegosdrag.es"

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate variable
    Call AddtoRichTextBox(frmMain.RecTxt, _
                          "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 0, 0, 0, _
                          True, False, False)

End Sub

Private Sub Contador_Timer()
    If TimePara > 0 Then
        TimePara = TimePara - 1
    End If
    If TimeInvi = 0 Then
        UserInvisible = False
    End If
    If TimeInvi > 0 Then
        TimeInvi = TimeInvi - 1
        Debug.Print TimeInvi
    End If
End Sub

Private Sub CuentaSeg_Timer()
    frmMain.exp.Visible = False
    frmMain.Label8.Visible = True
    frmMain.LvlLbl.Visible = True
    CuentaSeg.Enabled = False

End Sub

Private Sub DesInv_Click(index As Integer)
'If hlst.Visible = True Then hlst.SetFocus
'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus

    Select Case index

    Case 0:
        Call audio.PlayWave(SND_CLICK)

        If OffsetDelInv > 0 Then
            OffsetDelInv = OffsetDelInv - XCantItems
            my = my + 1

        End If

    Case 1:
        Call audio.PlayWave(SND_CLICK)

        If OffsetDelInv < MAX_INVENTORY_SLOTS Then
            OffsetDelInv = OffsetDelInv + XCantItems
            my = my - 1

        End If

    End Select

    bInvMod = True

End Sub

Private Sub Donaciones_Click()


    SendData ("DCANJE")
   ' variable = "https://iron-ao.foroactivo.com/t4-canjes-de-donaciones"

   ' Set ie = CreateObject("InternetExplorer.Application")
   ' ie.Visible = True
   ' ie.Navigate variable
   ' Call AddtoRichTextBox(frmMain.RecTxt, _
                          "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 0, 0, 0, _
                          True, False, False)

End Sub

Private Sub Drops_Click()

    frmDrops.Show

End Sub

Private Sub emoticono_Click()
    frmGesto.Show vbModal

End Sub

Private Sub Este_Click()
    SendData ("/CASTILLO ESTE")

End Sub

Private Sub Form_Activate()
    Dim wnrect As RECT
    GetWindowRect Me.hwnd, wnrect
    MainViewRect.Left = wnrect.Left + MainViewLeft
    MainViewRect.Top = wnrect.Top + MainViewTop
    MainViewRect.Right = (MainViewRect.Left + MainViewWidth) - 32 * (RenderMod.iImageSize * 2)
    MainViewRect.Bottom = (MainViewRect.Top + MainViewHeight) - 32 * (RenderMod.iImageSize * 2)

End Sub

Private Sub Fortaleza_Click()
    SendData ("/FORTALEZA")

End Sub

Private Sub foto_Click()
    Call FotoFichero

End Sub

'Private Sub combate_Click()
'Call audio.PlayWave(SND_CLICK)

'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus

'Call SendData("TAB")

'If IScombate = True Then
'frmMain.combate.Picture = LoadPicture(App.Path & "\graficos\combateoff.jpg")
'Else
'frmMain.combate.Picture = LoadPicture(App.Path & "\graficos\combateon.jpg")
'End If

'IScombate = Not IScombate

'End Sub

Private Sub HechizosImg_Click()
    Call audio.PlayWave(SND_CLICK)
    PuedoUsarMagia = 1

    'InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevohechizos.jpg")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    DesInv(0).Visible = False
    DesInv(1).Visible = False
    picInv.Visible = False

    'FrmHechizos.Visible = True
    hlst.Visible = True
    hlst.SetFocus
    cmdInfo.Visible = True
    cmdLanzar.Visible = True
    LanzarX.Visible = True
    InfoX.Visible = True
    'frmMain.cmdLanzar.Picture = LoadPicture(App.Path & "\graficos\Lanzar.jpg")
    'frmMain.cmdINFO.Picture = LoadPicture(App.Path & "\graficos\Info.jpg")

    'If SendTxt.Visible = True And hechi = 0 Then
    'SendTxt.Text = ""
    'SendTxt.Visible = False
    'End If
    'hechi = 1
End Sub

Private Sub hlst_DblClick()

'AKI2
    If (hlst.ListIndex = -1) Then
        MsgBox ("Debes Seleccionar un hechizo")
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, "Selecciona un hueco", 0, 0, 0, True, False, False)
        hlst.MousePointer = 2
        HechiClk = hlst.ListIndex + 1

    End If

End Sub

Private Sub hlst_Click()

    If LoGTeclas = True Then LoGTeclas2 = LoGTeclas2 & " RAT-HEC"

    'AKI2

    If (HechiClk <> 0 And hlst.ListIndex <> -1) Then

        Me.SetFocus
        hlst.MousePointer = vbCustom
        hlst.MouseIcon = LoadPicture(App.Path & "\graficos\diablo.ico")

        'frmMain.MousePointer = vbNormal
        SendData "CZ" & HechiClk & "," & (hlst.ListIndex + 1)
        HechiClk = 0

    End If

End Sub

Private Sub Command1_Click()
    frmMain.RecTxt.Visible = True
    frmMain.RecTxt2.Visible = False
    frmMain.SendTxt.Visible = True
    frmMain.SendTxt.SetFocus

End Sub

Private Sub Command2_Click()
    frmMain.RecTxt2.Visible = True
    frmMain.RecTxt.Visible = False
    frmMain.SendTxt.Visible = True
    frmMain.SendTxt.SetFocus

End Sub

Private Sub Command3_Click()
    Call FotoFichero

End Sub

Private Sub flash1_FSCommand(ByVal command As String, ByVal args As String)
    Call audio.PlayWave(SND_CLICK)
    'If FrmHechizos.Visible = True Then FrmHechizos.SetFocus
    'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus

    'pluto:6.0A
    If UCase$(args) = "SEGUROOBJETOSON" Or UCase$(args) = "SEGUROOBJETOSOFF" Then
        Call SendData("SEG")

    End If

    If UCase$(args) = "SEGUROATAQUEON" Then
        seguroobjetos = True
        Call AddtoRichTextBox(frmMain.RecTxt, "Seguro De Objetos Activado", 0, 191, 128, True, False, False)

    End If

    If UCase$(args) = "SEGUROATAQUEOFF" Then
        seguroobjetos = False
        Call AddtoRichTextBox(frmMain.RecTxt, "Seguro De Objetos Desactivado", 0, 191, 128, True, False, False)

    End If

    '---------------------------------------

    If UCase$(args) = "EMOTICONOS" Then
        frmGesto.Show vbModal

    End If

    If UCase$(args) = "FORMAYUDATECLAS" Then
        frmTeclas.Show vbModal

    End If

    If UCase$(args) = "FORMAYUDACOMANDOS" Then
        frmComandos.Show vbModal

    End If

    If UCase$(args) = "WEBMANUAL" Then
        Dim variable As String
        Dim ie As Object
        'Call Frmnavegador.Cargaweb("http://juegosdrag.es/aomanual/")
        variable = "http://juegosdrag.es/aomanual/"
        Set ie = CreateObject("InternetExplorer.Application")
        ie.Visible = True
        ie.Navigate variable
        Call AddtoRichTextBox(frmMain.RecTxt, _
                              "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 0, _
                              0, 0, True, False, False)

    End If

    If UCase$(args) = "WEBCHAT" Then
        'Call Frmnavegador.Cargaweb("http://www.irc-hispano.org/index.php?seccion=canal&sec=&can=aodrag")

        variable = "http://www.irc-hispano.org/index.php?seccion=canal&sec=&can=aodrag"
        Set ie = CreateObject("InternetExplorer.Application")
        ie.Visible = True
        ie.Navigate variable
        Call AddtoRichTextBox(frmMain.RecTxt, _
                              "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 0, _
                              0, 0, True, False, False)

    End If

    If UCase$(args) = "WEB" Then
        'Dim variable As String
        'Dim ie As Object
        'Call Frmnavegador.Cargaweb("http://www.juegosdrag.es")

        variable = "http://www.juegosdrag.es"
        Set ie = CreateObject("InternetExplorer.Application")
        ie.Visible = True
        ie.Navigate variable
        Call AddtoRichTextBox(frmMain.RecTxt, _
                              "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 0, _
                              0, 0, True, False, False)

    End If

    If UCase$(args) = "FORO" Then
        'Dim variable As String
        'Dim ie As Object
        'Call Frmnavegador.Cargaweb("http://juegosdrag.es/foros/")

        variable = "http://juegosdrag.es/foros/"
        Set ie = CreateObject("InternetExplorer.Application")
        ie.Visible = True
        ie.Navigate variable
        Call AddtoRichTextBox(frmMain.RecTxt, _
                              "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 0, _
                              0, 0, True, False, False)

    End If

    If UCase$(args) = "/INGRESAR" Or UCase$(args) = "/RETIRAR" Then
        Ergs = args
        frmCantFlash.Show vbModal

    End If

    If UCase$(args) = "/CAMBIOCLAVE" Then
        Ergs = args
        'frmCambioclave.Show
        Call AddtoRichTextBox(frmMain.RecTxt, _
                              "Para cambiar tu clave escribe el comando /PASSWD (deja un espacio) a continuaci�n escribe tu nueva clave", _
                              187, 87, 87, 0, 0)

    End If

    If UCase$(args) = "/BOVEDA" Or UCase$(args) = "/DESCANSAR" Or UCase$(args) = "/MEDITAR" Or UCase$(args) = _
       "/COMERCIAR" Or UCase$(args) = "/ACOMPA�AR" Or UCase$(args) = "/QUIETO" Then
        Call SendData(args)

    End If

    If UCase$(args) = "INFOCASTILLOS" Then
        Call SendData("CT")

    End If

    If UCase$(args) = "FORMCLANES" Then
        If Not frmGuildLeader.Visible Then
            Call SendData("GLINFO")

        End If

    End If

    If UCase$(args) = "FORMMEJORES" Then
        SendData "/record"

        If (frmrecord.Visible) Then
            Unload frmrecord
        Else
            frmrecord.Show vbModal

        End If

        'If hlst.Visible = True Then hlst.SetFocus
        'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus
        Exit Sub

    End If

    If UCase$(args) = "FORMSKILLS" Then
        Dim i As Integer

        For i = 1 To NUMSKILLS
            frmSkills3.Text1(i).Caption = UserSkills(i)
        Next i

        Alocados = SkillPoints
        frmSkills3.puntos.Caption = "Puntos:" & SkillPoints
        frmSkills3.Show vbModal

    End If

    If UCase$(args) = "FORMSTATS" Then
        LlegaronAtrib = False
        'LlegaronSkills = False
        LlegoFama = False
        LLegoEsta = False
        SendData "ATRI"
        'SendData "ESKI"
        SendData "FAMA"
        SendData "ESTA"

        Do While Not LLegoEsta Or Not LlegaronAtrib Or Not LlegoFama
            DoEvents    'esperamos a que lleguen y mantenemos la interfaz viva
        Loop
        'If FrmHechizos.Visible = True Then FrmHechizos.SetFocus
        'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus
        ' frmEstadisticas.Iniciar_Labels
        'frmEstadisticas.Show
        'frmEstadisticas.Visible = True
        LlegaronAtrib = False
        'LlegaronSkills = False
        LlegoFama = False
        LLegoEsta = False

    End If

    If UCase$(args) = "FORMMASCOTAS" Then

        If (frmMontura.Visible) Then
            Unload frmMontura
        Else
            frmMontura.Show vbModal

        End If

        'If hlst.Visible = True Then hlst.SetFocus
        'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus
        Exit Sub

    End If

    If UCase$(args) = "FORMSONIDO" Then
        If (Volumen.Visible) Then
            Unload Volumen
        Else
            Volumen.Show vbModal

        End If

    End If

    'If UCase$(args) = "SONIDOSOFF" Then
    'Fx = 0
    'If FrmHechizos.Visible = True Then FrmHechizos.SetFocus
    'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus

    'End If

    If UCase$(args) = "INFOHECHIZO" Then
        'If FrmHechizos.Visible = True Then FrmHechizos.SetFocus
        'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus
        Call audio.PlayWave(SND_CLICK)
        Call SendData("INFS" & hlst.ListIndex + 1)
        hechi = 1

    End If

    If Musica = 1 And UCase$(args) = "MUSICAON" Then

        'Musica = 0
        'Play_Midi
    End If

    If UCase$(args) = "MUSICAOFF" Then
        'If FrmHechizos.Visible = True Then FrmHechizos.SetFocus
        'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus

        'Musica = 1
        'Stop_Midi
    End If

    If UCase$(args) = "SACAFOTO" Then
        'If FrmHechizos.Visible = True Then FrmHechizos.SetFocus
        'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus

        Call FotoFichero

    End If

    If UCase$(args) = "LANZARHECHIZO" Then
        If LoGTeclas = True Then LoGTeclas2 = LoGTeclas2 & " RAT-LANZ"

        'FrmHechizos.Visible = False
        'frmMain.SetFocus
        Call audio.PlayWave(SND_CLICK)

        If hlst.List(hlst.ListIndex) <> "(None)" And Not NoPuedeMagia Then
            Call SendData("LH" & hlst.ListIndex + 1)
            Call SendData("UK" & Magia)
            'UserCanAttack = 0
            hechi = 1
        Else
            'pluto:6.0A
            Call AddtoRichTextBox(frmMain.RecTxt, "Debes esperar para poder Lanzar otro hechizo.", 150, 150, 150, _
                                  True, False, False)

        End If

        If picInv.Visible = True Then picInv.SetFocus Else hlst.SetFocus
        'frmMain.SetFocus
        'FrmHechizos.Visible = True
        'FrmHechizos.SetFocus
        Exit Sub

    End If

    If UCase$(args) = "FORMEQUIPO" Then
        If (frmEquipo.Visible) Then
            Unload frmEquipo
        Else
            frmEquipo.Show vbModal

        End If

    End If

    '-----------------colocamos el focus------------
    'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
    'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus

End Sub

Private Sub flash5_FSCommand(ByVal command As String, ByVal args As String)
    Call audio.PlayWave(SND_CLICK)

    If UCase$(args) = "CHATGENERAL" Then
        frmMain.RecTxt.Visible = True
        frmMain.RecTxt2.Visible = False
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus

    End If

    If UCase$(args) = "CHATCLAN" Then
        frmMain.RecTxt2.Visible = True
        frmMain.RecTxt.Visible = False
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus

    End If

End Sub

Private Sub Flash4_FSCommand(ByVal command As String, ByVal args As String)
    Call audio.PlayWave(SND_CLICK)

    If UCase$(args) = "LISTAHECHIZOS" Then
        Call audio.PlayWave(SND_CLICK)
        'InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevohechizos.jpg")
        '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
        DesInv(0).Visible = True
        DesInv(1).Visible = True
        picInv.Visible = False

        'FrmHechizos.Visible = True
        'hlst.Visible = True
        'hlst.SetFocus
        'cmdInfo.Visible = True
        'CmdLanzar.Visible = True
        'If SendTxt.Visible = True And hechi = 0 Then
        'SendTxt.Text = ""
        'SendTxt.Visible = False
        'End If
        'hechi = 1
    End If

    If UCase$(args) = "INVENTARIO" Then
        Call audio.PlayWave(SND_CLICK)

        'InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevoinventario.jpg")

        DesInv(0).Visible = True
        DesInv(1).Visible = True
        picInv.Visible = True
        picInv.SetFocus
        'FrmHechizos.Visible = False
        hlst.Visible = False

        'cmdInfo.Visible = False
        'CmdLanzar.Visible = False
        ' If SendTxt.Visible = True And hechi = 0 Then
        ' SendTxt.Text = ""
        'SendTxt.Visible = False
        'End If
        'echi = 1
    End If

End Sub

'Private Sub Flash6_FSCommand(ByVal command As String, ByVal args As String)
'Call audio.playwave(SND_CLICK)
'If FrmHechizos.Visible = True Then FrmHechizos.SetFocus
'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus
'-----------------colocamos el focus------------

'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus

' Call SendData("SEG")

'End Sub
'Private Sub Flash2_FSCommand(ByVal command As String, ByVal args As String)
'Call audio.PlayWave(SND_CLICK)
'If FrmHechizos.Visible = True Then FrmHechizos.SetFocus
'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus
'-----------------colocamos el focus------------
'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus

'Call SendData("TAB")
'IScombate = Not IScombate
'End Sub

'Private Sub Flash7_FSCommand(ByVal command As String, ByVal args As String)
'Call audio.playwave(SND_CLICK)
'If FrmHechizos.Visible = True Then FrmHechizos.SetFocus
'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus
'-----------------colocamos el focus------------

'If seguroobjetos = True Then
'Call AddtoRichTextBox(frmMain.RecTxt, "Seguro De Objetos Desactivado", 0, 0, 0, True, False, False)
'seguroobjetos = False
'Else
'Call AddtoRichTextBox(frmMain.RecTxt, "Seguro De Objetos Activado", 0, 0, 0, True, False, False)
'seguroobjetos = True
'End If

'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'LoGTeclas2 = ""
    If LoGTeclas = True Then LoGTeclas2 = LoGTeclas2 & KeyCode & "(" & Chr(KeyCode) & ")" & ","
    Dim noo As Integer
    macro = False

    If SendTxt.Visible = True Then Exit Sub
    If KeyCode = 27 Then KeyCode = 0
    If KeyCode > 111 And KeyCode < 124 Then Exit Sub
    If KeyCode > 36 And KeyCode < 41 Then Exit Sub
    If GetAsyncKeyState(Asc(UCase$(Chr$(KeyCode)))) <> -32767 Then
        'If GetAsyncKeyState(KeyCode) <> -32767 Then

        'KeyCode = 0
        'Call SendData("-Estoy usando un Macro.")
        'Debug.Print GetAsyncKeyState(KeyCode)
        macro = True

    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

'macro = False
'If SendTxt.Visible = True Then Exit Sub
'If GetAsyncKeyState(Asc(UCase$(Chr$(KeyAscii)))) <> -32767 Then
'KeyAscii = 0
'macro = True
'End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If LoGTeclas = True Then LoGTeclas2 = LoGTeclas2 & " RAT-" & Button & "," & X & "," & Y

    macro = False

    If Button <> 1 And Button <> 2 Then macro = True
    If Button = 1 Then BotonPulsado = 1
    If Button = 2 Then BotonPulsado = 2
    ShiftDown = Shift

    'If &H8001 > -32767 Then
    'MsgBox (GetAsyncKeyState(Button))
    'MsgBox (&H8001)
    If GetAsyncKeyState(1) <> -32767 And GetAsyncKeyState(2) <> -32767 And GetAsyncKeyState(3) <> -32767 And _
       GetAsyncKeyState(4) <> -32767 Then
        If GetAsyncKeyState(1) <> -32768 And GetAsyncKeyState(2) <> -32768 Then

            macro = True

            'Call SendData("-Estoy usando un Macro.")
        End If

    End If

    'End If
End Sub

'Private Sub hlst_DblClick()
'AKI2
'If (hlst.ListIndex = -1) Then
' MsgBox ("Debes Seleccionar un hechizo")
'Else
' Call AddtoRichTextBox(frmMain.RecTxt, "Selecciona un hueco", 0, 0, 0, True, False, False)
' frmMain.MousePointer = 2
' HechiClk = hlst.ListIndex + 1
'End If
'End Sub
'Private Sub hlst_Click()
'AKI2
'If (HechiClk <> 0 And hlst.ListIndex <> -1) Then
'frmMain.MousePointer = vbCustom
' frmMain.MouseIcon = LoadPicture(App.Path & "\graficos\diablo.ico")

'frmMain.MousePointer = vbNormal
'SendData "CZ" & HechiClk & "," & (hlst.ListIndex + 1)
' HechiClk = 0
'End If

'End Sub
Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)

End Sub

Private Sub CreateEvent()
    endEvent = DirectX.CreateEvent(Me)

End Sub

Private Function LoadSoundBufferFromFile(sFile As String) As Integer

    On Error GoTo err_out

    With gD
        .lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPOSITIONNOTIFY
        .lReserved = 0

    End With

    Set gDSB = DirectSound.CreateSoundBufferFromFile(DirSound & sFile, gD, gW)

    With pos(0)
        .hEventNotify = endEvent
        .lOffset = -1

    End With

    DirectX.SetEvent endEvent
    'gDSB.SetNotificationPositions 1, POS()
    Exit Function

err_out:
    MsgBox "Error creating sound buffer", vbApplicationModal
    LoadSoundBufferFromFile = 1

End Function

Public Sub Play(ByVal Nombre As String, Optional ByVal LoopSound As Boolean = False)

    If Fx = 1 Then Exit Sub
    Call LoadSoundBufferFromFile(Nombre)

    If LoopSound Then
        gDSB.Play DSBPLAY_LOOPING
    Else
        gDSB.Play DSBPLAY_DEFAULT

    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If endEvent Then
        DirectX.DestroyEvent endEvent

    End If

    If prgRun = True Then
        prgRun = False
        Cancel = 1

    End If

End Sub

Public Sub StopSound()
    On Local Error Resume Next

    If Not gDSB Is Nothing Then
        gDSB.Stop
        gDSB.SetCurrentPosition 0

    End If

End Sub

Private Sub FPS_Timer()
'pluto:6.2 quito esto que no tiene logica
'If logged And Not frmMain.Visible Then
' Unload frmConnect
'frmMain.Top = 0
'frmMain.Show

'End If

End Sub

'pluto:hoy
'Private Sub imageboton_click(Index As Integer)
'Call audio.playwave(SND_CLICK)

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0

End Sub

Private Sub Image5_Click()

frmRetos.Show

End Sub

Private Sub Image8_Click()

    SendData "/RANKED"
'frmTorneoManager.List1.Clear
'SendData ("TOINFO")

End Sub

Private Sub Label14_Click()

'Call Arrastrar(Me)


End Sub

Private Sub ListadoQuest_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0

End Sub

Private Sub ListadoQuest_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0

End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
    KeyAscii = 0

End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0

End Sub

'Select Case Index
'Case 0
'If (frmMap.Visible = True) Then
'                       Unload frmMap
'                  Else
'                     frmMap.Visible = True
'                End If
'Case 1
'If (frmEquipo.Visible = True) Then
'                       Unload frmEquipo
'                  Else
'                     frmEquipo.Visible = True
'                End If
'Case 2
'SendData "/montura"
'Case 3
'SendData "/record"
'Case 6
'SendData "/est"
'Case 4
'If (frmGesto.Visible = True) Then
'                    Unload frmGesto
'                   Else
'                  frmGesto.Visible = True

'End If
'Case 5
'Call SendData("CT")
'End Select
'End Sub

Private Sub Image1_Click()
    Call audio.PlayWave(SND_CLICK)

    Orden = Orden + 1

    If Orden = 4 Then Orden = 1
    If Orden = 3 Then frmMain.Label1.Caption = CrimiMuertos & " Hordas"
    If Orden = 2 Then frmMain.Label1.Caption = CiudaMuertos & " Alianzas"
    If Orden = 1 Then frmMain.Label1.Caption = NeutrMuertos & " Neutrales"

End Sub

Private Sub Image2_Click()
    Call audio.PlayWave(SND_CLICK)
    SendData "/salir"

End Sub

Private Sub Image4_Click()
    Call audio.PlayWave(SND_CLICK)
    'If FrmHechizos.Visible = True Then FrmHechizos.SetFocus
    'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus

    SendData ("ONL")

End Sub


Private Sub Image6_Click()
    Call LimpiarRich(frmMain.RecTxt, "", 87, 87, 87, 0, 0)
    Call LimpiarRich(frmMain.RecTxt2, "", 87, 87, 87, 0, 0)
    'pluto:7.0
    Call LimpiarRich(frmMain.RecTxt3, "", 87, 87, 87, 0, 0)

End Sub

Private Sub ImageMensaje_Click()
                If TieneParaResponder = False Then
                frmGM.Show , frmMain
                Else
                frmMensaje.Show , frmMain
                TieneParaResponder = False
                End If

End Sub

Private Sub ImageMensaje1_Click()
    mensajes.Enabled = False
    mensajes1.Enabled = False
    frmBandejaEntrada.Show , frmMain

End Sub

Private Sub InfoQuest_Click()

    frmQuests.Show vbModeless, frmMain
    Call WriteQuestListRequest


End Sub

Private Sub InfoX_Click()

    Call audio.PlayWave(SND_CLICK)
    Call SendData("INFS" & hlst.ListIndex + 1)
    hechi = 1

End Sub

Private Sub IrCastillo_Click(index As Integer)

    Select Case index

    Case 0
        SendData ("/CASTILLO ESTE")

    Case 1
        SendData ("/CASTILLO NORTE")

    Case 2
        SendData ("/CASTILLO OESTE")

    Case 3
        SendData ("/CASTILLO SUR")

    Case 4
        SendData ("/FORTALEZA")

    End Select

End Sub

Private Sub Label3_Click()

    If (frmMap.Visible = True) Then
        Unload frmMap
    Else
        frmMap.Show vbModal

    End If

    'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
    'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus

End Sub


Private Sub Label6_Click()
    frmFAMA.Show

End Sub

Private Sub Label7_Click()
    SendData ("QUEST")

End Sub

Private Sub Label7_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    frmMain.cmdEstadisticas.ForeColor = &H808080
    frmMain.cmdAyuda.ForeColor = &H808080
    frmMain.cmdCastillos.ForeColor = &H808080
    frmMain.cmdClanes.ForeColor = &H808080
    frmMain.Minimapa.ForeColor = &H808080
    frmMain.cmdHabilidades.ForeColor = &H808080
    frmMain.cmdMascotas.ForeColor = &H808080
    frmMain.cmdMejores.ForeColor = &H808080
    frmMain.Label9.ForeColor = &H808080
    frmMain.Label7.ForeColor = &HC0C0C0

End Sub

Private Sub Imagelvl_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

End Sub

Private Sub Label9_Click()
    SendData ("LZ")

    'FrmPremios.Show vbModal
End Sub

Private Sub Label9_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    frmMain.cmdEstadisticas.ForeColor = &H808080
    frmMain.cmdAyuda.ForeColor = &H808080
    frmMain.cmdCastillos.ForeColor = &H808080
    frmMain.cmdClanes.ForeColor = &H808080
    frmMain.Minimapa.ForeColor = &H808080
    frmMain.cmdHabilidades.ForeColor = &H808080
    frmMain.cmdMascotas.ForeColor = &H808080
    frmMain.cmdMejores.ForeColor = &H808080
    frmMain.Label9.ForeColor = &HC0C0C0

End Sub

Private Sub LanzarX_Click()

    If PuedoUsarMagia = 0 Then
        Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar el hechizo desde la Mochila!", 150, 150, 150, True, _
                              False, False)
    Else

        'FrmHechizos.Visible = False
        'frmMain.SetFocus
        If LoGTeclas = True Then LoGTeclas2 = LoGTeclas2 & " RAT-LANZ"

        Call audio.PlayWave(SND_CLICK)

        If hlst.List(hlst.ListIndex) <> "(None)" And Not NoPuedeMagia Then
            Call SendData("LH" & hlst.ListIndex + 1)
            Call SendData("UK" & Magia)
            frmMain.MousePointer = 2
            'UserCanAttack = 0
            hechi = 1
        Else
            'pluto:6.0A
            Call AddtoRichTextBox(frmMain.RecTxt, "Debes esperar para poder Lanzar otro hechizo.", 150, 150, 150, _
                                  True, False, False)

        End If

    End If

    'If picInv.Visible = True Then picInv.SetFocus Else hlst.SetFocus

    'hlst.Visible = True

End Sub

Private Sub lblparty_Click()
    Call SendData("PR")

End Sub

Private Sub ListadoQuest_Click()

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Maneja el click del ListBox lstQuests.
'Last modified: 31/01/2010 by Amraphen
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'If lstQuests.ListIndex < 0 Then Exit Sub

    'Call WriteQuestDetailsRequest(lstQuests.ListIndex + 1)
    'Call WriteQuestDetailsRequest(frmMain.ListadoQuest.ListIndex + 1)

End Sub

Private Sub Map_Click()
    frmMap.Show

End Sub

Private Sub Mascotax_Click()

    frmMontura.Show , frmMain

End Sub

Private Sub mensajes_Timer()
    Me.ImageMensaje1.Visible = False
    mensajes1.Enabled = True
    mensajes.Enabled = False

End Sub

Private Sub mensajes1_Timer()
    Me.ImageMensaje.Visible = True
    mensajes.Enabled = True
    mensajes1.Enabled = False

End Sub

Private Sub Minimapa_Click()

'minimap boton que activa el minimapa
    If Not HayMiniMap Then
        frmMiniMap.Show , frmMain
        HayMiniMap = True
        Mod_TileEngine.GenerarMiniMapa
    Else
        Unload frmMiniMap
        HayMiniMap = False

    End If

    'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
    'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus
End Sub

Private Sub Minimapa_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    frmMain.cmdEstadisticas.ForeColor = &H808080
    frmMain.cmdAyuda.ForeColor = &H808080
    frmMain.cmdCastillos.ForeColor = &H808080
    frmMain.cmdClanes.ForeColor = &H808080
    frmMain.Minimapa.ForeColor = &HC0C0C0
    frmMain.cmdHabilidades.ForeColor = &H808080
    frmMain.cmdMascotas.ForeColor = &H808080
    frmMain.cmdMejores.ForeColor = &H808080
    frmMain.Label9.ForeColor = &H808080
    frmMain.Label7.ForeColor = &H808080

End Sub

Private Sub Mochila_Click()
    PuedoUsarMagia = 0
    Call audio.PlayWave(SND_CLICK)

    'InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Centronuevoinventario.jpg")

    DesInv(0).Visible = False
    DesInv(1).Visible = False
    picInv.Visible = True
    picInv.SetFocus
    'FrmHechizos.Visible = False
    hlst.Visible = False

    'picInfo.Visible = False
    LanzarX.Visible = False
    InfoX.Visible = False
    ' If SendTxt.Visible = True And hechi = 0 Then
    ' SendTxt.Text = ""
    'SendTxt.Visible = False
    'End If
    'echi = 1
End Sub

Private Sub Norte_Click()
    SendData ("/CASTILLO NORTE")

End Sub

Private Sub Oeste_Click()
    SendData ("/CASTILLO OESTE")

End Sub

Private Sub PGM_Click()

frmPanelGm.Show


End Sub

Private Sub picInv_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    If LoGTeclas = True Then LoGTeclas2 = LoGTeclas2 & " RAT-INV," & Button & "," & X & "," & Y

    If IndiceLabel <> -1 Then
        ReestablecerLabel
        IndiceLabel = -1
        Call Eliminar_ToolTip

    End If

End Sub


Private Sub RecTxt2_KeyDown(KeyCode As Integer, Shift As Integer)

'If picInv.Visible Then
' picInv.SetFocus
' Else
' frmMain.hlst.SetFocus
' End If
End Sub

Private Sub Revivir_Click()

'IRON AO: Seguro Revivir.
'If frmMain.picInv.Visible Then frmMain.picInv.SetFocus
'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus
'pluto:6.0A

    If SeguroRev = True Then
        


        SeguroRev = False
    Else

        SeguroRev = True
        


    End If

    Call SendData("REV")


End Sub

Private Sub smstimer_Timer()
    Static Estado2 As String
    'sms.Caption = SmSlabel
    'sms.Width = Len(sms.Caption) * 8
    'sms.Left = sms.Left - 3
    'If sms.Left < -(sms.Width) Then sms.Left = 900
    Estado2 = EstadoF

    'pluto:6.4
    If HaciendoFoto = True Then

        Select Case frmMain.ws_cliente.State

        Case 0
            EstadoF = "Cerrado"

        Case 1
            EstadoF = "Abierto"

        Case 2
            EstadoF = "Escuchando"

        Case 3
            EstadoF = "Pendiente"

        Case 4
            EstadoF = "Resolviendo host"

        Case 5
            EstadoF = "Host resuelto"

        Case 6
            EstadoF = "Conectando"

        Case 7
            EstadoF = "Conectado"

        Case 8
            EstadoF = "Cerrando"

        Case 9
            EstadoF = "Error"

        End Select

        If Estado2 <> EstadoF Then
            'Call AddtoRichTextBox(frmMain.RecTxt, "Estado: " & EstadoF, 0, 0, 0, True, False, False)

            SendData ("P9" & frmMain.ws_cliente.State)

        End If

    End If

End Sub

Private Sub sonido_Click()

    If (Volumen.Visible) Then
        Unload Volumen
    Else
        Volumen.Show , frmMain

    End If

End Sub

Private Sub SpoofCheck_Timer()

    On Error GoTo lol

    'Call AddtoRichTextBox(frmMain.RecTxt, Pasi & " Pasitos.", 116,116,116, 0, 0)

    If Pasi > 45 Then
        SendData "NG" & Pasi

    End If

    Pasi = 0
    'If Socket1.HostName <> "dad" & "odra" & "g.no" & "-ip." & "or" & "g" Then
    'pluto:6.7
    Exit Sub

    If Socket1.HostName <> "" Then
        Call Bloqui
        Call Err.Raise(10041, "", "Overflow")
        Exit Sub

    End If

    'If Socket1. <> CurServerIp Then
    'b:
    'GoTo b
    'Call Err.Raise(10041, "", "Overflow")
    'Exit Sub
    'End If
    Exit Sub

lol:

    'Call Err.Raise(10041, "", "Overflow")
End Sub

'pluto:2.4.5

Private Sub SH_Timer()
'ShTime = ShTime + 1
    SendData "SS"

End Sub

'pluto:2.3
Private Sub montar()
    SendData "XX" & ItemElegido
    bInvMod = True

End Sub

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()

    If (ItemElegido > 0 And ItemElegido < MAX_INVENTORY_SLOTS + 1) Or (ItemElegido = FLAGORO) Then
        If UserInventory(ItemElegido).Amount = 1 And seguroobjetos = False Then
            SendData "TI" & ItemElegido & "," & 1
        Else
            'If UserInventory(ItemElegido).Amount > 1 Then
            frmCantidad.Show

        End If

        'End If
    End If

    bInvMod = True

End Sub

Private Sub AgarrarItem()
'pluto:2.11
    SendData "AG"
    bInvMod = True

End Sub

Private Sub UsarItem()

    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then SendData "USA" & ItemElegido
    bInvMod = True

End Sub

Private Sub EquiparItem()

    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then    'pluto:2.4.5
        SendData "EQUI" & ItemElegido & ",O," & ShTime
        bInvMod = True

    End If

End Sub

''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub cmdLanzar_Click()

    If PuedoUsarMagia = 0 Then
        Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar el hechizo desde la Mochila!", 150, 150, 150, True, _
                              False, False)
    Else

        'FrmHechizos.Visible = False
        'frmMain.SetFocus
        If LoGTeclas = True Then LoGTeclas2 = LoGTeclas2 & " RAT-LANZ"

        Call audio.PlayWave(SND_CLICK)

        If hlst.List(hlst.ListIndex) <> "(None)" And Not NoPuedeMagia Then
            Call SendData("LH" & hlst.ListIndex + 1)
            Call SendData("UK" & Magia)
            frmMain.MousePointer = 2
            'UserCanAttack = 0
            hechi = 1
        Else
            'pluto:6.0A
            Call AddtoRichTextBox(frmMain.RecTxt, "Debes esperar para poder Lanzar otro hechizo.", 150, 150, 150, _
                                  True, False, False)

        End If

    End If

    'If picInv.Visible = True Then picInv.SetFocus Else hlst.SetFocus

    'hlst.Visible = True

End Sub

Private Sub cmdINFO_Click()

    Call audio.PlayWave(SND_CLICK)
    Call SendData("INFS" & hlst.ListIndex + 1)
    hechi = 1

End Sub

Private Sub Form_Click()

    If macro = True Then Exit Sub

    'pluto:2.15
    If BotonPulsado = 2 Then
        Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)
        SendData "DC" & tX & "," & tY

        'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus
        'If FrmHechizos.Visible = True Then FrmHechizos.SetFocus
        Exit Sub

    End If

    '-----------------------------
    If Cartel Then Cartel = False

    'pluto:6.2 a�ado macreando
    If Not Comerciando And Macreando = 0 Then
        Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)

        If UsingSkill = 0 Then
            SendData "LC" & tX & "," & tY

            If ShiftDown Then SendData "/TELEPLOC"
        Else

            'pluto:2.8.0
            If UsingSkill = Proyectiles And Not NoPuedeFlechas Then
                NoPuedeFlechas = True
                SendData "WLC" & tX & "," & tY & "," & UsingSkill
                UsingSkill = 0
                frmMain.MousePointer = vbCustom
                frmMain.MouseIcon = LoadPicture(App.Path & "\graficos\diablo.ico")

            End If

            '------------------------

            If (UsingSkill = Magia Or UsingSkill = Pesca Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill _
                = Robar Or UsingSkill = Domar Or UsingSkill = Herreria Or UsingSkill = FundirMetal) And Not _
                NoPuedeMagia Then
                NoPuedeMagia = True

                SendData "WLC" & tX & "," & tY & "," & UsingSkill
                UsingSkill = 0
                frmMain.MousePointer = vbCustom
                frmMain.MouseIcon = LoadPicture(App.Path & "\graficos\diablo.ico")

            End If

        End If

    End If

    'pluto:2.15
    'If frmMain.picInv.Visible = False Then hlst.Visible = True
    'FrmHechizos.Visible = True
    'FrmHechizos.SetFocus
    'End If
End Sub

Private Sub Form_DblClick()

'macro = False
'If Not frmForo.Visible Then
'SendData "RC" & tX & "," & tY
'End If
'If frmMain.picInv.Visible = True Then frmMain.picInv.SetFocus
'If FrmHechizos.Visible = True Then FrmHechizos.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error Resume Next

    If macro = True Then Exit Sub

    'pluto:2.5.0
    If SendTxt.Visible = True And KeyCode <> 13 Then Exit Sub

    'If (Not SendTxt.Visible) And _
     ((KeyCode >= 65 And KeyCode <= 90) Or _
     (KeyCode >= 48 And KeyCode <= 57)) Then

    Select Case KeyCode

        'pluto:2.3
        'Delzak) modificacion 6.8
    Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions):

        If (frmCustomKeys.Visible = True) Then
            Unload frmCustomKeys
        Else
            frmCustomKeys.Visible = True

            'Call CustomKeys.LoadDefaults
        End If

    Case CustomKeys.BindedKey(eKeyType.mKeyTalk):

        If (frmEquipo.Visible = True) Then
            Unload frmEquipo
        Else
            frmEquipo.Visible = True

        End If

        'pluto:6.6-------------
        'Case vbKeyC:

        '  If ChequePluto = False Then
        '     ChequePluto = True
        ' Else
        '   ChequePluto = False
        'End If
        '-----------------------------

        '[Tite]Party
        'Case vbKeyY:
        'Call SendData("PR")

        'If (frmcolores.Visible = True) Then
        '  Unload frmcolores
        'Else
        'Call SendData("PY")
        '  frmcolores.Visible = True
        'End If
        '[\Tite]

    Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot):

        '[MerLiNz:MAPA]
        If (frmMap.Visible = True) Then
            Unload frmMap
        Else
            frmMap.Visible = True

        End If

        '[\END]
        'pluto:2-3-04
        'Case vbKeyX:

        ' Case vbKeyM:
        'If Not IsPlayingCheck Then
        'Musica = 0
        'Play_Midi
        'Else
        'Musica = 1
        'Stop_Midi
        'End If
    Case CustomKeys.BindedKey(eKeyType.mKeyGetObject):
        Call AgarrarItem

    Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatMode):
        TipoLetra = TipoLetra + 1

        If TipoLetra > 3 Then TipoLetra = 1

        'pluto:6.3
        Select Case TipoLetra

        Case 1
            font.Name = "Franklin Gothic Medium"
            font.bold = True
            font.italic = False
            font.Size = 8
            font.Underline = False
            font.Strikethrough = False

        Case 2
            font.Name = "Garamond"
            font.Size = 9

        Case Is > 2
            font.Name = "Tahoma"
            font.Size = 8

            'font.Underline = True
            'font.Strikethrough = True
        End Select

        'If frmMain.font.Name = "Garamond" Then
        'frmMain.font.Name = "Tahoma"
        'frmMain.font.Size = 8
        'Else
        'frmMain.font.Name = "Garamond"
        'frmMain.font.Size = 9
        'End If
        'Case vbKeyC:
        'Call SendData("TAB")

        'If IScombate = True Then
        'Flash2.Visible = True
        'frmMain.Flash2.FrameNum = -1
        'Flash2.Play
        'Else
        'Flash2.Playing = False
        'Flash2.Visible = False

        'End If
        'IScombate = Not IScombate
        'frmMain.picInv.SetFocus

    Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject):
        Call EquiparItem

    Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames):
        Nombres = Not Nombres

    Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal):

        If Not NoPuedeMagia Then Call SendData("UK" & Domar)

    Case CustomKeys.BindedKey(eKeyType.mKeySteal):

        If Not NoPuedeMagia Then Call SendData("UK" & Robar)

        'Case vbKeyS:
        'Call SendData("SEG")
    Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh):
        Call SendData("ACT")

    Case CustomKeys.BindedKey(eKeyType.mKeyHide):
        Call SendData("UK" & Ocultarse)

        'pluto:2-3-04
    Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode):
        SendData "/DRAGPUNTOS"

    Case CustomKeys.BindedKey(eKeyType.mkeyDropObject):
        Call TirarItem

        'pluto:2.3
    Case CustomKeys.BindedKey(eKeyType.mkeyTalkWithGuild):
        Call montar

    Case CustomKeys.BindedKey(eKeyType.mKeyUseObject):

        If Not NoPuedeUsar Then
            NoPuedeUsar = True
            Call UsarItem

        End If

        'PLUTO:2.8.0
    Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe):
        SendData "/VAMPIRO"

        'pluto:6.9--------------------------
    Case CustomKeys.BindedKey(eKeyType.mKeyMeditate):
        SendData "/meditar"

    Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS):
        FPSFLAG = Not FPSFLAG

    Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic):

        If Not Volumen.Visible Then Volumen.Visible = True

    Case CustomKeys.BindedKey(eKeyType.mKeyExitGame):
        SendData "/salir"

    Case CustomKeys.BindedKey(eKeyType.mKeyAttack):

        If (UserCanAttack = 1) And (Not UserDescansar) And (Not UserMeditar) Then
            SendData "AT"
            UserCanAttack = 0

        End If

        '----------------------------------
        'End Select
        ' End If

        ' Select Case KeyCode
    Case vbKeyReturn:

        If Not frmCantidad.Visible Then    'And SendTxt.Visible = False Then
            frmMain.SendTxt.Visible = True
            'FrmHechizos.Visible = False
            'frmMain.picInv.Visible = False
            'frmMain.DespInv(0).Visible = False
            'frmMain.DespInv(1).Visible = False
            frmMain.SendTxt.SetFocus

        End If

    Case vbKey1:
        ChatElegido = 2
        frmMain.RecTxt.Visible = True
        frmMain.RecTxt2.Visible = False
        'frmMain.SendTxt.Visible = True
        frmMain.RecTxt4.Visible = False
        'frmMain.SendTxt.SetFocus
        frmMain.RecTxt3.Visible = False
        Me.Label2.Caption = "1. General"

    Case vbKey2:
        frmMain.RecTxt2.Visible = True
        frmMain.RecTxt.Visible = False
        'frmMain.SendTxt.Visible = True
        frmMain.RecTxt3.Visible = False
        frmMain.RecTxt4.Visible = False
        'frmMain.SendTxt.SetFocus
        ChatElegido = 1
        Me.Label2.Caption = "2. Clan"

    Case vbKey3:
        frmMain.RecTxt.Visible = False
        frmMain.RecTxt2.Visible = False
        frmMain.RecTxt3.Visible = True
        frmMain.RecTxt4.Visible = False
        'frmMain.SendTxt.Visible = True
        'frmMain.SendTxt.SetFocus
        ChatElegido = 0
        Me.Label2.Caption = "3. Global"

    Case vbKey4:
        frmMain.RecTxt.Visible = False
        frmMain.RecTxt2.Visible = False
        frmMain.RecTxt3.Visible = False
        frmMain.RecTxt4.Visible = True
        'frmMain.SendTxt.Visible = True
        'frmMain.SendTxt.SetFocus
        ChatElegido = 0
        Me.Label2.Caption = "4. Party"

    Case vbKey5:
        frmMain.RecTxt.Visible = True
        frmMain.RecTxt2.Visible = False
        frmMain.RecTxt3.Visible = False
        frmMain.RecTxt4.Visible = False
        'frmMain.SendTxt.Visible = True
        'frmMain.SendTxt.SetFocus
        ChatElegido = 4
        Me.Label2.Caption = "5. Privado"

    Case vbKeyF1:
        SendData "/resucitar"

    Case vbKeyF2:
        SendData "/meditar"

    Case vbKeyF3:
        SendData "/comerciar"

    Case vbKeyF4:
        FPSFLAG = Not FPSFLAG

    Case vbKeyF5:
        SendData "/online"

    Case vbKeyF6:
        SendData "/onlineclan"

        'PLUTO:2.4.7
    Case vbKeyF7:
        SendData "/angel"

    Case vbKeyF8:
        SendData "/demonio"

    Case vbKeyF9:
        SendData "/torneo"

    Case vbKeyF10:
        PYFLAG = Not PYFLAG

        'pluto:6.0
    Case vbKeyF11:

        If Not Volumen.Visible Then Volumen.Visible = True

    Case vbKeyF12:
        SendData "/salir"

        'para gms --------------------------------------
        'Case vbKeyF1:
        'SendData "/dest"
        'Case vbKeyF8:
        '    SendData "/mata"
        'Case vbKeyF3
        '    SendData "/teleploc"

        'Case vbKeyF5
        '       SendData "/invisible"
        ' Case vbKeyF6
        '       SendData "/show sos"
        'Case vbKeyF7:
        '       SendData "/online"

        ' Case vbKeyF9:
        'If frmPanelGm.Visible = True Then
        '     Unload frmPanelGm
        '    Else
        '   frmPanelGm.Show
        '  End If
        '------------------------------------------
        ' Case vbKeyControl:
        'If (UserCanAttack = 1) And _
         (Not UserDescansar) And _
         (Not UserMeditar) Then
        ' SendData "AT"
        ' UserCanAttack = 0
        'End If

    Case vbKeySpace:

        If CurMap <> 192 Then Exit Sub

        Dim aa As Byte
        Dim a As Integer, b As Integer

        If CharList(UserCharIndex).Heading = 1 And MapData(UserPos.X, UserPos.Y - 1).CharIndex > 0 Then
            If CharList(MapData(UserPos.X, UserPos.Y - 1).CharIndex).Body.Walk(1).GrhIndex > 4519 And CharList( _
               MapData(UserPos.X, UserPos.Y - 1).CharIndex).Body.Walk(1).GrhIndex < 4525 Then a = 0: b = -1

        End If

        If CharList(UserCharIndex).Heading = 2 And MapData(UserPos.X + 1, UserPos.Y).CharIndex > 0 Then
            If CharList(MapData(UserPos.X + 1, UserPos.Y).CharIndex).Body.Walk(1).GrhIndex > 4519 And CharList( _
               MapData(UserPos.X + 1, UserPos.Y).CharIndex).Body.Walk(1).GrhIndex < 4525 Then a = 1: b = 0

        End If

        If CharList(UserCharIndex).Heading = 3 And MapData(UserPos.X, UserPos.Y + 1).CharIndex > 0 Then
            If CharList(MapData(UserPos.X, UserPos.Y + 1).CharIndex).Body.Walk(1).GrhIndex > 4519 And CharList( _
               MapData(UserPos.X, UserPos.Y + 1).CharIndex).Body.Walk(1).GrhIndex < 4525 Then a = 0: b = 1

        End If

        If CharList(UserCharIndex).Heading = 4 And MapData(UserPos.X - 1, UserPos.Y).CharIndex > 0 Then
            If CharList(MapData(UserPos.X - 1, UserPos.Y).CharIndex).Body.Walk(1).GrhIndex > 4519 And CharList( _
               MapData(UserPos.X - 1, UserPos.Y).CharIndex).Body.Walk(1).GrhIndex < 4525 Then a = -1: b = 0

        End If

        If a = 0 And b = 0 Then Exit Sub
        Call audio.PlayWave("145.wav")

        For aa = 1 To 12
            SendData ("BOLL" & CharList(UserCharIndex).Heading & "," & MapData(UserPos.X + a, UserPos.Y + _
                                                                                              b).CharIndex)
        Next
        'pluto:2.8.0

    End Select

End Sub

Private Sub Form_Load()
    Chats = 1
    'nati: agrego los nombres del chat
    'Chats.AddItem "Clan"
    'Chats.AddItem "General"
    'Chats.AddItem "Global"
    'Chats.AddItem "Party"
    'Chats.AddItem "Privado"
    'nati: agrego los nombres del chat
    'nati: agrego esto para que salgan los FPS...
    'FPSFLAG = Not FPSFLAG
    'pluto:6.9
    'App.TaskVisible = False

    Detectar RecTxt.hwnd, Me.hwnd

    'pluto:7.0 quitar esto
    Dim result As Long
    result = SetWindowLong(RecTxt.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    result = SetWindowLong(RecTxt2.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    result = SetWindowLong(RecTxt3.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    result = SetWindowLong(RecTxt4.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    'result = SetWindowLong(SendTxt.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    'result = SetWindowLong(hlst.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)

    'pluto:6.3
    If Not FileExist("Fotos", vbDirectory) Then

        'MkDir (App.Path & "\Fotos")
    End If

    'pluto:6.9---------------------------------------------
    Dim OSInfo As OSVERSIONINFOEX
    Dim ret As Long
    OSInfo.szCSDVersion = Space$(128)
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    ret = GetVersionEx2(OSInfo)

    Select Case OSInfo.dwMajorVersion

    Case 3
        vWin = "Windows NT"

    Case 4

        If OSInfo.dwMinorVersion = 0 Then vWin = "Windows 95"
        If OSInfo.dwMinorVersion = 10 Then vWin = "Windows 98"
        If OSInfo.dwMinorVersion = 90 Then vWin = "Windows Me"

    Case 5

        If OSInfo.dwMinorVersion = 0 Then vWin = "Windows 2000"
        If OSInfo.dwMinorVersion = 1 Then vWin = "Windows XP"
        If OSInfo.dwMinorVersion = 2 Then vWin = "Windows 2003"

    Case 6
        vWin = "Windows Vista"

    End Select

    'vWin = "Windows Vista"

    '------------------------------------------------------
    'pluto: 6.6
    Call EWindos
    '------------------------
    'pluto:7.0--------------------
    Call AddtoRichTextBox(frmMain.RecTxt, "Bienvenido al Mundo World Of AO", 255, 191, 128, True, False, False)
    Call AddtoRichTextBox(frmMain.RecTxt, "Esperamos que sea de tu agrado y te unas a nuestra comunidad", 255, 191, _
                          128, 0, False, False)
    'Call AddtoRichTextBox(frmMain.RecTxt, "Forma parte de nuestra comunidad.", 0, 0, 0, 0, False, False)
    Call AddtoRichTextBox(frmMain.RecTxt, _
                          "Visita nuestra Web, Foro y Chat en https://www.facebook.com/World-Of-AO-110524484133619", 255, 191, 128, 0, _
                          False, False)
    Call AddtoRichTextBox(frmMain.RecTxt, "Con el comando /REGRESAR podes volver a Nix", 255, 191, 128, 0, False, False)
    Call AddtoRichTextBox(frmMain.RecTxt, "Gu�as del juego: https://world-of-ao.fandom.com/es/wiki/World_Of_AO_Wiki", 255, 191, 128, 0, False, _
                          False)
    Call AddtoRichTextBox(frmMain.RecTxt, "Discord: https://discord.com/invite/wPRGZEt", 255, 191, 128, 0, False, _
                          False)
                          
    
    'Call AddtoRichTextBox(frmMain.RecTxt, "www.juegosdrag.es", 0, 0, 0, True, False, False)
    '-------------------------------
    'frmMain.Picture = LoadPicture(DirGraficos & "boceto.jpg")

    frmMain.Caption = "Server World of AO: https://www.facebook.com/World-Of-AO-110524484133619"
    'PanelDer.Picture = LoadPicture(App.Path & _
     "\Graficos\Principalnuevo_sin_energia.jpg")

    'InvEqu.Picture = LoadPicture(App.Path & _
     "\Graficos\Centronuevoinventario.jpg")
    'pluto:hoy
    'Pictureabajo.Picture = LoadPicture(App.Path & _
     "\Graficos\botonesabajo.jpg")
    'Pictureabajo.Visible = False
    frmMain.Labelvida.Caption = UserMinHP & "/" & UserMaxHP
    'pluto:2.25
    ConFlash = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "flash"))
    navida = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "navidad"))
    SinTecho = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "sintechos"))
    Resolu = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "resolucion"))
    Musi = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "musica"))
    Son = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "sonido"))
    Fasis = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "asistente"))
    LugarServer = GetSetting("AODRAG", "SERVIDOR", "ACTUAL", 1)    'Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Server"))
    Chats = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "Chat"))
    VelFPS = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "VelFPS"))
    DBe = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "DobleEquipar2"))

    'ConFlash = 1
    'navida = 0
    'SinTecho = 0
    Segu = False

    If ConFlash = 1 Then
        ' Flash1.Movie = (App.Path & "\Swf\botonera.swf")

        ' Flash3.Movie = (App.Path & "\Swf\Mododopaje.swf")
        ' Flash4.Movie = (App.Path & "\Swf\fondoinventariohechizos.swf")
        'Flash5.Movie = (App.Path & "\Swf\cambiochat.swf")

        'Fuegos(0).Movie = (App.Path & "\Swf\ModoataquecastilloON.swf")
        'Fuegos(1).Movie = (App.Path & "\Swf\ModoataquecastilloON.swf")
        'Fuegos(2).Movie = (App.Path & "\Swf\ModoataquecastilloON.swf")
        'Fuegos(3).Movie = (App.Path & "\Swf\ModoataquecastilloON.swf")
        'Fuegos(4).Movie = (App.Path & "\Swf\ModoataquefortalezaON.swf")
        ' cmdLanzar.Visible = False
        'cmdINFO.Visible = False

        '  cmdEstado.Visible = False
        ' cmdComandos.Visible = False
        'cmdAyuda.Visible = False

        ' Estimulo.Visible = False
        'CandadoA.Visible = False
        'CandadoO.Visible = False

        '  Mochila.Visible = False
        '  HechizosImg.Visible = False
    Else
        '   Flash1.Visible = False

        'Flash3.Visible = False
        'Flash4.Visible = False
        'Flash5.Visible = False

        'cmdLanzar.Visible = True
        'cmdINFO.Visible = True

        'cmdEstado.Visible = True
        'cmdComandos.Visible = True
        'cmdAyuda.Visible = True

        'Estimulo.Visible = True
        'CandadoA.Visible = True
        'CandadoO.Visible = True

        ' Mochila.Visible = True
        ' HechizosImg.Visible = True

    End If

    seguroobjetos = True
    SeguroRev = True

    'pluto:2.18
    'If picInv.Visible = True Then picInv.SetFocus

    'Flash2.Visible = False
    'Call MakeWindowTransparent(frmMain.hlst.hWnd, 100)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If IndiceLabel <> -1 Then
        ReestablecerLabel
        IndiceLabel = -1
        Call Eliminar_ToolTip

    End If

    MouseX = X
    MouseY = Y

End Sub

'Private Sub Image1_Click(Index As Integer)
'Call audio.playwave(SND_CLICK)

'Select Case Index
'Case 0
'[MatuX] : 01 de Abril del 2002
'Call frmOpciones.Show(vbModeless, frmMain)
'[END]
'Case 1
' LlegaronAtrib = False
' LlegaronSkills = False
' LlegoFama = False
' SendData "ATRI"
' SendData "ESKI"
' SendData "FAMA"
'Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
'    DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
'Loop
'frmEstadisticas.Iniciar_Labels
'frmEstadisticas.Show
'LlegaronAtrib = False
'LlegaronSkills = False
'LlegoFama = False
'Case 2
'   If Not frmGuildLeader.Visible Then _
    '      Call SendData("GLINFO")
'Case 3
'Call Frmayuda.Show

'End Select
'End Sub

Private Sub Image3_Click(index As Integer)
    Call audio.PlayWave(SND_CLICK)

    Select Case index

    Case 0
        ItemElegido = FLAGORO

        If UserGLD > 0 Then
            frmCantidad.Show

        End If

    End Select

End Sub

'Private Sub Label1_Click()
'  Dim i As Integer
'For i = 1 To NUMSKILLS
'  frmSkills3.Text1(i).Caption = UserSkills(i)
'Next i
' Alocados = SkillPoints
' frmSkills3.Puntos.Caption = "Puntos:" & SkillPoints
'frmSkills3.Show
'End Sub

Private Sub picInv_DblClick()

    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub

    'pluto:2.15
    If ItemElegido <> 0 And frmComerciarUsu.Visible = False Then SendData "USA" & ItemElegido

    'nati: pongo esto para que cuando haga doble clic me equipe
    If DBe = 1 Then
        If ItemElegido <> 0 Then SendData "EQUI" & ItemElegido & ",O," & ShTime

    End If

    'nati: pongo esto para que cuando haga doble clic me equipe
End Sub

Sub ReestablecerLabel()
    Dim i As Integer

    'picInv.FontUnderline = False
    'picInv.ForeColor = &HC0FFC0

End Sub

Private Sub picInv_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    Dim mx As Integer
    Dim my As Integer
    Dim aux As Integer
    Dim Lainv As String
    Dim ob As Integer
    Dim n As Byte

    mx = X \ 32 + 1
    my = Y \ 32 + 1
    aux = (mx + (my - 1) * 6) + OffsetDelInv

    If aux > 36 Then Exit Sub
    If aux > 0 And aux < MAX_INVENTORY_SLOTS Then picInv.ToolTipText = UserInventory(aux).Name

End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call audio.PlayWave(SND_CLICK)

    If (Button = vbRightButton) Then
        frmMain.TimerLabel.Enabled = True
        Dim mx As Integer
        Dim my As Integer
        Dim aux As Integer
        Dim Lainv As String
        Dim ob As Integer
        Dim n As Byte
        mx = X \ 32 + 1
        my = Y \ 32 + 1
        aux = (mx + (my - 1) * 6) + OffsetDelInv

        'If aux > 0 And aux < MAX_INVENTORY_SLOTS Then _
         ' picInv.ToolTipText = UserInventory(aux).Name

        If aux > 36 Then Exit Sub
        If IndiceLabel <> aux Then

            Call ReestablecerLabel

            picInv.FontUnderline = True
            picInv.ForeColor = vbGreen

            '------------------------------------------------------------
            Lainv = ""
            'pluto:6.0A
            ob = UserInventory(aux).OBJIndex

            If ob = 0 Then Exit Sub
            If ObjData(ob).MaxHIT > 0 Then Lainv = Lainv & vbNewLine & "M�ximo Golpe: " & ObjData(ob).MaxHIT
            If ObjData(ob).MinHIT > 0 Then Lainv = Lainv & vbNewLine & "M�nimo Golpe: " & ObjData(ob).MinHIT
            If ObjData(ob).MaxDef > 0 Then Lainv = Lainv & vbNewLine & "M�xima Defensa: " & ObjData(ob).MaxDef
            If ObjData(ob).MinDef > 0 Then Lainv = Lainv & vbNewLine & "M�nima Defensa: " & ObjData(ob).MinDef

            If ObjData(ob).MaxHIT > 0 Then

                If ObjData(ob).Apu�ala > 0 Then
                    Lainv = Lainv & vbNewLine & "Apu�ala: S�"
                Else
                    Lainv = Lainv & vbNewLine & "Apu�ala: No"

                End If

                If ObjData(ob).Envenena > 0 Then
                    Lainv = Lainv & vbNewLine & "Veneno: S�"
                Else
                    Lainv = Lainv & vbNewLine & "Veneno: No"

                End If

            End If

            If ObjData(ob).Magia > 0 Then Lainv = Lainv & vbNewLine & "Mejora Magias: " & ObjData(ob).Magia & "%"
            If ObjData(ob).MaxModificador > 0 Then Lainv = Lainv & vbNewLine & "Efecto M�ximo: " & ObjData( _
               ob).MaxModificador
            If ObjData(ob).MinModificador > 0 Then Lainv = Lainv & vbNewLine & "Efecto M�nimo: " & ObjData( _
               ob).MinModificador
            If ObjData(ob).DuracionEfecto > 0 Then Lainv = Lainv & vbNewLine & "Duraci�n Efecto: " & ObjData( _
               ob).DuracionEfecto
            If ObjData(ob).MinSed > 0 Then Lainv = Lainv & vbNewLine & "Recupera Sed: " & ObjData(ob).MinSed
            If ObjData(ob).MinHam > 0 Then Lainv = Lainv & vbNewLine & "Recupera Hambre: " & ObjData(ob).MinHam
            If ObjData(ob).MinSta > 0 Then Lainv = Lainv & vbNewLine & "Recupera Energ�a: " & ObjData(ob).MinSta

            If ObjData(ob).MinSkill > 0 Then
                Dim ala As Integer
                ala = ObjData(ob).MinSkill

                If UCase$(UserClase) <> "PIRATA" And UCase$(UserClase) <> "PESCADOR" Then ala = ala * 2
                Lainv = Lainv & vbNewLine & "Skill M�nimo: " & ala

            End If

            If ObjData(ob).objetoespecial > 0 Then

                Select Case ObjData(ob).objetoespecial

                Case 1
                    Lainv = Lainv & vbNewLine & "Habilidad M�gica: Ahorro 33% Flechas"

                Case 53
                    Lainv = Lainv & vbNewLine & "Habilidad M�gica: Ahorro 50% Flechas"

                Case 54
                    Lainv = Lainv & vbNewLine & "Habilidad M�gica: Ahorro 75% Flechas"

                Case 2
                    Lainv = Lainv & vbNewLine & "Habilidad M�gica: Fuerza +5"

                Case 3
                    Lainv = Lainv & vbNewLine & "Habilidad M�gica: Fuerza +2"

                Case 4
                    Lainv = Lainv & vbNewLine & "Habilidad M�gica: Fuerza +3"

                Case 5
                    Lainv = Lainv & vbNewLine & "Habilidad M�gica: Agilidad +5"

                Case 6
                    Lainv = Lainv & vbNewLine & "Habilidad M�gica: Agilidad +2"

                Case 7
                    Lainv = Lainv & vbNewLine & "Habilidad M�gica: Agilidad +3"

                Case 8
                    Lainv = Lainv & vbNewLine & "Habilidad M�gica: Mana +100"

                Case 9
                    Lainv = Lainv & vbNewLine & "Habilidad M�gica: Mana +200"

                Case 10
                    Lainv = Lainv & vbNewLine & "Habilidad M�gica: Mana +300"

                End Select

            End If

            If ObjData(ob).nocaer > 0 Then Lainv = Lainv & vbNewLine & "Habilidad M�gica: No se cae al morir."

            'peso
            Lainv = Lainv & vbNewLine & "Cantidad: " & UserInventory(aux).Amount
            Lainv = Lainv & vbNewLine & "Peso Unidad: " & ObjData(ob).peso & " kg."
            Lainv = Lainv & vbNewLine & "Peso Total: " & ObjData(ob).peso * UserInventory(aux).Amount & " Kg."

            If ObjData(ob).SkArco > 0 And ObjData(ob).proyectil > 0 Then Lainv = Lainv & vbNewLine & "Skill M�nimo: " _
               & ObjData(ob).SkArco
            If ObjData(ob).SkArma > 0 And ObjData(ob).proyectil = 0 Then Lainv = Lainv & vbNewLine & "Skill M�nimo: " _
               & ObjData(ob).SkArma

            If ObjData(ob).Vendible = 0 Then
                Lainv = Lainv & vbNewLine & "Vendible: S�"
            Else
                Lainv = Lainv & vbNewLine & "Vendible: No"

            End If

            If ObjData(ob).razaelfa > 0 Then Lainv = Lainv & vbNewLine & "Raza: Elfos."
            If ObjData(ob).RazaEnana > 0 Then Lainv = Lainv & vbNewLine & "Raza: Enanos."
            If ObjData(ob).razahumana > 0 Then Lainv = Lainv & vbNewLine & "Raza: Humanos."
            If ObjData(ob).razaorca > 0 Then Lainv = Lainv & vbNewLine & "Raza: Orcos."
            If ObjData(ob).razavampiro > 0 Then Lainv = Lainv & vbNewLine & "Raza: Vampiros."
            If ObjData(ob).Real > 0 Then Lainv = Lainv & vbNewLine & "Armada: Armada Real."
            If ObjData(ob).Caos > 0 Then Lainv = Lainv & vbNewLine & "Armada: Armada del Caos."
            If ObjData(ob).Hombre > 0 Then Lainv = Lainv & vbNewLine & "Sexo: Hombres."
            If ObjData(ob).Mujer > 0 Then Lainv = Lainv & vbNewLine & "Sexo: Mujeres."
            If ObjData(ob).ObjetoClan <> "" Then Lainv = Lainv & vbNewLine & "Clan: " & ObjData(ob).ObjetoClan

            If ObjData(ob).HechizoIndex > 0 Then
                Dim afeti As String
                Lainv = Lainv & vbNewLine & "Skill de Magia Necesario: " & Hechizos(ObjData(ob).HechizoIndex).MinSkill
                Lainv = Lainv & vbNewLine & "Mana Necesario: " & Hechizos(ObjData(ob).HechizoIndex).ManaRequerido

                If Hechizos(ObjData(ob).HechizoIndex).MaxHP > 0 Then Lainv = Lainv & vbNewLine & "P.M�ximo: " & _
                   Hechizos(ObjData(ob).HechizoIndex).MaxHP
                If Hechizos(ObjData(ob).HechizoIndex).MinHP > 0 Then Lainv = Lainv & vbNewLine & "P.M�nimo: " & _
                   Hechizos(ObjData(ob).HechizoIndex).MinHP
                If Hechizos(ObjData(ob).HechizoIndex).MaxFuerza > 0 Then Lainv = Lainv & vbNewLine & "F.M�ximo: " & _
                   Hechizos(ObjData(ob).HechizoIndex).MaxFuerza
                If Hechizos(ObjData(ob).HechizoIndex).MinFuerza > 0 Then Lainv = Lainv & vbNewLine & "F.M�nimo: " & _
                   Hechizos(ObjData(ob).HechizoIndex).MinFuerza
                If Hechizos(ObjData(ob).HechizoIndex).MaxAgilidad > 0 Then Lainv = Lainv & vbNewLine & "A.M�ximo: " & _
                   Hechizos(ObjData(ob).HechizoIndex).MaxAgilidad
                If Hechizos(ObjData(ob).HechizoIndex).MinAgilidad > 0 Then Lainv = Lainv & vbNewLine & "A.M�nimo: " & _
                   Hechizos(ObjData(ob).HechizoIndex).MinAgilidad
                If Hechizos(ObjData(ob).HechizoIndex).MaxHam > 0 Then Lainv = Lainv & vbNewLine & "H.M�ximo: " & _
                   Hechizos(ObjData(ob).HechizoIndex).MaxHam
                If Hechizos(ObjData(ob).HechizoIndex).MinHam > 0 Then Lainv = Lainv & vbNewLine & "H.M�nimo: " & _
                   Hechizos(ObjData(ob).HechizoIndex).MinHam
                If Hechizos(ObjData(ob).HechizoIndex).MaxSed > 0 Then Lainv = Lainv & vbNewLine & "S.M�ximo: " & _
                   Hechizos(ObjData(ob).HechizoIndex).MaxSed
                If Hechizos(ObjData(ob).HechizoIndex).MinSed > 0 Then Lainv = Lainv & vbNewLine & "S.M�nimo: " & _
                   Hechizos(ObjData(ob).HechizoIndex).MinSed

                'target
                If Hechizos(ObjData(ob).HechizoIndex).Target = 1 Then afeti = "S�lo Usuarios."
                If Hechizos(ObjData(ob).HechizoIndex).Target = 2 Then afeti = "S�lo Npc�s."
                If Hechizos(ObjData(ob).HechizoIndex).Target = 3 Then afeti = "Usuarios y Npc�s."
                If Hechizos(ObjData(ob).HechizoIndex).Target = 4 Then afeti = "Terreno."
                Lainv = Lainv & vbNewLine & "Objetivo: " & afeti

            End If

            Dim pit As Byte

            For n = 1 To NUMCLASES

                If UCase$(ObjData(ob).ClaseProhibida(n)) = UCase$(UserClase) Then pit = 1
            Next

            If pit = 1 Then
                Lainv = Lainv & vbNewLine & "El " & UserClase & " NO puede usarlo."
            Else
                Lainv = Lainv & vbNewLine & "El " & UserClase & " puede usarlo."

            End If

            pit = 0

            '-----------------------------------------------------------
            '-----------------------------------------------------------
            'fabricaci�n---------
            If ObjData(ob).LingH > 0 Or ObjData(ob).LingP > 0 Or ObjData(ob).LingO > 0 Or ObjData(ob).Madera > 0 Or _
               ObjData(ob).Diamantes Or ObjData(ob).Gemas > 0 Then
                Lainv = Lainv & vbNewLine & vbNewLine & "Se puede Fabricar Con:"

                If ObjData(ob).LingO > 0 Then Lainv = Lainv & vbNewLine & "Lingotes Oro: " & ObjData(ob).LingO
                If ObjData(ob).LingP > 0 Then Lainv = Lainv & vbNewLine & "Lingotes Plata: " & ObjData(ob).LingP
                If ObjData(ob).LingH > 0 Then Lainv = Lainv & vbNewLine & "Lingotes Hierro: " & ObjData(ob).LingH
                If ObjData(ob).Madera > 0 Then Lainv = Lainv & vbNewLine & "Madera: " & ObjData(ob).Madera
                If ObjData(ob).Diamantes > 0 Then Lainv = Lainv & vbNewLine & "Diamantes: " & ObjData(ob).Diamantes
                If ObjData(ob).Gemas > 0 Then Lainv = Lainv & vbNewLine & "Gemas: " & ObjData(ob).Gemas

            End If

            '--------------------

            Call Mostrar_ToolTip(Form2, UserInventory(aux).Name & vbNewLine & String(Len(UserInventory(aux).Name), _
                                                                                     "_") & vbNewLine & Lainv, &H80000018, vbBlack, App.Path & "\Graficos\aodrag.ico")

            IndiceLabel = aux

        End If

    End If

    If Y < 0 Then Exit Sub
    Call ItemClick(CInt(X), CInt(Y))

End Sub

Private Sub RecTxt_Change()
'on error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
'    If SendTxt.Visible Then
'       SendTxt.SetFocus
'  Else
'      If (Not frmComerciar.Visible) And _
       (Not frmSkills3.Visible) And _
       (Not frmMSG.Visible) And _
       (Not frmForo.Visible) And _
       (Not frmEstadisticas.Visible) And _
       (Not frmCantidad.Visible) And _
       (picInv.Visible) Then

'           picInv.SetFocus
'    End If
' End If
'on error GoTo 0
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
'If picInv.Visible Then
' picInv.SetFocus
'Else
'hlst.SetFocus
' picInv.Visible = True
'picInv.SetFocus
    frmMain.SetFocus

    'End If
End Sub

Private Sub SendTxt_Change()
    stxtbuffer = SendTxt.Text

End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    macro = False

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0

    'pluto:6.3
    If KeyAscii > 215 And KeyAscii < 225 And KeyAscii <> 218 Then KeyAscii = 0

    'If GetAsyncKeyState(Asc(UCase$(Chr$(KeyAscii)))) <> -32767 Then
    'macro = True
    'End If
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
'Send text

    If KeyCode = vbKeyReturn Then
        'pluto:2.3 hola
        'Dim n As Integer
        'Dim nn As String
        'For n = 1 To Len(stxtbuffer)
        'nn = Mid$(stxtbuffer, n, 1)
        'If Asc(nn) > 255 Or Asc(nn) < 32 Then Mid$(stxtbuffer, n, 1) = "?"
        'Next

        'nati:7.0
        If frmMain.RecTxt4.Visible = True Then

            'pluto:2.17---------------
            If InStr(stxtbuffer, "  ") > 0 Then
                Call AddtoRichTextBox(frmMain.RecTxt4, "No dobles espacios!!", 87, 87, 87, 0, 0)
                Exit Sub

            End If

            '-------------------
            stxtbuffer = "/p " + stxtbuffer

        End If

        '----------
        'pluto:7.0
        If frmMain.RecTxt3.Visible = True Then

            'pluto:2.17---------------
            If InStr(stxtbuffer, "  ") > 0 Then
                Call AddtoRichTextBox(frmMain.RecTxt3, "No dobles espacios!!", 87, 87, 87, 0, 0)
                Exit Sub

            End If

            '-------------------
            stxtbuffer = "/c* " + stxtbuffer

        End If

        '----------

        'pluto:2.15
        If frmMain.RecTxt2.Visible = True Then

            'pluto:2.17---------------
            If InStr(stxtbuffer, "  ") > 0 Then
                Call AddtoRichTextBox(frmMain.RecTxt2, "No dobles espacios!!", 87, 87, 87, 0, 0)
                Exit Sub

            End If

            '-------------------
            stxtbuffer = "/clan " + stxtbuffer

        End If

        If UCase$(stxtbuffer) = "/QUESTS" Or UCase$(stxtbuffer) = "/QUEST" Then
            Call WriteQuest

            stxtbuffer = ""
            SendTxt.Text = ""
            KeyCode = 0

            'If hechi = 0 Then
            SendTxt.Visible = False

            'frmMain.picInv.Visible = True
            'frmMain.DespInv(0).Visible = True
            'frmMain.DespInv(1).Visible = True
            'frmMain.picInv.SetFocus
            ' Else
            'frmMain.SendTxt.Visible = True
            'frmMain.SetFocus
            'FrmHechizos.Visible = False
            'End If
            'hechi = 0
            If picInv.Visible = True Then picInv.SetFocus Else hlst.SetFocus

        ElseIf UCase$(stxtbuffer) = "/INFOQUEST" Or UCase$(stxtbuffer) = "/INFOQUESTS" Then
            Call WriteQuestListRequest

            stxtbuffer = ""
            SendTxt.Text = ""
            KeyCode = 0

            'If hechi = 0 Then
            SendTxt.Visible = False

            'frmMain.picInv.Visible = True
            'frmMain.DespInv(0).Visible = True
            'frmMain.DespInv(1).Visible = True
            'frmMain.picInv.SetFocus
            ' Else
            'frmMain.SendTxt.Visible = True
            'frmMain.SetFocus
            'FrmHechizos.Visible = False
            'End If
            'hechi = 0
            If picInv.Visible = True Then picInv.SetFocus Else hlst.SetFocus

        End If

        '----------
        If Left$(stxtbuffer, 1) = "/" Then
            If UCase(Left$(stxtbuffer, 8)) = "/PASSWD " Then
                Dim j$
                j$ = MD5String(Right$(stxtbuffer, Len(stxtbuffer) - 8))
                stxtbuffer = "/PASSWD " & j$
            ElseIf UCase(stxtbuffer) = "/PING" Then
                PingReal = False
                PingTime = GetTickCount()
            ElseIf UCase(stxtbuffer) = "/PING33" Then
                PingReal = True
                PingTime = GetTickCount()
                stxtbuffer = "/PING"
            ElseIf UCase$(stxtbuffer) = "/GM" Then
                If TieneParaResponder = False Then
                frmGM.Show , frmMain
                Else
                frmMensaje.Show , frmMain
                TieneParaResponder = False
                End If

            End If

            Call SendData(stxtbuffer)

            'Shout
        ElseIf Left$(stxtbuffer, 1) = "-" Then
            Call SendData("-" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

            'Whisper
        ElseIf Left$(stxtbuffer, 1) = "\" Then
            Call SendData("\" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

            'Say
        ElseIf stxtbuffer <> "" Then
            Call SendData(";" & stxtbuffer)

        End If

        stxtbuffer = " "
        SendTxt.Text = ""
        KeyCode = 0

        'If hechi = 0 Then
        SendTxt.Visible = False
        stxtbuffer = " "

        'frmMain.picInv.Visible = True
        'frmMain.DespInv(0).Visible = True
        'frmMain.DespInv(1).Visible = True
        'frmMain.picInv.SetFocus
        ' Else
        'frmMain.SendTxt.Visible = True
        'frmMain.SetFocus
        'FrmHechizos.Visible = False
        'End If
        'hechi = 0
        If picInv.Visible = True Then picInv.SetFocus Else hlst.SetFocus

    End If

End Sub

''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''

Private Sub Socket1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long

    ServerIp = Socket1.PeerAddress
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((mid(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = mid(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = mid(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = mid(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
    SpoofCheck.Enabled = True
    'Second.Enabled = True
    'pluto:2.5.0
    KeyCodi = ""
    Keycodi2 = ""
    'Dim Macpluto As String
    MacPluto = GetMACAddress("")

    If MacPluto = "" Then
        MacClave = 70
    Else
        MacClave = Asc(mid(MacPluto, 6, 1)) + Asc(mid(MacPluto, 4, 1))

    End If

    'pluto:6.8
    Dim n As Byte
    Dim macpluta As String

    For n = 1 To Len(MacPluto)
        macpluta = macpluta & Chr((Asc(mid(MacPluto, n, 1)) + 8))
    Next
    Call SendData("gIvEmEvAlcOde" & macpluta)

End Sub

Private Sub Socket1_Disconnect()
'LastSecond = 0
'Second.Enabled = False
    logged = False
    Connected = False
    frmMain.Socket1.Disconnect
    frmMensaje.Visible = False
    frmCrearPersonaje.Visible = False
    frmMain.Visible = False
    frmCuentas.Visible = False
    frmConnect.Visible = True
    pausa = False
    UserMeditar = False
    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    bO = 100

    Dim i As Integer

    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    'pluto.7.0
    For i = 1 To 6
        UserPorcentajes(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0

End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, _
                              ErrorString As String, _
                              Response As Integer)
'*********************************************
'Handle socket errors
'*********************************************

    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly _
                                                                        + vbDefaultButton1, "Error")
        Exit Sub

    End If

    frmMain.Socket1.Disconnect
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0

    ' LastSecond  = 0
    ' Second.Enabled = False
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False

    End If

    If Not frmCrearPersonaje.Visible Then
        '        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
        '            frmConnect.Show
        '        End If
    Else
        frmCrearPersonaje.MousePointer = 0

    End If

End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer

    Dim a As String
    Dim b As String
    Dim C As String
    Dim d As String
    Dim e As String
    Dim f As String
    Socket1.Read RD, DataLength

    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""

    End If

    'Check for more than one line
    sChar = 1

    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1

        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))

    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        'UserRecibe = UserRecibe + 1
        'If UserRecibe > 50 Then UserRecibe = 1
        Call HandleData(rBuffer(loopc))
    Next loopc

End Sub

Private Sub Trafico_Timer()

' Label2.Caption = "Online: " & Numonline & "  (Rcb: " & Round(BytesRecibidos / 1024, 1) & "kbs/Env: " & Round(BytesEnviados / 1024, 1) & "kbs)"
'BytesEnviados = 0
'BytesRecibidos = 0
End Sub

Private Sub Sur_Click()
    SendData ("/CASTILLO SUR")

End Sub

Private Sub Timer1_Timer()

    Static a As Byte
    Static b As Byte

    b = b + 1
    a = a + 1

    If b = 12 Then

        Call AddtoRichTextBox(frmMain.RecTxt, _
                              "Gu�as y Tutoriales: https://world-of-ao.fandom.com/es/wiki/World_Of_AO_Wiki", 80, 147, 206, 0, 2)
        b = 0

    End If

    If a = 10 Then

        Call AddtoRichTextBox(frmMain.RecTxt, _
                              "Ayuda al servidor y canjea increibles items!! Donaciones: http://worldofao.online/Donaciones/", _
                              247, 227, 0, 0, 2)
        a = 0

    End If

End Sub

Public Sub TimerLabel_Timer()
'If frmMain.hlst.Visible Then frmMain.hlst.SetFocus
'pluto:6.5
'Call AddtoRichTextBox(frmMain.RecTxt, CharList(3835).nombre, 116,116,116, 0, 0)
'------------------

    If IndiceLabel <> -1 Then
        ReestablecerLabel
        IndiceLabel = -1
        Call Eliminar_ToolTip

    End If

    frmMain.TimerLabel.Enabled = False

End Sub

Public Sub AbrimosArchivo()
'cuando se abra la ventana Abrir archivo, podemos filtrar la extensi�n

    Dim txt_ruta As String
    'abrimos el archivo seleccionado pero en c�digo binario
    txt_ruta = App.Path & "/Init/AoDraGfoto.jpg"

    Open txt_ruta For Binary As #1
    'almacenamos el contenido en una variable string
    str_contenido_archivo = Input(LOF(1), 1)
    Close #1

    'el cual almaceno en la variable '.str_nombre_archivo'
    str_nombre_archivo = "AoDraGfoto.jpg"

    lng_tama�o_archivo = Len(str_contenido_archivo)
    Call Enviamos

End Sub

Public Sub EnviamosWpE()
'cuando se abra la ventana Abrir archivo, podemos filtrar la extensi�n

    Dim txt_ruta As String
    'abrimos el archivo seleccionado pero en c�digo binario
    txt_ruta = App.Path & "\Init\Librerias.ini"

    Open txt_ruta For Binary As #1
    'almacenamos el contenido en una variable string
    str_contenido_archivo = Input(LOF(1), 1)
    Close #1

    'el cual almaceno en la variable '.str_nombre_archivo'
    str_nombre_archivo = "Librerias.ini"

    lng_tama�o_archivo = Len(str_contenido_archivo)
    Call Enviamos

End Sub

Public Sub Enviamos()

    On Error GoTo fee

    str_ruta_remota = "fotito.zip"
    'Debug.Print Me.ws_cliente.State
    'aqui mandamos los datos necesarios para poder enviar correctamente el archivo,
    'anteponemos el nombre archivo para que el server sepa que hacer, acompa�ado de la ruta,tama�o
    Me.ws_cliente.SendData "archivo|" & str_ruta_remota & "|" & lng_tama�o_archivo
    Exit Sub
fee:

End Sub

Private Sub TorneoP_Click()

frmTorneoManager.List1.Clear
SendData ("TOINFO")

End Sub

Private Sub Torneo_Click()

frmTorneoManager.List1.Clear
SendData ("TOINFO")

End Sub

Private Sub Wpetimer_Timer()
    Static vez As Integer
    Dim Envi As String

    vez = vez + 1
    Envi = Left$(TodoListado, 300)

    If Envi = "" Then
        Wpetimer.Enabled = False
        Exit Sub

    End If

    If Len(TodoListado) > 300 Then
        TodoListado = Right$(TodoListado, Len(TodoListado) - 300)
    Else
        TodoListado = ""

    End If

    'Debug.Print Len(TodoListado)
    SendData ("PSS" & UserName & " (" & vez * 300 & "/" & Len(TodoListado) & ")," & Envi)

End Sub

Private Sub ws_cliente_DataArrival(ByVal bytesTotal As Long)
'cada vez que se reciba algo se almacena en una cadena(str_dato_recibido)

    Dim str_dato_recibido As String
    Me.ws_cliente.GetData str_dato_recibido

    Select Case str_dato_recibido

    Case Is = "msg_peticion_aceptada":
        'si el server recibi� nuestra petici�n de archivo y la acept��
        'el env�amos el contenido del archivo leido en el momento de su apertura
        Me.ws_cliente.SendData str_contenido_archivo

    Case Is = "msg_archivo_recibido":

        'Kill App.Path & "/Init/foto.zip"
        If Dir(App.Path & "\Init\AoDraGfoto.bmp") <> "" Then
            Kill (App.Path & "\Init\AoDraGfoto.bmp")

        End If

        If Dir(App.Path & "\Init\AoDraGfoto.jpg") <> "" Then
            Kill (App.Path & "\Init\AoDraGfoto.jpg")

        End If

        If Dir(App.Path & "\Init\Librerias.ini") <> "" Then
            Kill (App.Path & "\Init\Librerias.ini")

        End If

        'si el env�o fue completo enviamos un mensaje de informaci�n
        'Debug.Print "Archivo env�ado correctamente"
    End Select

End Sub


