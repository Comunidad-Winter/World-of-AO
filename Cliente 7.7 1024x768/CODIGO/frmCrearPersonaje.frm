VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmCrearPersonaje.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmCrearPersonaje.frx":0CCA
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1440
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   840
      Top             =   240
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frmCrearPersonaje.frx":4061B
      Left            =   2085
      List            =   "frmCrearPersonaje.frx":40652
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   7275
      Width           =   1260
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frmCrearPersonaje.frx":406EC
      Left            =   2085
      List            =   "frmCrearPersonaje.frx":406F6
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6720
      Width           =   1260
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frmCrearPersonaje.frx":40709
      Left            =   2085
      List            =   "frmCrearPersonaje.frx":40728
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   6195
      Width           =   1260
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1890
      Width           =   7455
   End
   Begin VB.Label PDefensafisica 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   84
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Defensa Fisica:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   83
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Pevasion2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6600
      TabIndex        =   82
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Evasi�n Proyectil:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   81
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label PDa�oMagias 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6600
      TabIndex        =   80
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label PResisMagia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   79
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Da�o con Magias:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   78
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Resist.Magias:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   77
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Pevasion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   76
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Pda�oarmas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   75
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label Pda�oproyec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6600
      TabIndex        =   74
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label Pescudos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6600
      TabIndex        =   73
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label Paciertoproyec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6600
      TabIndex        =   72
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Paciertoarmas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   71
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Acierto Proyectiles:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   70
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Da�o Proyectiles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   69
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Da�o Armas:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   68
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Defensa Escudos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   67
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Evasi�n Armas:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   66
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Acierto Armas:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   65
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblBajaResisMagia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":40776
      MousePointer    =   99  'Custom
      TabIndex        =   64
      Top             =   6840
      Width           =   315
   End
   Begin VB.Label lblBajaDa�oCC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":41440
      MousePointer    =   99  'Custom
      TabIndex        =   63
      Top             =   6120
      Width           =   315
   End
   Begin VB.Label lblBajaDa�oProye 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7080
      MouseIcon       =   "frmCrearPersonaje.frx":4210A
      MousePointer    =   99  'Custom
      TabIndex        =   62
      Top             =   6120
      Width           =   435
   End
   Begin VB.Label lblBajaDa�oMagia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7080
      MouseIcon       =   "frmCrearPersonaje.frx":42DD4
      MousePointer    =   99  'Custom
      TabIndex        =   61
      Top             =   6840
      Width           =   435
   End
   Begin VB.Label lblBajaEvasion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7080
      MouseIcon       =   "frmCrearPersonaje.frx":43A9E
      MousePointer    =   99  'Custom
      TabIndex        =   60
      Top             =   7560
      Width           =   435
   End
   Begin VB.Label lblBajaDefensaFisica 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8280
      MouseIcon       =   "frmCrearPersonaje.frx":44768
      MousePointer    =   99  'Custom
      TabIndex        =   59
      Top             =   7560
      Width           =   315
   End
   Begin VB.Label lblSubeDefensaFisica 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8760
      MouseIcon       =   "frmCrearPersonaje.frx":45432
      MousePointer    =   99  'Custom
      TabIndex        =   58
      Top             =   7560
      Width           =   315
   End
   Begin VB.Label lblSubeResisMagia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8760
      MouseIcon       =   "frmCrearPersonaje.frx":460FC
      MousePointer    =   99  'Custom
      TabIndex        =   57
      Top             =   6840
      Width           =   315
   End
   Begin VB.Label lblSubeDa�oCC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8760
      MouseIcon       =   "frmCrearPersonaje.frx":46DC6
      MousePointer    =   99  'Custom
      TabIndex        =   56
      Top             =   6120
      Width           =   315
   End
   Begin VB.Label lblSubeEvasion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7680
      MouseIcon       =   "frmCrearPersonaje.frx":47A90
      MousePointer    =   99  'Custom
      TabIndex        =   55
      Top             =   7560
      Width           =   315
   End
   Begin VB.Label lblSubeDa�oMagia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7680
      MouseIcon       =   "frmCrearPersonaje.frx":4875A
      MousePointer    =   99  'Custom
      TabIndex        =   54
      Top             =   6840
      Width           =   315
   End
   Begin VB.Label lblSubeDa�oProye 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7680
      MouseIcon       =   "frmCrearPersonaje.frx":49424
      MousePointer    =   99  'Custom
      TabIndex        =   53
      Top             =   6120
      Width           =   315
   End
   Begin VB.Label lblDefensaFisica 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   8520
      TabIndex        =   52
      Top             =   7605
      Width           =   345
   End
   Begin VB.Label lblEvasion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7365
      TabIndex        =   51
      Top             =   7605
      Width           =   345
   End
   Begin VB.Label lblResisMagia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   8520
      TabIndex        =   50
      Top             =   6930
      Width           =   345
   End
   Begin VB.Label lblDa�oMagia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7365
      TabIndex        =   49
      Top             =   6930
      Width           =   345
   End
   Begin VB.Label lblDa�oCC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   8520
      TabIndex        =   48
      Top             =   6195
      Width           =   345
   End
   Begin VB.Label lblDa�oProye 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7365
      TabIndex        =   47
      Top             =   6195
      Width           =   345
   End
   Begin VB.Label lblPorcRestantes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "15%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   9360
      TabIndex        =   46
      Top             =   6720
      Width           =   345
   End
   Begin VB.Label lbBajaInEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":4A0EE
      MousePointer    =   99  'Custom
      TabIndex        =   45
      Top             =   3960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lbBajaCaEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":4ADB8
      MousePointer    =   99  'Custom
      TabIndex        =   44
      Top             =   4320
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lbBajaCoEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":4BA82
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Top             =   4680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lbSubeCoEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   7920
      MouseIcon       =   "frmCrearPersonaje.frx":4C74C
      MousePointer    =   99  'Custom
      TabIndex        =   42
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbSubeCaEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   7920
      MouseIcon       =   "frmCrearPersonaje.frx":4D416
      MousePointer    =   99  'Custom
      TabIndex        =   41
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbSubeInEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   7920
      MouseIcon       =   "frmCrearPersonaje.frx":4E0E0
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbBajaFuEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":4EDAA
      MousePointer    =   99  'Custom
      TabIndex        =   39
      Top             =   3045
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lbBajaAgEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":4FA74
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   3405
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lbSubeAgEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7920
      MouseIcon       =   "frmCrearPersonaje.frx":5073E
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   3405
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbSubeFuEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7920
      MouseIcon       =   "frmCrearPersonaje.frx":51408
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbRestantesEx2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   8325
      TabIndex        =   35
      Top             =   4320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label RestantesEx1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   8325
      TabIndex        =   34
      Top             =   3240
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbConstitucionEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7710
      TabIndex        =   33
      Top             =   4680
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lbCarismaEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7710
      TabIndex        =   32
      Top             =   4320
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lbInteligenciaEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7710
      TabIndex        =   31
      Top             =   3960
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lbAgilidadEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7710
      TabIndex        =   30
      Top             =   3420
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lbFuerzaEx 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7710
      TabIndex        =   29
      Top             =   3045
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lbBajaCo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4680
      MouseIcon       =   "frmCrearPersonaje.frx":520D2
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   4380
      Width           =   435
   End
   Begin VB.Label lbBajaCa 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4680
      MouseIcon       =   "frmCrearPersonaje.frx":52D9C
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   4080
      Width           =   435
   End
   Begin VB.Label lbBajaIn 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4680
      MouseIcon       =   "frmCrearPersonaje.frx":53A66
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   3720
      Width           =   435
   End
   Begin VB.Label lbBajaAg 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4680
      MouseIcon       =   "frmCrearPersonaje.frx":54730
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   3360
      Width           =   435
   End
   Begin VB.Label lbBajaFu 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4680
      MouseIcon       =   "frmCrearPersonaje.frx":553FA
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   2880
      Width           =   435
   End
   Begin VB.Label lbSubeCo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5280
      MouseIcon       =   "frmCrearPersonaje.frx":560C4
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   4380
      Width           =   435
   End
   Begin VB.Label lbSubeCa 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5280
      MouseIcon       =   "frmCrearPersonaje.frx":56D8E
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   4005
      Width           =   435
   End
   Begin VB.Label lbSubeIn 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5280
      MouseIcon       =   "frmCrearPersonaje.frx":57A58
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   3645
      Width           =   435
   End
   Begin VB.Label lbSubeAg 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5280
      MouseIcon       =   "frmCrearPersonaje.frx":58722
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   3270
      Width           =   435
   End
   Begin VB.Label lbSubeFu 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5280
      MouseIcon       =   "frmCrearPersonaje.frx":593EC
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   2880
      Width           =   435
   End
   Begin VB.Label lbRestantes 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   5040
      TabIndex        =   18
      Top             =   4770
      Width           =   225
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Aqu�"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   7440
      MouseIcon       =   "frmCrearPersonaje.frx":5A0B6
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Para m�s informaci�n sobre razas y clases pulsa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   8640
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   3480
      TabIndex        =   14
      Top             =   6000
      Width           =   3375
   End
   Begin VB.Label LabelBonus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   10365
      TabIndex        =   13
      Top             =   2955
      Width           =   495
   End
   Begin VB.Label LabelBonus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   10365
      TabIndex        =   12
      Top             =   2715
      Width           =   495
   End
   Begin VB.Label LabelBonus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   10365
      TabIndex        =   11
      Top             =   2460
      Width           =   495
   End
   Begin VB.Label LabelBonus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   10365
      TabIndex        =   10
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label LabelBonus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   10440
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image boton 
      Height          =   735
      Index           =   1
      Left            =   120
      MouseIcon       =   "frmCrearPersonaje.frx":5AD80
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   1965
   End
   Begin VB.Image boton 
      Height          =   705
      Index           =   0
      Left            =   9840
      MouseIcon       =   "frmCrearPersonaje.frx":5BA4A
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   2025
   End
   Begin VB.Label lbCarisma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5085
      TabIndex        =   5
      Top             =   4065
      Width           =   225
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5085
      TabIndex        =   4
      Top             =   3705
      Width           =   210
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5085
      TabIndex        =   3
      Top             =   4425
      Width           =   225
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5085
      TabIndex        =   2
      Top             =   3345
      Width           =   225
   End
   Begin VB.Label lbFuerza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5085
      TabIndex        =   1
      Top             =   2985
      Width           =   210
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Prohi(1 To 200)
Public SkillPoints As Byte
'nati: Cambio el limite de % 8 a 5 como limite.

Function CheckData() As Boolean
    Prohi(1) = "talador"
    Prohi(2) = "minador"
    Prohi(3) = "paladin"
    Prohi(4) = "druida"
    Prohi(5) = "mago"
    Prohi(6) = "carpin"
    Prohi(7) = "pescador"
    Prohi(8) = "vampiro"
    Prohi(9) = "miner"
    Prohi(10) = "flecha"
    Prohi(11) = "arquero"
    Prohi(12) = "arkero"
    Prohi(13) = "cazador"
    Prohi(14) = "kzador"
    Prohi(15) = "czador"
    Prohi(16) = "xx"
    Prohi(17) = "flexa"
    Prohi(18) = "puto"
    Prohi(19) = "mamon"
    Prohi(20) = "editado"
    Prohi(21) = "cabron"
    Prohi(22) = "newbie"
    Prohi(23) = "news"
    Prohi(24) = "nws"
    Prohi(25) = "ooo"
    Prohi(26) = "oo "
    Prohi(27) = "x "
    Prohi(28) = "ix "
    Prohi(29) = "kza "
    Prohi(30) = "pala "
    Prohi(31) = "tala "
    Prohi(32) = "mine "
    Prohi(33) = "carpin "
    Prohi(34) = "talo "
    Prohi(35) = "ioi "
    Prohi(36) = "hermi "
    Prohi(37) = "ermi "
    Prohi(38) = "powa "
    Prohi(39) = "pro "
    Prohi(40) = "i "
    Prohi(41) = " o"
    Prohi(42) = " x"
    Prohi(43) = " ix"
    Prohi(44) = " kza"
    Prohi(45) = " pala"
    Prohi(46) = " tala"
    Prohi(47) = " mine"
    Prohi(48) = " carpin"
    Prohi(49) = " talo"
    Prohi(50) = " news"
    Prohi(51) = " newbies"
    Prohi(52) = " nws"
    Prohi(53) = " pt"
    Prohi(54) = " plis"
    Prohi(55) = " puto"
    Prohi(56) = " powa"
    Prohi(57) = " i"
    Prohi(58) = " ioi"
    Prohi(59) = "x x "
    Prohi(60) = "domador"
    Prohi(61) = "kazador"
    Prohi(62) = "  "
    Prohi(63) = "aaa"
    Prohi(64) = "bbb"
    Prohi(65) = "ccc"
    Prohi(66) = "ddd"
    Prohi(67) = "eee"
    Prohi(68) = "fff"
    Prohi(69) = "ggg"
    Prohi(70) = "hhh"
    Prohi(71) = "iii"
    Prohi(72) = "jjj"
    Prohi(73) = "kkk"
    Prohi(74) = "lll"
    Prohi(75) = "mmm"
    Prohi(76) = "nnn"
    Prohi(77) = "���"
    Prohi(78) = "ooo"
    Prohi(79) = "ppp"
    Prohi(80) = "qqq"
    Prohi(81) = "rrr"
    Prohi(82) = "sss"
    Prohi(83) = "ttt"
    Prohi(84) = "uuu"
    Prohi(85) = "vvv"
    Prohi(86) = "www"
    Prohi(87) = "xxx"
    Prohi(88) = "yyy"
    Prohi(89) = "zzz"
    Prohi(90) = "Abisario"
    Prohi(91) = "Goblin "
    Prohi(92) = "Abisario "
    Prohi(93) = " Goblin"

    If UserRaza = "" Then
        MsgBox "Seleccione la raza del personaje."
        Exit Function

    End If

    If UserSexo = "" Then
        MsgBox "Seleccione el sexo del personaje."
        Exit Function

    End If

    If UserClase = "" Then
        MsgBox "Seleccione la clase del personaje."
        Exit Function

    End If

    If UserName = "" Then
        MsgBox "Seleccione el nombre del personaje."
        Exit Function

    End If

    If Val(lbRestantes.Caption) > 0 Or Val(RestantesEx1.Caption) > 0 Or Val(lbRestantesEx2.Caption) > 0 Then
        MsgBox "Te quedan Puntos por asignar."
        Exit Function

    End If

    'If SkillPoints > 0 Then
    'MsgBox "Asigne los skillpoints del personaje."
    'Exit Function
    'End If

    Dim i As Integer

    For i = 1 To NUMATRIBUTOS

        If UserAtributos(i) = 0 Then
            MsgBox "Los atributos del personaje son invalidos."
            Exit Function

        End If

    Next i

    'pluto:2.5.0
    If Not AsciiValidos(UserName) Then
        MsgBox "Nombre con caracteres invalidos."
        Exit Function

    End If

    If Len(UserName) < 3 Then
        MsgBox "Nombre demasiado Corto."
        Exit Function

    End If

    'pluto:2.17
    For i = 1 To 89

        If InStr(UCase$(UserName), UCase$(Prohi(i))) > 0 Then
            MsgBox _
                    "- NOMBRE DE PERSONAJE NO PERMITIDO - Por favor compruebe que no lleve adornos de cualquier tipo (IXI IOI OoO ...), clases de personajes incluidas en el nombre (Talador, Minero...), Dobles espacios, Palabras mal sonantes o cualquier otro tipo de cosas que no lo hagan apropiado para un juego de Rol. Gracias."
            Exit Function

        End If

    Next

    'pluto:7.0 quito advertencias
    'frmAdverten.Show vbModal
    'If NameCorrecto = True Then
    CheckData = True
    'NameCorrecto = False
    'Else
    'CheckData = False
    'End If

End Function

Private Sub boton_Click(index As Integer)

    Call audio.PlayWave(SND_CLICK)

    Select Case index

    Case 0

        'Dim i As Integer
        'Dim k As Object
        'i = 1
        'For Each k In Skill
        '   UserSkills(i) = k.Caption
        '  i = i + 1
        ' Next

        UserName = txtNombre.Text

        If Right(UserName, 1) = " " Then
            UserName = RTrim(UserName)
            MsgBox "Nombre invalido, se han removido los espacios al final del nombre"

        End If

        UserRaza = lstRaza.List(lstRaza.ListIndex)
        UserSexo = lstGenero.List(lstGenero.ListIndex)
        UserClase = lstProfesion.List(lstProfesion.ListIndex)

        'pluto:7.0 Incompatibilidades razas/clases
        Select Case UserRaza

        Case "Humano"

            'If UserClase = "Bardo" Or UserClase = "Cazador" Then
            'MsgBox ("Los Humanos no pueden ser: Bardos y Cazadores")
            'Exit Sub
            'End If
        Case "Elfo"

            If UserClase = "Guerrero" Or UserClase = "Cazador" Or UserClase = "Pirata" Or UserClase = _
               "Bandido" Or UserClase = "Ladron" Or UserClase = "Le�ador" Or UserClase = "Ermita�o" Or _
               UserClase = "Pescador" Or UserClase = "Carpintero" Or UserClase = "Herrero" Or UserClase _
               = "Minero" Or UserClase = "Domador" Then
                MsgBox ( _
                       "Los Elfos no pueden ser: Guerreros, Cazadores, Piratas, Bandidos, Ladrones, Le�adores, Ermita�os, Pescadores, Carpinteros, Herreros, Mineros o Domadores.")
                Exit Sub

            End If

        Case "Elfo Oscuro"

            If UserClase = "Bardo" Then
                MsgBox ("Los Elfos Oscuros no pueden ser: Bardos")
                Exit Sub

            End If

        Case "Enano"

            If UserClase = "Pirata" Or UserClase = "Druida" Or UserClase = "Mago" Or UserClase = "Arquero" Then
                MsgBox ("Los Enanos no pueden ser: Piratas, Druidas, Magos o Arqueros")
                Exit Sub

            End If

        Case "Gnomo"
            'If UserClase = "Druida" Or UserClase = "Pirata" Or UserClase = "Guerrero" Then
            'MsgBox ("Los Gnomos no pueden ser: Druidas, Piratas o Guerreros")
            'Exit Sub
            'End If

        Case "Goblin"

            If UserClase = "Mago" Then
                MsgBox ("Los Goblins no pueden ser: Magos.")
                Exit Sub

            End If

        Case "Orco"

            If UserClase = "Paladin" Or UserClase = "Bardo" _
               Or UserClase = "Druida" Or UserClase = "Mago" Or UserClase = "Ladron" Or UserClase = _
               "Arquero" Then
                MsgBox ( _
                       "Los Orcos no pueden ser: Paladines, Bardos, Druidas,Mago, Ladrones o Arqueros.")
                Exit Sub

            End If

        Case "Vampiro"
            'If UserClase = "Arquero" Or UserClase = "Druida" Or UserClase = "Bardo" Then
            'MsgBox ("Los Vampiros no pueden ser: Arqueros, Druidas o Bardos")
            'Exit Sub
            'End If

        Case "Abisario"

            If UserClase = "Cazador" Or UserClase = "Druida" Or UserClase = "Mago" Or UserClase = "Arquero" Then
                MsgBox ("Los Abisarios no pueden ser: Cazadores, Druidas, Arqueros o Magos.")
                Exit Sub

            End If

        End Select

        UserAtributos(1) = Val(lbFuerzaEx.Caption)
        UserAtributos(2) = Val(lbInteligenciaEx.Caption)
        UserAtributos(3) = Val(lbAgilidadEx.Caption)
        UserAtributos(4) = Val(lbCarismaEx.Caption)
        UserAtributos(5) = Val(lbConstitucionEx.Caption)
        'pluto:7.0
        UserPorcentajes(1) = Val(lblDa�oProye.Caption)
        UserPorcentajes(2) = Val(lblDa�oCC.Caption)
        UserPorcentajes(3) = Val(lblDa�oMagia.Caption)
        UserPorcentajes(4) = Val(lblResisMagia.Caption)
        UserPorcentajes(5) = Val(lblEvasion.Caption)
        UserPorcentajes(6) = Val(lblDefensaFisica.Caption)

        'pluto:7.0
        'If Not CheckData() Then Exit Sub
        'frmCrearPersonaje.Visible = False
        'FrmConfirmarPersonaje.Show vbModal
        'Exit Sub

        UserHogar = Label5.Caption

        If Not CheckData() Then Exit Sub    'frmPasswd.Show vbModal
        If MsgBox("�Esta seguro que desea crear este personaje?", vbYesNo) = vbYes Then
            SendNewChar = True
            Me.MousePointer = 11

            '   If Not frmMain.Socket1.Connected Then
            '      frmMain.Socket1.HostName = CurServerIp
            '      frmMain.Socket1.RemotePort = CurServerPort
            '     Load frmMain.Socket1
            '        frmMain.Socket1.Connect
            'Else
            'pluto:2.5.0
            KeyCodi = ""
            Keycodi2 = ""
            'Call SendData("gIvEmEvAlcOde")
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

            ' End If
        End If

    Case 1
        'If Musica = 0 Then
        CurMidi = DirMidi & "2.mid"
        LoopMidi = 1
        Call audio.PlayMIDI(CStr(CurMidi) & ".mid", LoopMidi)
        'End If

        'frmConnect.FONDO.Picture = LoadPicture(App.Path & "\Graficos\conectar8.jpg")
        Unload frmCrearPersonaje
        Me.Visible = False
        frmCuentas.Visible = True

    Case 2

        If TimeDado = True Then
            Call audio.PlayWave(SND_DICE)
            'Call TirarDados
            TimeDado = False

        End If

    End Select

End Sub

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

    Randomize Timer

    RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound

    If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function

Public Sub TirarDados()
    Dim a18 As Byte
novalen:
    a18 = 0
    lbFuerza.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
    lbInteligencia.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
    lbAgilidad.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
    lbCarisma.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
    lbConstitucion.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))

    If Val(lbFuerza.Caption) < 17 Then lbFuerza.Caption = Val(lbFuerza.Caption) + 2
    If Val(lbInteligencia.Caption) < 17 Then lbInteligencia.Caption = Val(lbInteligencia.Caption) + 2
    If Val(lbAgilidad.Caption) < 17 Then lbAgilidad.Caption = Val(lbAgilidad.Caption) + 2
    If Val(lbCarisma.Caption) < 17 Then lbCarisma.Caption = Val(lbCarisma.Caption) + 2
    If Val(lbConstitucion.Caption) < 17 Then lbConstitucion.Caption = Val(lbConstitucion.Caption) + 2

    'pluto:2.17
    If Val(lbFuerza.Caption) = 18 Then a18 = a18 + 1
    If Val(lbInteligencia.Caption) = 18 Then a18 = a18 + 1
    If Val(lbAgilidad.Caption) = 18 Then a18 = a18 + 1
    If Val(lbCarisma.Caption) = 18 Then a18 = a18 + 1
    If Val(lbConstitucion.Caption) = 18 Then a18 = a18 + 1
    If a18 > 2 Then GoTo novalen

End Sub

Private Sub Command1_Click(index As Integer)
    Call audio.PlayWave(SND_CLICK)

    Dim Indice
    'If Index Mod 2 = 0 Then
    '   If SkillPoints > 0 Then
    '      Indice = Index \ 2
    '     Skill(Indice).Caption = Val(Skill(Indice).Caption) + 1
    '    SkillPoints = SkillPoints - 1
    'End If
    'Else
    '   If SkillPoints < 10 Then

    '      Indice = Index \ 2
    '     If Val(Skill(Indice).Caption) > 0 Then
    '        Skill(Indice).Caption = Val(Skill(Indice).Caption) - 1
    '       SkillPoints = SkillPoints + 1
    '  End If
    'End If
    'End If

    'puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
    SkillPoints = 10
    'Puntos.Caption = SkillPoints
    Me.Picture = LoadPicture(App.Path & "\graficos\CrearPj2.JPG")
    'imgHogar.Picture = LoadPicture(App.Path & "\graficos\CP-Ullathorpe.jpg")

    'pluto:7.0
    'Call TirarDados
    lbFuerza.Caption = 16
    lbInteligencia.Caption = 16
    lbAgilidad.Caption = 16
    lbCarisma.Caption = 16
    lbConstitucion.Caption = 16
    lbRestantes.Caption = 6
    Dim i As Integer
    lstProfesion.Clear

    For i = LBound(ListaClases) To UBound(ListaClases)
        lstProfesion.AddItem ListaClases(i)
    Next i

    lstProfesion.ListIndex = 0
    lstGenero.ListIndex = 0
    lstRaza.ListIndex = 0
    MemoAgi = 0
    MemoFue = 0
    'lstProfesion.ListIndex = 1

    'Image1.Picture = LoadPicture(App.Path & "\graficos\" & lstProfesion.Text & ".jpg")
    'Call TirarDados

End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label3_Click()
    Dim ie As Object
    Dim variable As String
    variable = "http://juegosdrag.es/aomanual/"
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate variable

End Sub

Private Sub lbBajaAg_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lbRestantes.Caption)
    Sube = Val(lbAgilidad.Caption)

    If Sube > 16 Then
        Restantes = Restantes + 1
        lbRestantes.Caption = Restantes
        Sube = Sube - 1
        lbAgilidad.Caption = Sube
        lbAgilidadEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbBajaAgEx_Click()
    Dim Restantes As Byte
    Dim Sube As Byte

    Restantes = Val(Me.RestantesEx1.Caption)
    Sube = Val(lbAgilidadEx.Caption)

    If Sube > Val(lbAgilidad.Caption) And MemoAgi > 0 Then
        Restantes = Restantes + 1
        RestantesEx1.Caption = Restantes
        Sube = Sube - 1
        MemoAgi = MemoAgi - 1
        lbAgilidadEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbBajaCaEx_Click()
    Dim Restantes As Byte
    Dim Sube As Byte

    Restantes = Val(Me.lbRestantesEx2.Caption)
    Sube = Val(lbCarismaEx.Caption)

    If Sube > Val(lbCarisma.Caption) Then
        Restantes = Restantes + 1
        lbRestantesEx2.Caption = Restantes
        Sube = Sube - 1

        lbCarismaEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbBajaCoEx_Click()
    Dim Restantes As Byte
    Dim Sube As Byte

    Restantes = Val(Me.lbRestantesEx2.Caption)
    Sube = Val(lbConstitucionEx.Caption)

    If Sube > Val(lbConstitucion.Caption) Then
        Restantes = Restantes + 1
        lbRestantesEx2.Caption = Restantes
        Sube = Sube - 1

        lbConstitucionEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbBajaFuEx_Click()
    Dim Restantes As Byte
    Dim Sube As Byte

    Restantes = Val(Me.RestantesEx1.Caption)
    Sube = Val(lbFuerzaEx.Caption)

    If Sube > Val(lbFuerza.Caption) And MemoFue > 0 Then
        Restantes = Restantes + 1
        RestantesEx1.Caption = Restantes
        Sube = Sube - 1
        MemoFue = MemoFue - 1

        lbFuerzaEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbBajaInEx_Click()
    Dim Restantes As Byte
    Dim Sube As Byte

    Restantes = Val(Me.lbRestantesEx2.Caption)
    Sube = Val(lbInteligenciaEx.Caption)

    If Sube > Val(lbInteligencia.Caption) Then
        Restantes = Restantes + 1
        lbRestantesEx2.Caption = Restantes
        Sube = Sube - 1

        lbInteligenciaEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lblBajaDa�oCC_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lblPorcRestantes.Caption)
    Sube = Val(lblDa�oCC.Caption)

    If Sube > 0 Then
        Restantes = Restantes + 1
        lblPorcRestantes.Caption = Restantes & "%"
        Sube = Sube - 1
        lblDa�oCC.Caption = Sube & "%"
        lstProfesion_Click

    End If

End Sub

Private Sub lblBajaDa�oMagia_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lblPorcRestantes.Caption)
    Sube = Val(lblDa�oMagia.Caption)

    If Sube > 0 Then
        Restantes = Restantes + 1
        lblPorcRestantes.Caption = Restantes & "%"
        Sube = Sube - 1
        lblDa�oMagia.Caption = Sube & "%"
        lstProfesion_Click

    End If

End Sub

Private Sub lblBajaDa�oProye_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lblPorcRestantes.Caption)
    Sube = Val(lblDa�oProye.Caption)

    If Sube > 0 Then
        Restantes = Restantes + 1
        lblPorcRestantes.Caption = Restantes & "%"
        Sube = Sube - 1
        lblDa�oProye.Caption = Sube & "%"
        lstProfesion_Click

    End If

End Sub

Private Sub lblBajaDefensaFisica_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lblPorcRestantes.Caption)
    Sube = Val(lblDefensaFisica.Caption)

    If Sube > 0 Then
        Restantes = Restantes + 1
        lblPorcRestantes.Caption = Restantes & "%"
        Sube = Sube - 1
        lblDefensaFisica.Caption = Sube & "%"
        lstProfesion_Click

    End If

End Sub

Private Sub lblBajaEvasion_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lblPorcRestantes.Caption)
    Sube = Val(lblEvasion.Caption)

    If Sube > 0 Then
        Restantes = Restantes + 1
        lblPorcRestantes.Caption = Restantes & "%"
        Sube = Sube - 1
        lblEvasion.Caption = Sube & "%"
        lstProfesion_Click

    End If

End Sub

Private Sub lblBajaResisMagia_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lblPorcRestantes.Caption)
    Sube = Val(lblResisMagia.Caption)

    If Sube > 0 Then
        Restantes = Restantes + 1
        lblPorcRestantes.Caption = Restantes & "%"
        Sube = Sube - 1
        lblResisMagia.Caption = Sube & "%"
        lstProfesion_Click

    End If

End Sub

Private Sub lblSubeDa�oCC_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lblPorcRestantes.Caption)
    Sube = Val(lblDa�oCC.Caption)

    If Restantes > 0 And Sube < 5 Then
        Restantes = Restantes - 1
        lblPorcRestantes.Caption = Restantes & "%"
        Sube = Sube + 1
        lblDa�oCC = Sube & "%"
        lstProfesion_Click

    End If

End Sub

Private Sub lblSubeDa�oMagia_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lblPorcRestantes.Caption)
    Sube = Val(lblDa�oMagia.Caption)

    If Restantes > 0 And Sube < 5 Then
        Restantes = Restantes - 1
        lblPorcRestantes.Caption = Restantes & "%"
        Sube = Sube + 1
        lblDa�oMagia = Sube & "%"
        lstProfesion_Click

    End If

End Sub

Private Sub lblSubeDa�oProye_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lblPorcRestantes.Caption)
    Sube = Val(lblDa�oProye.Caption)

    If Restantes > 0 And Sube < 5 Then
        Restantes = Restantes - 1
        lblPorcRestantes.Caption = Restantes & "%"
        Sube = Sube + 1
        lblDa�oProye = Sube & "%"
        lstProfesion_Click

    End If

End Sub

Private Sub lblSubeDefensaFisica_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lblPorcRestantes.Caption)
    Sube = Val(lblDefensaFisica.Caption)

    If Restantes > 0 And Sube < 5 Then
        Restantes = Restantes - 1
        lblPorcRestantes.Caption = Restantes & "%"
        Sube = Sube + 1
        lblDefensaFisica = Sube & "%"
        lstProfesion_Click

    End If

End Sub

Private Sub lblSubeEvasion_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lblPorcRestantes.Caption)
    Sube = Val(lblEvasion.Caption)

    If Restantes > 0 And Sube < 5 Then
        Restantes = Restantes - 1
        lblPorcRestantes.Caption = Restantes & "%"
        Sube = Sube + 1
        lblEvasion = Sube & "%"
        lstProfesion_Click

    End If

End Sub

Private Sub lblSubeResisMagia_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lblPorcRestantes.Caption)
    Sube = Val(lblResisMagia.Caption)

    If Restantes > 0 And Sube < 5 Then
        Restantes = Restantes - 1
        lblPorcRestantes.Caption = Restantes & "%"
        Sube = Sube + 1
        lblResisMagia = Sube & "%"
        lstProfesion_Click

    End If

End Sub

Private Sub lbSubeAg_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lbRestantes.Caption)
    Sube = Val(lbAgilidad.Caption)

    If Restantes > 0 And Sube < 18 Then
        Restantes = Restantes - 1
        lbRestantes.Caption = Restantes
        Sube = Sube + 1
        lbAgilidad.Caption = Sube
        lbAgilidadEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbBajaFu_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lbRestantes.Caption)
    Sube = Val(lbFuerza.Caption)

    If Sube > 16 Then
        Restantes = Restantes + 1
        lbRestantes.Caption = Restantes
        Sube = Sube - 1
        lbFuerza.Caption = Sube
        lbFuerzaEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbSubeAgEx_Click()
    Dim Restantes As Byte
    Dim Sube As Byte

    Restantes = Val(Me.RestantesEx1.Caption)
    Sube = Val(lbAgilidadEx.Caption)

    If Restantes > 0 And MemoAgi < 3 Then
        Restantes = Restantes - 1
        RestantesEx1.Caption = Restantes
        Sube = Sube + 1
        MemoAgi = MemoAgi + 1

        lbAgilidadEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbSubeCaEx_Click()
    Dim Restantes As Byte
    Dim Sube As Byte

    Restantes = Val(Me.lbRestantesEx2.Caption)
    Sube = Val(lbCarismaEx.Caption)

    If Restantes > 0 Then
        Restantes = Restantes - 1
        lbRestantesEx2.Caption = Restantes
        Sube = Sube + 1

        lbCarismaEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbSubeCoEx_Click()
    Dim Restantes As Byte
    Dim Sube As Byte

    Restantes = Val(Me.lbRestantesEx2.Caption)
    Sube = Val(lbConstitucionEx.Caption)

    If Restantes > 0 Then
        Restantes = Restantes - 1
        lbRestantesEx2.Caption = Restantes
        Sube = Sube + 1

        lbConstitucionEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbSubeFu_Click()
    Dim Restantes As Byte
    Dim Sube As Byte

    Restantes = Val(lbRestantes.Caption)
    Sube = Val(lbFuerza.Caption)

    If Restantes > 0 And Sube < 18 Then
        Restantes = Restantes - 1
        lbRestantes.Caption = Restantes
        Sube = Sube + 1

        lbFuerza.Caption = Sube
        lbFuerzaEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbBajaco_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lbRestantes.Caption)
    Sube = Val(lbConstitucion.Caption)

    If Sube > 16 Then
        Restantes = Restantes + 1
        lbRestantes.Caption = Restantes
        Sube = Sube - 1
        lbConstitucion.Caption = Sube
        lbConstitucionEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbSubeco_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lbRestantes.Caption)
    Sube = Val(lbConstitucion.Caption)

    If Restantes > 0 And Sube < 18 Then
        Restantes = Restantes - 1
        lbRestantes.Caption = Restantes
        Sube = Sube + 1
        lbConstitucion.Caption = Sube
        lbConstitucionEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbBajaca_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lbRestantes.Caption)
    Sube = Val(lbCarisma.Caption)

    If Sube > 16 Then
        Restantes = Restantes + 1
        lbRestantes.Caption = Restantes
        Sube = Sube - 1
        lbCarisma.Caption = Sube
        lbCarismaEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbSubeca_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lbRestantes.Caption)
    Sube = Val(lbCarisma.Caption)

    If Restantes > 0 And Sube < 18 Then
        Restantes = Restantes - 1
        lbRestantes.Caption = Restantes
        Sube = Sube + 1
        lbCarisma.Caption = Sube
        lbCarismaEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbBajain_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lbRestantes.Caption)
    Sube = Val(lbInteligencia.Caption)

    If Sube > 16 Then
        Restantes = Restantes + 1
        lbRestantes.Caption = Restantes
        Sube = Sube - 1
        lbInteligencia.Caption = Sube
        lbInteligenciaEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbSubeFuEx_Click()
    Dim Restantes As Byte
    Dim Sube As Byte

    Restantes = Val(Me.RestantesEx1.Caption)
    Sube = Val(lbFuerzaEx.Caption)

    If Restantes > 0 And MemoFue < 3 Then
        Restantes = Restantes - 1
        RestantesEx1.Caption = Restantes
        Sube = Sube + 1
        MemoFue = MemoFue + 1

        lbFuerzaEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbSubein_Click()
    Dim Restantes As Byte
    Dim Sube As Byte
    Restantes = Val(lbRestantes.Caption)
    Sube = Val(lbInteligencia.Caption)

    If Restantes > 0 And Sube < 18 Then
        Restantes = Restantes - 1
        lbRestantes.Caption = Restantes
        Sube = Sube + 1
        lbInteligencia.Caption = Sube
        lbInteligenciaEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lbSubeInEx_Click()
    Dim Restantes As Byte
    Dim Sube As Byte

    Restantes = Val(Me.lbRestantesEx2.Caption)
    Sube = Val(lbInteligenciaEx.Caption)

    If Restantes > 0 Then
        Restantes = Restantes - 1
        lbRestantesEx2.Caption = Restantes
        Sube = Sube + 1

        lbInteligenciaEx.Caption = Sube
        lstProfesion_Click

    End If

End Sub

Private Sub lstProfesion_Click()

    On Error Resume Next

    Select Case UCase$(lstProfesion.List(lstProfesion.ListIndex))

    Case "MAGO"
        Me.Pevasion.Caption = 80
        Me.Paciertoarmas.Caption = 50
        Me.Paciertoproyec.Caption = 50
        Me.Pda�oarmas.Caption = 50
        Me.Pda�oproyec.Caption = 50
        Me.Pescudos.Caption = 60
        Me.PDa�oMagias.Caption = 120
        Me.PResisMagia.Caption = 120

    Case "GUERRERO"
        Me.Pevasion.Caption = 100
        Me.Paciertoarmas.Caption = 100
        Me.Paciertoproyec.Caption = 80
        Me.Pda�oarmas.Caption = 110
        Me.Pda�oproyec.Caption = 100
        Me.Pescudos.Caption = 100
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

    Case "CAZADOR"
        Me.Pevasion.Caption = 90
        Me.Paciertoarmas.Caption = 80
        Me.Paciertoproyec.Caption = 120
        Me.Pda�oarmas.Caption = 90
        Me.Pda�oproyec.Caption = 90
        Me.Pescudos.Caption = 80
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

    Case "PALADIN"
        Me.Pevasion.Caption = 90
        Me.Paciertoarmas.Caption = 85
        Me.Paciertoproyec.Caption = 75
        Me.Pda�oarmas.Caption = 90
        Me.Pda�oproyec.Caption = 80
        Me.Pescudos.Caption = 100
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

    Case "BANDIDO"
        Me.Pevasion.Caption = 90
        Me.Paciertoarmas.Caption = 85
        Me.Paciertoproyec.Caption = 90
        Me.Pda�oarmas.Caption = 80
        Me.Pda�oproyec.Caption = 75
        Me.Pescudos.Caption = 80
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

    Case "ASESINO"
        Me.Pevasion.Caption = 110
        Me.Paciertoarmas.Caption = 85
        Me.Paciertoproyec.Caption = 75
        Me.Pda�oarmas.Caption = 90
        Me.Pda�oproyec.Caption = 80
        Me.Pescudos.Caption = 80
        Me.PDa�oMagias.Caption = 100
        Me.PResisMagia.Caption = 100

    Case "PIRATA"
        Me.Pevasion.Caption = 100
        Me.Paciertoarmas.Caption = 95
        Me.Paciertoproyec.Caption = 85
        Me.Pda�oarmas.Caption = 90
        Me.Pda�oproyec.Caption = 85
        Me.Pescudos.Caption = 85
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

    Case "LADRON"
        Me.Pevasion.Caption = 110
        Me.Paciertoarmas.Caption = 75
        Me.Paciertoproyec.Caption = 80
        Me.Pda�oarmas.Caption = 80
        Me.Pda�oproyec.Caption = 75
        Me.Pescudos.Caption = 70
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

    Case "BARDO"
        Me.Pevasion.Caption = 120
        Me.Paciertoarmas.Caption = 80
        Me.Paciertoproyec.Caption = 70
        Me.Pda�oarmas.Caption = 80
        Me.Pda�oproyec.Caption = 70
        Me.Pescudos.Caption = 75
        Me.PDa�oMagias.Caption = 110
        Me.PResisMagia.Caption = 110

    Case "CLERIGO"
        Me.Pevasion.Caption = 80
        Me.Paciertoarmas.Caption = 70
        Me.Paciertoproyec.Caption = 70
        Me.Pda�oarmas.Caption = 80
        Me.Pda�oproyec.Caption = 70
        Me.Pescudos.Caption = 90
        Me.PDa�oMagias.Caption = 110
        Me.PResisMagia.Caption = 110

    Case "DRUIDA"
        Me.Paciertoarmas.Caption = 70
        Me.Pevasion.Caption = 80
        Me.Paciertoproyec.Caption = 75
        Me.Pda�oarmas.Caption = 75
        Me.Pda�oproyec.Caption = 75
        Me.Pescudos.Caption = 75
        Me.PDa�oMagias.Caption = 100
        Me.PResisMagia.Caption = 100

    Case "ARQUERO"
        Me.Paciertoarmas.Caption = 50
        Me.Pevasion.Caption = 80
        Me.Paciertoproyec.Caption = 120
        Me.Pda�oproyec.Caption = 130
        Me.Pda�oarmas.Caption = 50
        Me.Pescudos.Caption = 60
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

    Case "PESCADOR"
        Me.Paciertoarmas.Caption = 60
        Me.Pevasion.Caption = 80
        Me.Paciertoproyec.Caption = 65
        Me.Pda�oarmas.Caption = 60
        Me.Pda�oproyec.Caption = 60
        Me.Pescudos.Caption = 70
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

    Case "LE�ADOR"
        Me.Paciertoarmas.Caption = 60
        Me.Pevasion.Caption = 80
        Me.Paciertoproyec.Caption = 70
        Me.Pda�oarmas.Caption = 70
        Me.Pda�oproyec.Caption = 60
        Me.Pescudos.Caption = 70
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

    Case "MINERO"
        Me.Paciertoarmas.Caption = 60
        Me.Pevasion.Caption = 80
        Me.Paciertoproyec.Caption = 65
        Me.Pda�oarmas.Caption = 75
        Me.Pda�oproyec.Caption = 70
        Me.Pescudos.Caption = 70
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

    Case "HERRERO"

        Me.Paciertoarmas.Caption = 60
        Me.Pevasion.Caption = 80
        Me.Paciertoproyec.Caption = 65
        Me.Pda�oarmas.Caption = 75
        Me.Pda�oproyec.Caption = 70
        Me.Pescudos.Caption = 70
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

    Case "CARPINTERO"
        Me.Paciertoarmas.Caption = 60
        Me.Pevasion.Caption = 80
        Me.Paciertoproyec.Caption = 70
        Me.Pda�oarmas.Caption = 70
        Me.Pda�oproyec.Caption = 70
        Me.Pescudos.Caption = 70
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

    Case "ERMITA�O"
        Me.Paciertoarmas.Caption = 80
        Me.Pevasion.Caption = 80
        Me.Paciertoproyec.Caption = 75
        Me.Pda�oarmas.Caption = 80
        Me.Pda�oproyec.Caption = 75
        Me.Pescudos.Caption = 80
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

    Case "DOMADOR"
        Me.Paciertoarmas.Caption = 50
        Me.Pevasion.Caption = 80
        Me.Paciertoproyec.Caption = 50
        Me.Pda�oarmas.Caption = 50
        Me.Pda�oproyec.Caption = 50
        Me.Pescudos.Caption = 60
        Me.PDa�oMagias.Caption = 90
        Me.PResisMagia.Caption = 90

        'Case Else
        '   Me.Pevasion.Caption = 80
    End Select

    'pluto:7.0 a�adiendo bonus raza
    If UCase$(lstRaza.List(lstRaza.ListIndex)) = "ELFO OSCURO" Then
        Me.Pevasion.Caption = Me.Pevasion.Caption + 10

    End If

    If UCase$(lstRaza.List(lstRaza.ListIndex)) = "ENANO" Then
        Me.Pescudos.Caption = Me.Pescudos.Caption + 20

    End If

    If UCase$(lstRaza.List(lstRaza.ListIndex)) = "GNOMO" Then
        Me.Paciertoarmas.Caption = Me.Paciertoarmas.Caption + 10

    End If

    'pluto:7.0 Caculando potencial
    Dim valo As Double
    'evasion
    valo = Val(Me.Pevasion.Caption) / 100
    Me.Pevasion.Caption = Round(Me.lbAgilidadEx.Caption * valo)
    Me.Pevasion.Caption = Round(Me.Pevasion.Caption + porcentaje(Me.Pevasion.Caption, Val(lblEvasion.Caption)))
    Me.Pevasion2.Caption = Me.Pevasion
    'acierto armas
    valo = Val(Me.Paciertoarmas.Caption) / 100
    Me.Paciertoarmas.Caption = Round(Me.lbAgilidadEx.Caption * valo)
    'acierto proyec
    valo = Val(Me.Paciertoproyec.Caption) / 100
    Me.Paciertoproyec.Caption = Round(Me.lbAgilidadEx.Caption * valo)
    'def.escudos
    valo = Val(Me.Pescudos.Caption) / 100
    Me.Pescudos.Caption = Round(valo * 10)
    'resistencia magias
    valo = 20 * Val(Me.PResisMagia.Caption) / 100
    Me.PResisMagia.Caption = Round(valo + porcentaje(valo, Val(lblResisMagia.Caption)))
    'da�o magias
    valo = 20 * Val(Me.PDa�oMagias.Caption) / 100
    valo = valo + porcentaje(valo, 3)
    Me.PDa�oMagias.Caption = Round(valo + porcentaje(valo, Val(lblDa�oMagia.Caption)))

    'da�o armas c/c
    valo = Val(Me.Pda�oarmas.Caption) / 100
    Me.Pda�oarmas.Caption = ((3 * 3) + ((3 / 5) * (Me.lbFuerzaEx.Caption - 15) + 2) * valo)
    Me.Pda�oarmas.Caption = Round(Me.Pda�oarmas.Caption + porcentaje(Me.Pda�oarmas.Caption, Val(lblDa�oCC.Caption)))
    'da�o proyectiles
    valo = Val(Me.Pda�oproyec.Caption) / 100
    Me.Pda�oproyec.Caption = (3 * 4) + ((3 / 5) * (Me.lbFuerzaEx.Caption - 15) + 2) * valo
    Me.Pda�oproyec.Caption = Round(Me.Pda�oproyec.Caption + porcentaje(Me.Pda�oproyec.Caption, Val( _
                                                                                               lblDa�oProye.Caption)))
    'defensa fisica
    Me.PDefensafisica.Caption = 20

    'Image1.Picture = LoadPicture(App.Path & "\graficos\" & lstProfesion.Text & ".jpg")
End Sub

Private Sub lstRaza_Click()
'pluto:2.17

    Select Case UCase$(lstRaza.List(lstRaza.ListIndex))

        'pluto:7.0
    Case "HUMANO"
        ' LabelBonus(0).Caption = "+1"
        ' LabelBonus(1).Caption = "+2"
        ' LabelBonus(4).Caption = "+2"
        ' LabelBonus(2).Caption = "+1"
        'LabelBonus(3).Caption = "+0"
        Label1.Caption = "Humanos: Reciben mejores efectos al tomar Pociones."
        Label5.Caption = "Dungeon Newbie"

    Case "ELFO"
        ' LabelBonus(0).Caption = "-1"
        ' LabelBonus(1).Caption = "+2"
        ' LabelBonus(2).Caption = "+2"
        ' LabelBonus(3).Caption = "+2"
        ' LabelBonus(4).Caption = "+1"
        Label1.Caption = "Elfos: Gastan un 15% menos de mana al usar magias."
        Label5.Caption = "Dungeon Newbie"

    Case "ELFO OSCURO"
        'LabelBonus(0).Caption = "+1"
        'LabelBonus(1).Caption = "+2"
        'LabelBonus(2).Caption = "-2"
        'LabelBonus(3).Caption = "+2"
        'LabelBonus(4).Caption = "+1"
        Label1.Caption = _
        "Elfos Oscuros: Dope de Agilidad permanente. Invisibilidad +33% duraci�n y obtiene +10 en Evasi�n."
        Label5.Caption = "Dungeon Newbie"

    Case "ENANO"
        ' LabelBonus(0).Caption = "+3"
        'LabelBonus(1).Caption = "-1"
        'LabelBonus(2).Caption = "-3"
        'LabelBonus(3).Caption = "+0"
        'LabelBonus(4).Caption = "+3"
        Label1.Caption = "Enanos: Dope Fuerza Permanente. Reduce 50% tiempo Paralisis y +20 en Defensa de Escudos."
        Label5.Caption = "Dungeon Newbie"

    Case "GNOMO"
        ' LabelBonus(0).Caption = "-4"
        'LabelBonus(1).Caption = "+3"
        'LabelBonus(2).Caption = "+3"
        'LabelBonus(3).Caption = "+0"
        'LabelBonus(4).Caption = "+1"
        Label1.Caption = "Gnomos: Tienen un  15% de evitar Paralisis y obtienen +10 en Ataque con Armas."
        Label5.Caption = "Dungeon Newbie"

    Case "ORCO"
        ' LabelBonus(0).Caption = "+4"
        'LabelBonus(1).Caption = "-3"
        'LabelBonus(2).Caption = "-6"
        'LabelBonus(3).Caption = "+0"
        'LabelBonus(4).Caption = "+3"
        Label1.Caption = "Orcos: Poseen Habilidad BESERKER. (Consultar Manual para m�s informaci�n)"
        Label5.Caption = "Dungeon Newbie"

    Case "VAMPIRO"
        '  LabelBonus(0).Caption = "+2"
        ' LabelBonus(1).Caption = "+2"
        ' LabelBonus(2).Caption = "+0"
        ' LabelBonus(3).Caption = "+0"
        'LabelBonus(4).Caption = "+2"
        Label5.Caption = "Dungeon Newbie"
        Label1.Caption = "Vampiros: Regeneran Salud. Transformarci�n en Murcielagos. Teleportaci�n a Ciudades."
        Label5.Caption = "Dungeon Newbie"

    Case "ABISARIO"
        Label5.Caption = "Dungeon Newbie"
        Label1.Caption = "Abisarios: Golpes cr�ticos x1.25 a otros jugadores (1/15). Adem�s 10% de quedar con 1 de vida al recibir golpe y no morir."


    Case "GOBLIN"
        Label5.Caption = "Dungeon Newbie"
        Label1.Caption = _
        "GOBLINS: Roba oro por golpe. Con invisibilidad 30% de no oirse tus pasos. Adem�s 25% no se caiga el inventario al morir."

    End Select

    frmCrearPersonaje.Label2.Visible = True
    frmCrearPersonaje.Label3.Visible = True

    'pluto:7.0 dope agilidad en elfos
    If UCase$(lstRaza.List(lstRaza.ListIndex)) = "ELFO OSCURO" Then

        If Me.lbAgilidadEx < 25 Then Me.lbAgilidadEx = Me.lbAgilidadEx + 13
    Else

        If Me.lbAgilidadEx > 25 Then Me.lbAgilidadEx = Me.lbAgilidadEx - 13

    End If

    'pluto:7.0 dope fuerza enanos
    If UCase$(lstRaza.List(lstRaza.ListIndex)) = "ENANO" Then

        If Me.lbFuerzaEx < 25 Then Me.lbFuerzaEx = Me.lbFuerzaEx + 13
    Else

        If Me.lbFuerzaEx > 25 Then Me.lbFuerzaEx = Me.lbFuerzaEx - 13

    End If

    'actualizo marcadores
    lstProfesion_Click

End Sub

Private Sub Timer1_Timer()
    TimeDado = True

End Sub

Private Sub Timer2_Timer()
    Static FONDO As Boolean

    If Val(frmCrearPersonaje.lbRestantes.Caption) > 0 And FONDO = False Then
        FONDO = True
        Me.Picture = LoadPicture(App.Path & "\graficos\CrearPj2.JPG")
        Me.lbAgilidadEx.Visible = False
        Me.lbConstitucionEx.Visible = False
        Me.lbFuerzaEx.Visible = False
        Me.lbCarismaEx.Visible = False
        Me.lbInteligenciaEx.Visible = False
        Me.lbBajaAgEx.Visible = False
        Me.lbBajaCaEx.Visible = False
        Me.lbBajaCoEx.Visible = False
        Me.lbBajaFuEx.Visible = False
        Me.lbBajaInEx.Visible = False
        Me.lbSubeAgEx.Visible = False
        Me.lbSubeCaEx.Visible = False
        Me.lbSubeCoEx.Visible = False
        Me.lbSubeFuEx.Visible = False
        Me.lbSubeInEx.Visible = False

        Me.lbRestantesEx2.Visible = False
        Me.RestantesEx1.Visible = False

        Me.lbAgilidadEx.Caption = Me.lbAgilidad.Caption
        Me.lbFuerzaEx.Caption = Me.lbFuerza.Caption
        Me.lbConstitucionEx.Caption = Me.lbConstitucion.Caption
        Me.lbCarismaEx.Caption = Me.lbCarisma.Caption
        Me.lbInteligenciaEx.Caption = Me.lbInteligencia.Caption
        Me.RestantesEx1.Caption = 2
        Me.lbRestantesEx2 = 3
        lstProfesion.ListIndex = 0
        lstGenero.ListIndex = 0
        lstRaza.ListIndex = 0
        MemoAgi = 0
        MemoFue = 0

    End If

    If Val(frmCrearPersonaje.lbRestantes.Caption) = 0 And FONDO = True Then
        FONDO = False
        Me.Picture = LoadPicture(App.Path & "\graficos\CrearPj1.JPG")
        Me.lbAgilidadEx.Visible = True
        Me.lbConstitucionEx.Visible = True
        Me.lbFuerzaEx.Visible = True
        Me.lbCarismaEx.Visible = True
        Me.lbInteligenciaEx.Visible = True
        Me.lbRestantesEx2.Visible = True
        Me.RestantesEx1.Visible = True
        Me.lbBajaAgEx.Visible = True
        Me.lbBajaCaEx.Visible = True
        Me.lbBajaCoEx.Visible = True
        Me.lbBajaFuEx.Visible = True
        Me.lbBajaInEx.Visible = True
        Me.lbSubeAgEx.Visible = True
        Me.lbSubeCaEx.Visible = True
        Me.lbSubeCoEx.Visible = True
        Me.lbSubeFuEx.Visible = True
        Me.lbSubeInEx.Visible = True

    End If

    If Val(frmCrearPersonaje.lbRestantes.Caption) = 0 And Val(frmCrearPersonaje.lbRestantesEx2.Caption) = 0 Then
        lstRaza.Enabled = True
        lstProfesion.Enabled = True
        lstGenero.Enabled = True
    Else
        lstRaza.Enabled = False
        lstProfesion.Enabled = False
        lstGenero.Enabled = False
        lstProfesion.ListIndex = 0
        lstGenero.ListIndex = 0
        lstRaza.ListIndex = 0

    End If

End Sub

Private Sub txtNombre_Change()
    txtNombre.Text = LTrim(txtNombre.Text)

End Sub

Private Sub txtNombre_GotFocus()
    MsgBox _
            "Sea cuidadoso al seleccionar el nombre de su personaje,decida bi�n que letras poner en May�sculas y cuales en Min�sculas porque luego no podr�n ser cambiadas. Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotaci�n politica los administradores borrar�n su personaje y no habr� ninguna posibilidad de recuperarlo."
    Call audio.PlayWave("crear.wav")

End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(Chr(KeyAscii))

    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

