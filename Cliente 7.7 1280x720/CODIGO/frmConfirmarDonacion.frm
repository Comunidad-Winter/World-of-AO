VERSION 5.00
Begin VB.Form frmConfirmarDonacion 
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmConfirmarDonacion.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9000
      Top             =   5160
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Asignar DragCreditos"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   3615
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   4080
      MouseIcon       =   "frmConfirmarDonacion.frx":36BEE
      MousePointer    =   99  'Custom
      Picture         =   "frmConfirmarDonacion.frx":378B8
      Top             =   6360
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   8760
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DragCreditos Disponibles:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DragCreditos Necesarios:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   5640
      Width           =   495
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   1200
      Picture         =   "frmConfirmarDonacion.frx":3BB71
      Top             =   2520
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   1335
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   6735
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   6600
      MouseIcon       =   "frmConfirmarDonacion.frx":3E16D
      MousePointer    =   99  'Custom
      Picture         =   "frmConfirmarDonacion.frx":3EE37
      Top             =   6360
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   1560
      MouseIcon       =   "frmConfirmarDonacion.frx":43307
      MousePointer    =   99  'Custom
      Picture         =   "frmConfirmarDonacion.frx":43FD1
      Top             =   6360
      Visible         =   0   'False
      Width           =   1680
   End
End
Attribute VB_Name = "frmConfirmarDonacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.Timer1.Enabled = True

End Sub

Private Sub Image1_Click()

    If Val(Me.Label2) < Val(Me.Label3) Then
        Me.Label1.Caption = _
        "No tienes suficientes DragCreditos. Visita www.juegosdrag.es y en la secci�n Donaciones podr�s obtener toda la informaci�n de como conseguir DragCreditos."

    End If

    SendData (Me.Label6.Caption)

    Select Case Left$(Me.Label6.Caption, 5)

    Case "DRAC1"
        Me.Label1.Caption = _
        "Ok!! Cambio realizado con �xito. Es posible que tengas que salir y volver a entrar al personaje para ver los cambios."

    Case "DRAC2"
        Me.Label1.Caption = _
        "Ok!! Cambio realizado con �xito. Es posible que tengas que salir y volver a entrar al personaje para ver los cambios."

    Case "DRAC3"
        Me.Label1.Caption = _
        "Ok!! Cambio realizado con �xito. Es posible que tengas que salir y volver a entrar al personaje para ver los cambios."

    Case "DRAC4"
        Me.Label1.Caption = _
        "Ok!! Cambio realizado con �xito. Es posible que tengas que salir y volver a entrar al personaje para ver los cambios."

    Case "DRAC5"
        Me.Label1.Caption = _
        "Ok!! Cambio realizado con �xito. Es posible que tengas que salir y volver a entrar al personaje para ver los cambios."

    Case "DRAC6"
        Me.Label1.Caption = _
        "Ok!! Cambio realizado con �xito. Es posible que tengas que salir y volver a entrar al personaje para ver los cambios."

    Case "DRAC7"
        Me.Label1.Caption = _
        "Ok!! Has conseguido nuevas solicitudes de clan. Es posible que tengas que salir y volver a entrar al personaje para ver los cambios."

    Case "DRAC8"
        Me.Label1.Caption = "Ok!! El nuevo objeto ha sido a�adido a tu inventario."

    End Select

    ' Me.Label1.Caption = "Ok!! Es posible que tengas que salir y volver a entrar al personaje para ver los cambios."
    Timer1.Enabled = False
    Me.Image1.Visible = False
    Me.Image2.Visible = False
    Me.Image4.Visible = True

End Sub

Private Sub Image2_Click()
    frmdragcreditos.ncreditos.Caption = Me.Label3
    Unload Me
    frmdragcreditos.Show vbModal

End Sub

Private Sub Image4_Click()
    Unload Me

End Sub

Private Sub Timer1_Timer()
'Debug.Print Me.Image3.Width
'Debug.Print Me.Image3.Left

    If frmConfirmarDonacion.Label6.Caption <> "" Then
        Me.Image1.Visible = True
        Me.Image2.Visible = True
        Me.Image4.Visible = False
    Else
        Me.Image1.Visible = False
        Me.Image2.Visible = False
        Me.Image4.Visible = True

    End If

End Sub
