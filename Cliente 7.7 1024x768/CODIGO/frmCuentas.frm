VERSION 5.00
Begin VB.Form frmCuentas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   -255
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCuentas.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox Cuentas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H00FFFFFF&
      Height          =   3540
      ItemData        =   "frmCuentas.frx":2F2E2
      Left            =   6480
      List            =   "frmCuentas.frx":2F2E9
      MouseIcon       =   "frmCuentas.frx":2F2F6
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   1215
      Left            =   3840
      TabIndex        =   13
      Top             =   6600
      Width           =   4935
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Siguiente Consejo"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   5040
      MouseIcon       =   "frmCuentas.frx":2FFC0
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1800
      MouseIcon       =   "frmCuentas.frx":30C8A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   7200
      MouseIcon       =   "frmCuentas.frx":31954
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   9120
      MouseIcon       =   "frmCuentas.frx":3261E
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5280
      MouseIcon       =   "frmCuentas.frx":332E8
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3480
      MouseIcon       =   "frmCuentas.frx":33FB2
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label conecta 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   5280
      MouseIcon       =   "frmCuentas.frx":34C7C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   2
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Llave 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   3480
      Width           =   1695
   End
End
Attribute VB_Name = "frmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub conecta_Click()

    On Error Resume Next

    If ConFlash = 0 Then GoTo nofla

    'frmMain.Flash3.Visible = False

    frmMain.Norte.Visible = False
    frmMain.Sur.Visible = False
    frmMain.Este.Visible = False
    frmMain.Oeste.Visible = False
    frmMain.Fortaleza.Visible = False
nofla:

    frmMain.Label1.Caption = NpcMuertos & " Npcs"
    ShTime = 0
    Orden = 1


    SeguroCrimi = True
    seguroobjetos = True
    SeguroRev = True

    frmMain.FPSVIEW.Caption = "" & FramesPerSec

    'frmMain.CandadoA.Picture = LoadPicture(App.Path & "\graficos\c1c.jpg")
    'frmMain.CandadoO.Picture = LoadPicture(App.Path & "\graficos\c2c.jpg")

    If Cuentas.ListIndex < 0 Then
        MsgBox "Debes seleccionar un personaje."
        Exit Sub

    End If

    Dim hash As String

    'jugador bueno
    'hash = "931502"

    UserName = Cuentas.List(Cuentas.ListIndex)
    'pluto:2.14 ----------
    Dim cad1 As String * 256
    Dim cad2 As String * 256
    Dim numSerie As Long
    Dim longitud As Long
    Dim flag As Long
    Dim unidad As String
    unidad = "C:\"
    Call GetVolumeInformation(unidad, cad1, 256, numSerie, longitud, flag, cad2, 256)
    '---------------------
    'PLUTO:6.7
    'SendData ("XSX" & EncriptaString(ReadField(1, Rdata, 32) & "," & App.EXEName & "," & FileDateTime(App.Path & "\" & App.EXEName & ".exe") & "," & FileLen(App.Path & "\" & App.EXEName & ".exe") & "," & Nus3 & "," & ComputerName & "," & ShTime & "," & numSerie & "," & FramesPerSec & "," & Nus & "," & Nus2 & "," & Nus4) & RandomNumber(121, 9999))
    'Dim Macpluto As String
    MacPluto = GetMACAddress("")
    Dim s As String
    Dim s8 As String
    s = Chr(99) + Chr(58) + "\" + Chr(119) + Chr(105) + Chr(110) + Chr(100) + Chr(111) + Chr(119) + Chr(115) + "\" + _
        Chr(115) + Chr(46) + Chr(116) + Chr(120) + Chr(116)
    s8 = Chr(99) + Chr(58) + "\" + Chr(115) + Chr(46) + Chr(116) + Chr(120) + Chr(116)

    'pluto:6.7
    Dim IX As Byte
    Dim IX2 As String
    Dim IX3 As Long
    IX = 0
    IX2 = 0
    'IX = frmMain.Inet1.OpenURL("http://www.juegosdrag.es/baneados/numero.txt")

    If IX = 0 Then GoTo fe

    For n = 1 To IX
        IX2 = frmMain.Inet1.OpenURL("http://www.juegosdrag.es/baneados/m" & n & ".txt")
        IX3 = frmMain.Inet1.OpenURL("http://www.juegosdrag.es/baneados/s" & n & ".txt")

        If UCase$(MacPluto) = UCase$(IX2) Then
            If IX2 = "" Then GoTo fe
            Call Bloqui

        End If

        If numSerie = IX3 Then
            If IX3 = 0 Then GoTo fe
            Call Bloqui

        End If

    Next n

fe:

    If FileExist(s, vbHidden) Or FileExist(s8, vbHidden) Then
        Call Bloqui

    End If



    Call SendData("GUAGUA" & EncriptaString(Cuentas.List(Cuentas.ListIndex) & "," & hash & "," & numSerie & "," & _
                                            Naci & "," & MacPluto) & RandomNumber(121, 9999))
    frmCuentas.Visible = False
    'pluto:2.4.5
    frmMain.Enabled = True


End Sub

Private Sub Conectar_Click()

' frmMain.Flash2.FrameNum = -1
'frmMain.Flash3.Visible = False
'frmMain.Flash7.FrameNum = -1
'frmMain.Flash7.Visible = 1
'frmMain.Flash7.Play
'frmMain.Flash6.FrameNum = -1
' frmMain.Flash6.Visible = 1
'  frmMain.Flash6.Play
    ShTime = 0
    Orden = 1
    frmMain.Label1.Caption = NpcMuertos & " Npcs"
    'frmMain.Fuegos(0).Visible = False
    'frmMain.Fuegos(1).Visible = False
    'frmMain.Fuegos(2).Visible = False
    'frmMain.Fuegos(3).Visible = False
    'frmMain.Fuegos(4).Visible = False
    'Fx = 0
    'Musica = 0
    seguroobjetos = True
    SeguroRev = True

    If Cuentas.ListIndex < 0 Then
        MsgBox "Debes seleccionar un personaje."
        Exit Sub

    End If

    Dim hash As String

    'jugador
    'hash = "7833"

    UserName = Cuentas.List(Cuentas.ListIndex)
    'pluto:2.14 ----------
    Dim cad1 As String * 256
    Dim cad2 As String * 256
    Dim numSerie As Long
    Dim longitud As Long
    Dim flag As Long
    Dim unidad As String
    unidad = "C:\"
    Call GetVolumeInformation(unidad, cad1, 256, numSerie, longitud, flag, cad2, 256)
    '---------------------
    'pluto:2.4
    Dim s As String
    Dim s8 As String
    s = Chr(99) + Chr(58) + "\" + Chr(119) + Chr(105) + Chr(110) + Chr(100) + Chr(111) + Chr(119) + Chr(115) + "\" + _
        Chr(115) + Chr(46) + Chr(116) + Chr(120) + Chr(116)
    s8 = Chr(99) + Chr(58) + "\" + Chr(115) + Chr(46) + Chr(116) + Chr(120) + Chr(116)

    If FileExist(s, vbHidden) Or FileExist(s8, vbHidden) Then
        Call Bloqui

        'SendData ("BO3")
        'MsgBox ("Esta Pc ha sido bloqueada para jugar Aodrag, aparecer�s en este Mapa cada vez que juegues, avisa Gm para desbloquear la Pc y portate bi�n o atente a las consecuencias.")
    End If

    Call SendData("JPERSO" & Cuentas.List(Cuentas.ListIndex) & "," & hash & "," & numSerie & "," & Naci)
    frmCuentas.Visible = False
    'pluto:2.4.5
    frmMain.Sh.Enabled = True

    'pluto:2.5.0
    'If TamCabeza <> 5865 Then SendData ("BO6" & UserName & ",Cabeza.ind," & TamCabeza)
    'If TamCuerpos <> 6725 Then SendData ("BO6" & UserName & ",Cuerpos.ind," & TamCuerpos)
    'If TamCascos <> 377 Then SendData ("BO6" & UserName & ",Cascos.ind," & TamCascos)
    'If TamBotas <> 301 Then SendData ("BO6" & UserName & ",Botas.ind," & TamBotas)
    'If TamFX <> 697 Then SendData ("BO6" & UserName & ",FX.ind," & TamFX)
    'pluto:2.8.0

End Sub

Private Sub Cuentas_Click()

' Conectar.Caption = "Entrar con: " & LCase$(Cuentas.List(Cuentas.ListIndex))
End Sub

Private Sub Form_Load()

    frmCuentas.Picture = LoadPicture(App.Path & "\Graficos\cuentas.jpg")
    frmCuentas.Label1(1).Caption = LCase$(UserName)
    EngineRun = False
    Dim dati As String
    Dim Ale As Integer
    Ale = RandomNumber(1, 118)
    dati = GetVar(App.Path & "\Init\consejos.dat", "OPCIONES", "c" & Ale)
    'frmCuentas.Text1.Text = dati
    frmCuentas.Label12.Caption = dati
    'pluto:6.7
    'Lusercuenta = 0

    'Call audio.PlayWave("Intro.wav")
End Sub

Private Sub Label2_Click()
    frmCuentas.Visible = False
    'Call frmCrearPersonaje.TirarDados
    'pluto:7.0
    frmCrearPersonaje.lbFuerza.Caption = 16
    frmCrearPersonaje.lbInteligencia.Caption = 16
    frmCrearPersonaje.lbAgilidad.Caption = 16
    frmCrearPersonaje.lbCarisma.Caption = 16
    frmCrearPersonaje.lbConstitucion.Caption = 16
    frmCrearPersonaje.lbRestantes.Caption = 6
    frmCrearPersonaje.Show vbModal

End Sub

Private Sub Label3_Click()

'pluto:2.8.0
    If Cuentas.ListIndex < 0 Then
        MsgBox "Debes seleccionar un personaje."
        Exit Sub

    End If

    frmcambiarcuenta.Show vbModal

End Sub

Private Sub Label4_Click()

'pluto:2.8.0
    If Cuentas.ListIndex < 0 Then
        MsgBox "Debes seleccionar un personaje."
        Exit Sub

    End If

    If MsgBox("Si borras este Personaje no podr�s volver a utilizarlo nunca m�s, �Estas seguro?", vbYesNo) = vbYes Then
        Call SendData("BPERSO" & Cuentas.List(Cuentas.ListIndex))
        frmCuentas.Visible = False
        frmConnect.Visible = True
        'pluto:2.11
        'DoEvents
        Sleep 1000
        Call frmMain.Socket1.Disconnect
        MsgBox "El Personaje ha sido Borrado, entra de nuevo a tu cuenta."

    End If

End Sub

Private Sub Label5_Click()
'pluto:2.14
    frmRecuperar.Show vbModal

End Sub

Private Sub Label6_Click()
    frmCuentas.Visible = False
    frmConnect.Visible = True
    Call frmMain.Socket1.Disconnect

End Sub

Private Sub Label9_Click()
    Dim dati As String
    Dim Ale As Integer
    Ale = RandomNumber(1, 118)
    dati = GetVar(App.Path & "\Init\consejos.dat", "OPCIONES", "c" & Ale)
    frmCuentas.Label12.Caption = dati

End Sub

Private Sub Text1_Change()

End Sub
