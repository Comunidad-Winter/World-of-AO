VERSION 5.00
Begin VB.Form frmPanelGm 
   BackColor       =   &H00000000&
   Caption         =   "Panel de Gm"
   ClientHeight    =   8250
   ClientLeft      =   7575
   ClientTop       =   435
   ClientWidth     =   3975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   8250
   ScaleWidth      =   3975
   Begin VB.Frame Frame 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7575
      Index           =   3
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton cmdHABLAR 
         Caption         =   "Hablar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         MouseIcon       =   "frmPanelGm.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdRMSG2 
         Caption         =   "Mapa"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         MouseIcon       =   "frmPanelGm.frx":0CCA
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdDEV 
         Caption         =   "/DEV"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         MouseIcon       =   "frmPanelGm.frx":1994
         MousePointer    =   99  'Custom
         TabIndex        =   38
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox cboListaUsus 
         Height          =   285
         Left            =   1440
         TabIndex        =   36
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton ONLINE 
         Caption         =   "/ONLINE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         MouseIcon       =   "frmPanelGm.frx":265E
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdINVISIBLE 
         Caption         =   "/INVISIBLE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         MouseIcon       =   "frmPanelGm.frx":3328
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdDONDE 
         Caption         =   "/DONDE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         MouseIcon       =   "frmPanelGm.frx":3FF2
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdTELEP 
         Caption         =   "/TELEP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         MouseIcon       =   "frmPanelGm.frx":4CBC
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdIRA 
         Caption         =   "/IRA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         MouseIcon       =   "frmPanelGm.frx":5986
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdCARCEL 
         Caption         =   "/CARCEL"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         MouseIcon       =   "frmPanelGm.frx":6650
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdINFO 
         Caption         =   "/INFO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2640
         MouseIcon       =   "frmPanelGm.frx":731A
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdSTAT 
         Caption         =   "/FICHA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2640
         MouseIcon       =   "frmPanelGm.frx":7FE4
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdINV 
         Caption         =   "/INV"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2640
         MouseIcon       =   "frmPanelGm.frx":8CAE
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdSKILLS 
         Caption         =   "/SKILLS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2640
         MouseIcon       =   "frmPanelGm.frx":9978
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdREVIVIR 
         Caption         =   "/REVIVIR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         MouseIcon       =   "frmPanelGm.frx":A642
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdECHAR 
         Caption         =   "/ECHAR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         MouseIcon       =   "frmPanelGm.frx":B30C
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdBAN 
         Caption         =   "/BAN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         MouseIcon       =   "frmPanelGm.frx":BFD6
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdUNBAN 
         Caption         =   "/UNBAN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         MouseIcon       =   "frmPanelGm.frx":CCA0
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSUM 
         Caption         =   "/SUM"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         MouseIcon       =   "frmPanelGm.frx":D96A
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdBORRARPENA 
         Caption         =   "/UNCARCEL"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         MouseIcon       =   "frmPanelGm.frx":E634
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdLASTIP 
         Caption         =   "/IPNICK"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2640
         MouseIcon       =   "frmPanelGm.frx":F2FE
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdCC 
         Caption         =   "/CC"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         MouseIcon       =   "frmPanelGm.frx":FFC8
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdCT 
         Caption         =   "CreaTeleport"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         MouseIcon       =   "frmPanelGm.frx":10C92
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   6480
         Width           =   1335
      End
      Begin VB.CommandButton cmdDT 
         Caption         =   "DestruyeTeleport"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         MouseIcon       =   "frmPanelGm.frx":1195C
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton cmdLLUVIA 
         Caption         =   "/LLUVIA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         MouseIcon       =   "frmPanelGm.frx":12626
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CommandButton cmdMASSDEST 
         Caption         =   "Limpiar Objetos del Mapa"
         CausesValidation=   0   'False
         Height          =   915
         Left            =   1440
         MouseIcon       =   "frmPanelGm.frx":132F0
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   6360
         Width           =   735
      End
      Begin VB.CommandButton cmdCI 
         Caption         =   "Crear Objeto"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         MouseIcon       =   "frmPanelGm.frx":13FBA
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   6480
         Width           =   1335
      End
      Begin VB.CommandButton cmdDEST 
         Caption         =   "/DEST"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         MouseIcon       =   "frmPanelGm.frx":14C84
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton cmdNENE 
         Caption         =   "/NENE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         MouseIcon       =   "frmPanelGm.frx":1594E
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdONLINEGM 
         Caption         =   "/ONLINEGM"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         MouseIcon       =   "frmPanelGm.frx":16618
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   5760
         Width           =   1215
      End
      Begin VB.CommandButton cmdONLINEMAP 
         Caption         =   "/ONLINEMAP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         MouseIcon       =   "frmPanelGm.frx":172E2
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdBORRAR_SOS 
         Caption         =   "/BORRAR SOS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         MouseIcon       =   "frmPanelGm.frx":17FAC
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdSHOW_SOS 
         Caption         =   "/SHOW SOS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         MouseIcon       =   "frmPanelGm.frx":18C76
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdRMSG 
         Caption         =   "General"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         MouseIcon       =   "frmPanelGm.frx":19940
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdSMSG 
         Caption         =   "Cartel"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         MouseIcon       =   "frmPanelGm.frx":1A60A
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdHORA 
         Caption         =   "/HORA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         MouseIcon       =   "frmPanelGm.frx":1B2D4
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CommandButton cmdGMSG 
         Caption         =   "Clan"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         MouseIcon       =   "frmPanelGm.frx":1BF9E
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Mensajes"
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
         Left            =   120
         TabIndex        =   39
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.Line Line6 
         BorderColor     =   &H0000FFFF&
         X1              =   0
         X2              =   3960
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0000FFFF&
         X1              =   0
         X2              =   3960
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line4 
         BorderColor     =   &H0000FFFF&
         X1              =   0
         X2              =   3960
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000FFFF&
         X1              =   0
         X2              =   3960
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FFFF&
         X1              =   120
         X2              =   3960
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         X1              =   0
         X2              =   3960
         Y1              =   2880
         Y2              =   2880
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   1200
      MouseIcon       =   "frmPanelGm.frx":1CC68
      TabIndex        =   0
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "World of AO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   35
      Top             =   7560
      Width           =   3135
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cboListaUsus_Validate(Cancel As Boolean)
    Cancel = True

End Sub

Private Sub cmdBAN_Click()
'/BAN
    Dim tStr As String
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then
        tStr = InputBox("Escriba el motivo del ban.", "BAN a " & nick)

        If LenB(tStr) <> 0 Then If MsgBox("�Seguro desea banear a " & nick & "?", vbYesNo, "Atencion!") = vbYes Then _
           SendData "/BAN " & tStr & "@" & nick

    End If

End Sub

Private Sub cmdBORRAR_SOS_Click()

'/BORRAR SOS
    If MsgBox("�Seguro desea borrar el SOS?", vbYesNo, "Atencion!") = vbYes Then Call SendData("/BORRAR SOS")

End Sub

Private Sub cmdBORRARPENA_Click()
'/BORRARPENA
    Dim tStr As String
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then

        If MsgBox("�Seguro desea borrar la pena " & tStr & " a " & nick & "?", vbYesNo, "Atencion!") = vbYes Then _
           SendData "/UNCARCEL " & nick

    End If

End Sub

Private Sub cmdCARCEL_Click()
'/CARCEL
    Dim tStr As String
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then
        'tStr = InputBox("Escriba el motivo de la pena.", "Carcel a " & Nick)

        tStr = InputBox("Indique el tiempo de condena (entre 0 y 60 minutos).", "Carcel a " & nick)
        'We use the Parser to control the command format
        SendData "/CARCEL " & tStr & " " & nick

    End If

End Sub

Private Sub cmdCC_Click()
'/CC
    SendData "/cc"

End Sub

Private Sub cmdCI_Click()
'/CI
    Dim tStr As String

    tStr = InputBox("Indique el n�mero del objeto a crear.", "Crear Objeto")

    If LenB(tStr) <> 0 Then If MsgBox("�Seguro desea crear el objeto " & tStr & "?", vbYesNo, "Atencion!") = vbYes _
       Then SendData "/mod " & UserName & " objetox " & tStr

End Sub

Private Sub cmdCT_Click()
'/CT
    Dim tStr As String

    tStr = InputBox("Indique la posici�n donde lleva el portal (MAPA X Y).", "Crear Portal")

    If LenB(tStr) <> 0 Then SendData "/CT " & tStr

End Sub

Private Sub cmdDEST_Click()

'/DEST
    If MsgBox("�Seguro desea destruir el objeto sobre el que esta parado?", vbYesNo, "Atencion!") = vbYes Then _
       SendData "/dest"

End Sub

Private Sub cmdDEV_Click()
'/SUM
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then SendData ("/DEV " & nick)

End Sub

Private Sub cmdDONDE_Click()
'/DONDE
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then Call SendData("/DONDE " & nick)

End Sub

Private Sub cmdDT_Click()

'DT
    If MsgBox("�Seguro desea destruir el portal?", vbYesNo, "Atencion!") = vbYes Then SendData "/DT"

End Sub

Private Sub cmdECHAR_Click()
'/ECHAR
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then SendData "/ECHAR " & nick

End Sub

Private Sub cmdGMSG_Click()
'/GMSG
    Dim tStr As String

    tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de GM")

    If LenB(tStr) <> 0 Then SendData "/CLAN " & tStr

End Sub

Private Sub cmdHABLAR_Click()

    Dim tStr As String

    tStr = InputBox("Escriba el mensaje.", "Hablar con el GM")

    If LenB(tStr) <> 0 Then Call SendData(";" & tStr)

End Sub

Private Sub cmdHORA_Click()
'/HORA
    SendData "/HORA"

End Sub

Private Sub cmdINFO_Click()
'/INFO
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then SendData "/INFO " & nick

End Sub

Private Sub cmdINV_Click()
'/INV
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then SendData "/INV " & nick

End Sub

Private Sub cmdINVISIBLE_Click()
'/INVISIBLE
    SendData "/invisible"

End Sub

Private Sub cmdIRA_Click()
'/IRA
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then Call SendData("/IRA " & nick)

End Sub

Private Sub cmdLASTIP_Click()
'/LASTIP
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then SendData "/IPNICK " & nick

End Sub

Private Sub cmdLLUVIA_Click()
'/LLUVIA
    SendData "/lluvia"

End Sub

Private Sub cmdMASSDEST_Click()

'/MASSDEST
    If MsgBox("�Seguro desea destruir todos los items del mapa?", vbYesNo, "Atencion!") = vbYes Then SendData "/limpiar"

End Sub

Private Sub cmdNENE_Click()
'/NENE
    Dim tStr As String

    tStr = InputBox("Indique el mapa.", "N�mero de NPCs enemigos.")

    If LenB(tStr) <> 0 Then SendData "/nene " & tStr

End Sub

Private Sub cmdONLINEGM_Click()
'/ONLINEGM
    SendData "/ONLINEGM"

End Sub

Private Sub cmdONLINEMAP_Click()
'/ONLINEMAP
    Dim tStr As String

    tStr = InputBox("Indique el mapa.", "N�mero de Jugadores Online.")

    If LenB(tStr) <> 0 Then SendData "/ONLINEMAPA " & tStr

End Sub

Private Sub cmdREVIVIR_Click()
'/REVIVIR
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then SendData ("/REVIVIR " & nick)

End Sub

Private Sub cmdRMSG_Click()
'/RMSG
    Dim tStr As String

    tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de RoleMaster")

    If LenB(tStr) <> 0 Then Call SendData("/RMSG " & tStr)

End Sub

Private Sub cmdRMSG2_Click()
'/RMSG
    Dim tStr As String

    tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de RoleMaster")

    If LenB(tStr) <> 0 Then Call SendData("/RMSG2 " & tStr)

End Sub

Private Sub cmdSHOW_SOS_Click()
'/SHOW SOS
    Call SendData("/SHOW SOS")

End Sub

Private Sub cmdSKILLS_Click()
'/SKILLS
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then SendData "/SKILLS " & nick

End Sub

Private Sub cmdSMSG_Click()
'/SMSG
    Dim tStr As String

    tStr = InputBox("Escriba el mensaje.", "Mensaje de sistema")

    If LenB(tStr) <> 0 Then SendData ("/SMSG " & tStr)

End Sub

Private Sub cmdSTAT_Click()
'/STAT
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then SendData "/FICHA " & nick

End Sub

Private Sub cmdSUM_Click()
'/SUM
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then SendData ("/SUM " & nick)

End Sub

Private Sub cmdTELEP_Click()
'/TELEP
    Dim tStr As String
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then
        tStr = InputBox("Indique la posici�n (MAPA X Y).", "Transportar a " & nick)

        If LenB(tStr) <> 0 Then SendData ("/TELEP " & nick & " " & tStr)

    End If

End Sub

Private Sub cmdUNBAN_Click()
'/UNBAN
    Dim nick As String

    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then If MsgBox("�Seguro desea unbanear a " & nick & "?", vbYesNo, "Atencion!") = vbYes Then _
       SendData "/UNBAN " & nick

End Sub

Private Sub Form_Load()

'Call showTab(1)
'Call cmdActualiza_Click
End Sub

'Private Sub cmdActualiza_Click()
'SendData "/ADONLINE"

'End Sub

Private Sub cmdCerrar_Click()
    Unload Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me

End Sub

Private Sub ONLINE_Click()
    SendData "/ONLINE"

End Sub
