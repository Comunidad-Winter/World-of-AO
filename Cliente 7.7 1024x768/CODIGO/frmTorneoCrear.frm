VERSION 5.00
Begin VB.Form frmTorneoCrear 
   BorderStyle     =   0  'None
   Caption         =   "Crear Torneo AoDraG"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   4320
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Todos Vs Todos"
         Height          =   315
         Left            =   0
         Picture         =   "frmTorneoCrear.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Eliminatoria"
         DisabledPicture =   "frmTorneoCrear.frx":0929
         Height          =   315
         Left            =   0
         Picture         =   "frmTorneoCrear.frx":5A3CB
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4320
      Picture         =   "frmTorneoCrear.frx":5ACF4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Restringir Nivel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1560
      Picture         =   "frmTorneoCrear.frx":5B9A9
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3060
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ocho Participantes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1560
      Picture         =   "frmTorneoCrear.frx":5C2D2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1740
      Width           =   2655
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Uno Contra Uno"
      DisabledPicture =   "frmTorneoCrear.frx":5CBFB
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1560
      Picture         =   "frmTorneoCrear.frx":B669D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Value           =   -1  'True
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Crear Torneo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      Picture         =   "frmTorneoCrear.frx":B6FC6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "0"
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   4320
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Text            =   "15"
         Top             =   120
         Width           =   495
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "M�nimo"
         DisabledPicture =   "frmTorneoCrear.frx":B7C7B
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         Picture         =   "frmTorneoCrear.frx":11171D
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "M�ximo"
         Height          =   315
         Left            =   0
         Picture         =   "frmTorneoCrear.frx":112046
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4320
      TabIndex        =   18
      Top             =   2400
      Width           =   1695
      Begin VB.OptionButton Option7 
         Caption         =   "Apostar Personaje"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         Picture         =   "frmTorneoCrear.frx":11296F
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   1695
      End
      Begin VB.OptionButton Option8 
         Caption         =   "15 Min. C�rcel"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         Picture         =   "frmTorneoCrear.frx":113298
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1080
      Picture         =   "frmTorneoCrear.frx":113BC1
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1080
      Picture         =   "frmTorneoCrear.frx":114293
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Premio del Ganador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   13
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cuota Inscripci�n"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Torneo AoDraG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Elige las Opciones y crea un Torneo para AoDraG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   5535
   End
End
Attribute VB_Name = "frmTorneoCrear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()

    If Check1.value = 1 Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False
        Option4.value = 1
        Option3.value = 0
        Text2.Text = 15

    End If

End Sub

Private Sub Command1_Click()

    If Val(Text2.Text) < 1 Then
        Beep
        Label6.Caption = "Nivel M�ximo o M�nimo no v�lido"
        Exit Sub

    End If

    If Val(Text2.Text) < Val(frmMain.LvlLbl.Caption) And Option3.value = True Then
        Beep
        Label6.Caption = "Nivel M�ximo no puede ser inferior al tuyo."
        Exit Sub

    End If

    If Val(Text2.Text) > 45 And Option4.value = True Then
        Beep
        Label6.Caption = "Nivel M�nimo no v�lido (Tope:45)"
        Exit Sub

    End If

    'organizar torneo
    Dim Ttip, Tcua, Tpj, Tmax, Tmin As Byte
    Dim Tins As Long

    frmTorneoParticipar.List1.Clear

    Tpj = 0
    Tins = 0
    Tmax = 0
    Tmin = 0
    Tcua = 0
    Tins = Val(Text1.Text)

    If Option1.value = True Then Ttip = 2 Else Ttip = 1
    If Option1.value = True And Option6.value = True Then Tcua = 1
    If Option1.value = True And Option5.value = True Then Tcua = 2
    If Option2.value = True And Option7.value = True Then Tpj = 1: Tins = 0
    If Option2.value = True And Option8.value = True Then Tpj = 2: Tins = 0

    If Check1.value = 1 And Option3.value = True Then Tmax = Val(Text2.Text)
    If Check1.value = 1 And Option4.value = True Then Tmin = Val(Text2.Text)
    frmTorneoCrear.Visible = False
    frmTorneoParticipar.Visible = True
    frmTorneoParticipar.Label3.Caption = UserName
    frmTorneoParticipar.Label5.Caption = Tins

    If Ttip = 2 Then frmTorneoParticipar.Label7.Caption = Val(Tins * 8) Else frmTorneoParticipar.Label7.Caption = Val( _
       Tins)
    If Tpj = 1 Then frmTorneoParticipar.Label7.Caption = "Personaje": frmTorneoParticipar.Label7.ForeColor = vbRed
    If Tpj = 2 Then frmTorneoParticipar.Label7.Caption = "15 de C�rcel": frmTorneoParticipar.Label7.ForeColor = vbRed
    If Ttip = 1 Then frmTorneoParticipar.Label9.Caption = "1 vs 1": frmTorneoParticipar.Label7.Caption = Val(Tins)
    If Tcua = 2 Then frmTorneoParticipar.Label9.Caption = "Eliminatoria"
    If Tcua = 1 Then frmTorneoParticipar.Label9.Caption = "Todos vs"
    If Tmin > 0 Then frmTorneoParticipar.Label11.Caption = "M�n.Level " & Tmin
    If Tmax > 0 Then frmTorneoParticipar.Label11.Caption = "M�x.Level " & Tmax
    If Tmin = 0 And Tmax = 0 Then frmTorneoParticipar.Label11.Caption = "Ninguna"
    SendData ("TO2" & UserName & "," & Ttip & "," & Tcua & "," & Tpj & "," & Tmax & "," & Tmin & "," & Tins)

End Sub

Private Sub Command2_Click()
    frmTorneoCrear.Visible = False
    SendData "/DTT"

End Sub

Private Sub Option1_Click()
    Image2.Visible = True
    Image1.Visible = False
    Label3.Caption = Val(Text1.Text) * 8
    Frame2.Visible = True
    Option5.Visible = True
    Option6.Visible = True
    Option6.value = True
    Option7.value = False
    Option7.Visible = False
    Option8.value = False
    Option8.Visible = False
    Frame3.Visible = False

End Sub

Private Sub Option2_Click()
    Image1.Visible = True
    Image2.Visible = False
    Label3.Caption = Val(Text1.Text)
    Frame2.Visible = False
    Option5.Visible = False
    Option6.Visible = False
    Option6.value = True
    Option7.value = False
    Option7.Visible = True
    Option8.value = False
    Option8.Visible = True
    Frame3.Visible = True
    Text1.Visible = True

End Sub

Private Sub Option7_Click()

    Label3.Caption = "El Pj de su Rival"
    Text1.Visible = False

End Sub

Private Sub Option8_Click()

    Label3.Caption = "15 de C�rcel"
    Text1.Visible = False

End Sub

Private Sub Text1_Change()

    If Option1.value = True Then Label3.Caption = Val(Text1.Text) * 8 Else Label3.Caption = Val(Text1.Text)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        If (index <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0

        End If

    End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        If (index <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0

        End If

    End If

End Sub
