VERSION 5.00
Begin VB.Form frmBanquero 
   BorderStyle     =   0  'None
   Caption         =   "Finanzas de AodraG"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form4"
   ScaleHeight     =   7185
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image7 
      Height          =   300
      Left            =   2040
      MouseIcon       =   "frmBanquero.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmBanquero.frx":0CCA
      Top             =   3960
      Width           =   2220
   End
   Begin VB.Image Image6 
      Height          =   300
      Left            =   2040
      MouseIcon       =   "frmBanquero.frx":490D
      MousePointer    =   99  'Custom
      Picture         =   "frmBanquero.frx":55D7
      Top             =   3360
      Width           =   2220
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   2040
      MouseIcon       =   "frmBanquero.frx":AF03
      MousePointer    =   99  'Custom
      Picture         =   "frmBanquero.frx":BBCD
      Top             =   2760
      Width           =   2220
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   2040
      MouseIcon       =   "frmBanquero.frx":F578
      MousePointer    =   99  'Custom
      Picture         =   "frmBanquero.frx":10242
      Top             =   2160
      Width           =   2220
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   3480
      MouseIcon       =   "frmBanquero.frx":158BA
      MousePointer    =   99  'Custom
      Picture         =   "frmBanquero.frx":16584
      Top             =   6360
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   1080
      MouseIcon       =   "frmBanquero.frx":1B646
      MousePointer    =   99  'Custom
      Picture         =   "frmBanquero.frx":1C310
      Top             =   6360
      Width           =   1485
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      MouseIcon       =   "frmBanquero.frx":2141F
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dispones de:"
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
      Height          =   240
      Left            =   720
      MouseIcon       =   "frmBanquero.frx":220E9
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finanzas WOAO"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2145
   End
End
Attribute VB_Name = "frmBanquero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Depos As Byte

Private Sub Form_Load()
'Cargamos la interfase
    Me.Picture = LoadPicture(App.Path & "\Graficos\ventanas.jpg")
    Call audio.PlayWave("comerciante3.wav")

End Sub

Private Sub Image1_Click()
    Call audio.PlayWave(SND_CLICK)
    'If Text1.Text = "" Then
    'MsgBox ("Introduce la Cantidad")
    'Exit Sub
    'End If

    If Depos = 1 Then
        SendData ("/DEPOSITAR " & Val(Text1.Text))
    ElseIf Depos = 2 Then
        SendData ("/RETIRAR " & Val(Text1.Text))

    End If

    Unload Me

End Sub

Private Sub Image2_Click()
    Call audio.PlayWave(SND_CLICK)
    Unload Me

End Sub

Private Sub Image3_Click()
    Label8.Visible = True
    Text1.Visible = True
    'Image4.Visible = True
    Depos = 1
    Call audio.PlayWave(SND_CLICK)

End Sub

Private Sub Image5_Click()
    Label8.Visible = True
    Text1.Visible = True
    'Image4.Visible = True
    Depos = 2
    Call audio.PlayWave(SND_CLICK)

End Sub

Private Sub Image6_Click()
    Call audio.PlayWave(SND_CLICK)
    'Call SendData("/BOVEDA")

    Unload Me
    frmCajas.Show vbModal

End Sub

Private Sub Image7_Click()
'Call SendData("/BOVEDA")
    Call audio.PlayWave(SND_CLICK)
    Unload Me
    frmCajas.Show vbModal

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        If (index <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0

        End If

    End If

End Sub
