VERSION 5.00
Begin VB.Form Frmayuda 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Ayuda Online Aodrag"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image8 
      Height          =   300
      Left            =   2160
      Picture         =   "Frmayuda.frx":0000
      Top             =   1920
      Width           =   1680
   End
   Begin VB.Image Image7 
      Height          =   300
      Left            =   2160
      MouseIcon       =   "Frmayuda.frx":3A56
      MousePointer    =   99  'Custom
      Picture         =   "Frmayuda.frx":4720
      Top             =   4800
      Width           =   1650
   End
   Begin VB.Image Image6 
      Height          =   300
      Left            =   2160
      MouseIcon       =   "Frmayuda.frx":9974
      MousePointer    =   99  'Custom
      Picture         =   "Frmayuda.frx":A63E
      Top             =   4320
      Width           =   1650
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   2160
      MouseIcon       =   "Frmayuda.frx":F13B
      MousePointer    =   99  'Custom
      Picture         =   "Frmayuda.frx":FE05
      Top             =   3840
      Width           =   1650
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   2160
      MouseIcon       =   "Frmayuda.frx":14776
      MousePointer    =   99  'Custom
      Picture         =   "Frmayuda.frx":15440
      Top             =   3360
      Width           =   1650
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   2160
      MouseIcon       =   "Frmayuda.frx":1A0BC
      MousePointer    =   99  'Custom
      Picture         =   "Frmayuda.frx":1AD86
      Top             =   2880
      Width           =   1650
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   2160
      MouseIcon       =   "Frmayuda.frx":1FD5A
      MousePointer    =   99  'Custom
      Picture         =   "Frmayuda.frx":20A24
      Top             =   2400
      Width           =   1650
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   2280
      MouseIcon       =   "Frmayuda.frx":2551E
      Picture         =   "Frmayuda.frx":261E8
      Top             =   6360
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menú de Ayuda"
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
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1560
      Width           =   4095
   End
End
Attribute VB_Name = "Frmayuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim variable As String
Dim ie As Object

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()

End Sub

Private Sub Form_Load()
    Frmayuda.Picture = LoadPicture(App.Path & "\Graficos\ventanas.jpg")

End Sub

Private Sub Image1_Click()
    Unload Me

End Sub

Private Sub Image2_Click()
    Unload Me
    frmTeclas.Show vbModal

End Sub

Private Sub Image3_Click()
    Unload Me
    frmComandos.Show vbModal

End Sub

Private Sub Image4_Click()

    variable = "https://web.archive.org/web/20130621072803/http://www.juegosdrag.es/aomanual/?sec=clases2#ir"

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate variable
    Call AddtoRichTextBox(frmMain.RecTxt, _
                          "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 255, 255, 255, _
                          True, False, False)

End Sub

Private Sub Image5_Click()

    variable = "http://www.Worldofao.online"

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate variable
    Call AddtoRichTextBox(frmMain.RecTxt, _
                          "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 255, 255, 255, _
                          True, False, False)

End Sub

Private Sub Image6_Click()

    variable = "https://world-of-ao.fandom.com/es/wiki/World_Of_AO_Wiki"

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate variable
    Call AddtoRichTextBox(frmMain.RecTxt, _
                          "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 255, 255, 255, _
                          True, False, False)

End Sub

Private Sub Image7_Click()

    variable = "https://discord.com/invite/wPRGZEt"

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate variable
    Call AddtoRichTextBox(frmMain.RecTxt, _
                          "Web abierta en el explorer, minimiza el juego con las teclas ALT + TAB para poder ver la web.", 255, 255, 255, _
                          True, False, False)

End Sub

Private Sub Image8_Click()

Call frmConstruir.Show

End Sub
