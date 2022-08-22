VERSION 5.00
Begin VB.Form frmCajas 
   BorderStyle     =   0  'None
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   Picture         =   "frmCajas.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      Height          =   525
      Left            =   2280
      MouseIcon       =   "frmCajas.frx":1B6E5
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":1C3AF
      Top             =   6240
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   5
      Left            =   3600
      MouseIcon       =   "frmCajas.frx":21471
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":2213B
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   4
      Left            =   960
      MouseIcon       =   "frmCajas.frx":28D2E
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":299F8
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   3
      Left            =   3600
      MouseIcon       =   "frmCajas.frx":305EB
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":312B5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   2
      Left            =   960
      MouseIcon       =   "frmCajas.frx":37EA8
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":38B72
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   1
      Left            =   3600
      MouseIcon       =   "frmCajas.frx":3F765
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":4042F
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1275
      Index           =   0
      Left            =   960
      MouseIcon       =   "frmCajas.frx":47022
      MousePointer    =   99  'Custom
      Picture         =   "frmCajas.frx":47CEC
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "frmCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim n As Byte

    For n = 0 To 3
        frmCajas.Image1(n).Picture = LoadPicture(DirGraficos & "baul2.jpg")
    Next

End Sub

Private Sub Image1_Click(index As Integer)
    Dim index2 As Byte
    index2 = index + 1
    SendData ("/BOVEDA" & index2)
    Unload Me

End Sub

Private Sub Image2_Click()
    Unload Me
    frmBanquero.Show vbModal

End Sub
