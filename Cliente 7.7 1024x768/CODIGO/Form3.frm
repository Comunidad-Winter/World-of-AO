VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Colabora con un click en la publicidad"
   ClientHeight    =   2025
   ClientLeft      =   3645
   ClientTop       =   2835
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SHDocVwCtl.WebBrowser Navegador 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7335
      ExtentX         =   12938
      ExtentY         =   3201
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pulsa sobre la Publicidad para Continuar"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1800
      Width           =   4695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bandas_HeightChanged(ByVal NewHeight As Single)
    Form_Resize

End Sub

Sub Cargaweb2(ByVal laweb2 As String)
'Frmnavegador.Visible = True
    Navegador.Navigate laweb2
    Form3.Show vbModal

End Sub

Private Sub Form_Resize()

    If Me.WindowState = 1 Then Exit Sub

    With Navegador

        .Width = 7680
        .Height = 1600

    End With

End Sub

Private Sub Label2_Click()
    Unload Form3
    frmConnect.Show

End Sub

Private Sub Navegador_DownloadComplete()
    Static a As Byte

    'DoEvents
    a = a + 1

    If a > 2 Then
        a = 0
        Unload Form3
        frmConnect.Show

    End If

End Sub

Private Sub Navegador_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Unload Form3
    frmConnect.Show

End Sub
