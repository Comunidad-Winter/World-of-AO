VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmnavegador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servidor AodraG - www.juegosdrag.es"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar Estado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5550
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin ComCtl3.CoolBar Bandas 
      Align           =   1  'Align Top
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   2143
      BandCount       =   2
      _CBWidth        =   7230
      _CBHeight       =   1215
      _Version        =   "6.0.8169"
      Child1          =   "BarraEstandar"
      MinHeight1      =   810
      Width1          =   975
      NewRow1         =   0   'False
      Caption2        =   "Direccion"
      Child2          =   "Direccion"
      MinHeight2      =   315
      Width2          =   975
      NewRow2         =   -1  'True
      Begin VB.ComboBox Direccion 
         Height          =   315
         Left            =   975
         TabIndex        =   4
         Text            =   "http://www.juegosdrag.es"
         Top             =   870
         Width           =   6165
      End
      Begin MSComctlLib.ImageList Imagenes 
         Left            =   6360
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmnavegador.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmnavegador.frx":005E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmnavegador.frx":00BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmnavegador.frx":011A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmnavegador.frx":0178
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar BarraEstandar 
         Height          =   810
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1429
         ButtonWidth     =   1561
         ButtonHeight    =   1429
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "Imagenes"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Atras"
               Key             =   "ATRAS"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Adelante"
               Key             =   "ADELANTE"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Home"
               Key             =   "HOME"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Actualizar"
               Key             =   "ACTUALIZAR"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Detener"
               Key             =   "DETENER"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
   End
   Begin SHDocVwCtl.WebBrowser Navegador 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   7215
      ExtentX         =   12726
      ExtentY         =   7646
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
End
Attribute VB_Name = "Frmnavegador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bandas_HeightChanged(ByVal NewHeight As Single)
    Form_Resize

End Sub

Private Sub BarraEstandar_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next

    Select Case Button.key

        Case "ATRAS": Navegador.GoBack

        Case "ADELANTE": Navegador.GoForward

        Case "DETENER": Navegador.Stop

        Case "HOME": Navegador.GoHome

        Case "ACTUALIZAR": Navegador.Refresh

    End Select

End Sub

Private Sub Direccion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Navegador.Navigate Direccion.Text

    End If

End Sub

Sub Cargaweb(ByVal laweb As String)
    'Frmnavegador.Visible = True

    Navegador.Navigate laweb
    Frmnavegador.Show vbModal

End Sub

Private Sub Form_Resize()

    If Me.WindowState = 1 Then Exit Sub

    With Navegador
        .Top = Bandas.Height + 30
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - (Bandas.Height + Estado.Height + 30)

    End With

End Sub

Private Sub Navegador_DownloadComplete()
    Direccion.Text = Navegador.LocationURL

End Sub

Private Sub Navegador_StatusTextChange(ByVal Text As String)

    'Estado.Panels("ESTADO").Text = Text
End Sub

Private Sub Navegador_TitleChange(ByVal Text As String)
    Me.Caption = Text & " - Servidor AodraG v.5.1"

End Sub
