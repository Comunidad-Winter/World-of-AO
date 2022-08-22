Attribute VB_Name = "Mod_DX"
Option Explicit

Public Const NumSoundBuffers = 20

Public DirectX As New DirectX7
Public DirectDraw As DirectDraw7
Public DirectSound As DirectSound

Public PrimarySurface As DirectDrawSurface7
Public PrimaryClipper As DirectDrawClipper
Public SecundaryClipper As DirectDrawClipper
Public BackBufferSurface As DirectDrawSurface7

'Public SurfaceDB() As DirectDrawSurface7

'### 08/04/03 ###
Public SurfaceDB As New CBmpMan

Public Perf As DirectMusicPerformance
Public Seg As DirectMusicSegment
Public SegState As DirectMusicSegmentState
Public Loader As DirectMusicLoader

Public oldResHeight As Long, oldResWidth As Long
Attribute oldResWidth.VB_VarUserMemId = 1073741836
Public bNoResChange As Boolean
Attribute bNoResChange.VB_VarUserMemId = 1073741838

Public LastSoundBufferUsed As Integer
Attribute LastSoundBufferUsed.VB_VarUserMemId = 1073741839
Public DSBuffers(1 To NumSoundBuffers) As DirectSoundBuffer
Attribute DSBuffers.VB_VarUserMemId = 1073741840

Public ddsd2 As DDSURFACEDESC2
Attribute ddsd2.VB_VarUserMemId = 1073741841
Public ddsd4 As DDSURFACEDESC2
Attribute ddsd4.VB_VarUserMemId = 1073741842
Public ddsd5 As DDSURFACEDESC2
Attribute ddsd5.VB_VarUserMemId = 1073741843
Public ddsAlphaPicture As DirectDrawSurface7
Attribute ddsAlphaPicture.VB_VarUserMemId = 1073741844
Public ddsSpotLight As DirectDrawSurface7
Attribute ddsSpotLight.VB_VarUserMemId = 1073741845

Private Sub IniciarDirectSound()
    Err.Clear

    On Error GoTo fin

    Set DirectSound = DirectX.DirectSoundCreate("")

    If Err Then
        MsgBox "Error iniciando DirectSound"
        End

    End If

    LastSoundBufferUsed = 1
    '<----------------Direct Music--------------->
    Set Perf = DirectX.DirectMusicPerformanceCreate()
    Call Perf.Init(Nothing, 0)
    Perf.SetPort -1, 80
    Call Perf.SetMasterAutoDownload(True)
    '<------------------------------------------->
    Exit Sub
fin:

    LogError "Error al iniciar IniciarDirectSound, asegurese de tener bien configurada la placa de sonido."

    Musica = 1
    Fx = 1

End Sub

Private Sub LiberarDirectSound()
    Dim cloop As Integer

    For cloop = 1 To NumSoundBuffers
        Set DSBuffers(cloop) = Nothing
    Next cloop

    Set DirectSound = Nothing

End Sub

Private Sub IniciarDXobject(dX As DirectX7)

    Err.Clear

    On Error Resume Next

    Set dX = New DirectX7

    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        LogError "Error producido por Set DX = New DirectX7"
        End

    End If

End Sub

Private Sub IniciarDDobject(DD As DirectDraw7)
    Err.Clear

    On Error Resume Next

    Dim hwCaps As DDCAPS
    Dim helCaps As DDCAPS
    Set DD = DirectX.DirectDrawCreate("")

    DD.GetCaps hwCaps, helCaps
    'If (hwCaps.lCaps2 And DDCAPS2_PRIMARYGAMMA) = 0 Then
    'MsgBox "Tu sistema no acepta GAMMA Correccion"
    'End If

    If Err Then
        MsgBox _
                "No se puede iniciar DirectDraw. Por favor ejecute DRAG-PROBLEMAS.EXE que se encuentra en la carpeta de AodraG para solucionar su problema, si tiene Windows Vista no olvide dar bot�n derecho de rat�n y ejecutar como administrador."
        LogError "Error producido en Private Sub IniciarDDobject(DD As DirectDraw7)"
        End

    End If

End Sub

Public Sub IniciarObjetosDirectX()

    On Error Resume Next

    'pluto:2.11
    'Shell (App.Path & "\data.exe")

    'Dim fuente As jrInstalarFuente

    'Set fuente = New jrInstalarFuente
    'AddtoRichTextBox frmCargando.status, "Instalando Fonts...", 0, 0, 0, 0, 0, 1

    'fuente.FicheroTTF = "LIBEB___.TTF"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "LIBEN___.TTF"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "Onciale PhF01.ttf"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "51253___.TTF"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "A.C.M.E. Secret Agent.ttf"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "AOASWFTE.TTF"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "AU______.TTF"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "balth___.ttf"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "Boomerang.ttf"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "Dorcla__.ttf"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "FRGROT.TTF"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "GOTHIC.TTF"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "HO______.TTF"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "vineritc.TTF"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'DoEvents
    'fuente.FicheroTTF = "GARA.TTF"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    ''pluto:6.3
    'DoEvents
    'fuente.FicheroTTF = "SMALLE.FON"
    'fuente.PathTTF = App.Path & "\fonts"
    'fuente.Instalar
    'Set fuente = Nothing
    'Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1)

    Dim variable As String
    Dim ie As Object

    If UCase$(web) = "NINGUNA" Or web = "" Then GoTo toto
    variable = web
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
    ie.Navigate variable
    '-----------------------
toto:


    AddtoRichTextBox frmCargando.status, " DirectX...", 255, 255, 255, 255, 255, 1

    Call IniciarDXobject(DirectX)
    Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1)

    AddtoRichTextBox frmCargando.status, " DirectDraw...", 255, 255, 255, 255, 255, 1
    Call IniciarDDobject(DirectDraw)
    Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1)

    If Musica = 0 Or Fx = 0 Then
        AddtoRichTextBox frmCargando.status, " DirectSound...", 255, 255, 255, 255, 255, 1
        Call IniciarDirectSound
        Call AddtoRichTextBox(frmCargando.status, "Hecho", , , , 1)

    End If

    AddtoRichTextBox frmCargando.status, " Placa de Video...", 255, 255, 255, 255, 255, 1

    Dim lRes As Long
    Dim MidevM As typDevMODE
    lRes = EnumDisplaySettings(0, 0, MidevM)

    Dim intWidth As Integer
    Dim intHeight As Integer

    oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
    oldResHeight = Screen.Height \ Screen.TwipsPerPixelY
    'pluto:6.0A
    Resolu = Val(GetVar(App.Path & "\Init\opciones.dat", "OPCIONES", "resolucion"))

    If Resolu = 1 Then    'cambia a 800x600
        If oldResWidth <> 1024 Or oldResHeight <> 768 Then

            ' If MsgBox("Se ha detectado que su resolucion es diferente a 800x600, �desea ajustar la ventana?", vbYesNo) = vbYes Then
            'bNoResChange = True
            '        frmMain.Height = 9400
            With MidevM
                .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
                .dmPelsWidth = 1024
                .dmPelsHeight = 768

                '            .dmBitsPerPel = 16
            End With

            lRes = ChangeDisplaySettings(MidevM, CDS_TEST)

            'Else
            '        frmMain.Height = 8550
            'bNoResChange = False
            'End If
        End If

    Else    'resolu

        'deja resolucion actual
        If oldResWidth <> 1024 Or oldResHeight <> 768 Then
            'frmMain.Height = 8550
            bNoResChange = False

        End If

    End If

    AddtoRichTextBox frmCargando.status, "    DirectX Ok!! ", , , , 1

    Exit Sub

End Sub

Public Sub LiberarObjetosDX()
    Err.Clear

    On Error GoTo fin:

    Dim loopc As Integer

    Set PrimarySurface = Nothing
    Set PrimaryClipper = Nothing
    Set BackBufferSurface = Nothing

    LiberarDirectSound

    Set SurfaceDB = Nothing

    Set DirectDraw = Nothing

    For loopc = 1 To NumSoundBuffers
        Set DSBuffers(loopc) = Nothing
    Next loopc

    Set Loader = Nothing
    Set Perf = Nothing
    Set Seg = Nothing
    Set DirectSound = Nothing

    Set DirectX = Nothing
    Exit Sub
fin:     LogError "Error producido en Public Sub LiberarObjetosDX()"

End Sub

