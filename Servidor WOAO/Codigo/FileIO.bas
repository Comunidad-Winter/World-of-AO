Attribute VB_Name = "ES"
Option Explicit



Public Sub CargarPremiosList()
    Dim p As Integer, loopc As Integer
    p = val(GetVar(App.Path & "\Dat\Premios.dat", "INIT", "NumPremios"))
    'canjeo [Dylan.-]
    ReDim PremiosList(p) As tPremiosCanjes


    For loopc = 1 To p
        PremiosList(loopc).ObjName = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopc, "Nombre")
        PremiosList(loopc).ObjIndexP = val(GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopc, "NumObj"))
        PremiosList(loopc).ObjRequiere = val(GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopc, "Requiere"))
        PremiosList(loopc).ObjMaxAt = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopc, "AtaqueMaximo")
        PremiosList(loopc).ObjMinAt = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopc, "AtaqueMinimo")
        PremiosList(loopc).ObjMindef = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopc, "DefensaMinima")
        PremiosList(loopc).ObjMaxdef = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopc, "DefensaMaxima")
        PremiosList(loopc).ObjMinAtMag = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopc, "AtaqueMagicoMinimo")
        PremiosList(loopc).ObjMaxAtMag = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopc, "AtaqueMagicoMaximo")
        PremiosList(loopc).ObjMinDefMag = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopc, "DefensaMagicaMinima")
        PremiosList(loopc).ObjMaxDefMag = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopc, "DefensaMagicaMaxima")
        PremiosList(loopc).ObjDescripcion = GetVar(App.Path & "\Dat\Premios.dat", "PREMIO" & loopc, "Descripcion")
    Next loopc
End Sub


    Public Sub CargarPremiosListD()
        Dim p As Integer, loopc As Integer
        p = val(GetVar(App.Path & "\Dat\Donaciones.dat", "INIT", "NumPremios"))
   'canjeo [Dylan.-]
        ReDim PremiosListD(p) As tPremiosCanjesD
       
           
        For loopc = 1 To p
            PremiosListD(loopc).ObjName = GetVar(App.Path & "\Dat\Donaciones.dat", "PREMIO" & loopc, "Nombre")
            PremiosListD(loopc).ObjIndexP = val(GetVar(App.Path & "\Dat\Donaciones.dat", "PREMIO" & loopc, "NumObj"))
            PremiosListD(loopc).ObjRequiere = val(GetVar(App.Path & "\Dat\Donaciones.dat", "PREMIO" & loopc, "Requiere"))
            PremiosListD(loopc).ObjMaxAt = GetVar(App.Path & "\Dat\Donaciones.dat", "PREMIO" & loopc, "AtaqueMaximo")
            PremiosListD(loopc).ObjMinAt = GetVar(App.Path & "\Dat\Donaciones.dat", "PREMIO" & loopc, "AtaqueMinimo")
            PremiosListD(loopc).ObjMindef = GetVar(App.Path & "\Dat\Donaciones.dat", "PREMIO" & loopc, "DefensaMinima")
            PremiosListD(loopc).ObjMaxdef = GetVar(App.Path & "\Dat\Donaciones.dat", "PREMIO" & loopc, "DefensaMaxima")
            PremiosListD(loopc).ObjMinAtMag = GetVar(App.Path & "\Dat\Donaciones.dat", "PREMIO" & loopc, "AtaqueMagicoMinimo")
            PremiosListD(loopc).ObjMaxAtMag = GetVar(App.Path & "\Dat\Donaciones.dat", "PREMIO" & loopc, "AtaqueMagicoMaximo")
            PremiosListD(loopc).ObjMinDefMag = GetVar(App.Path & "\Dat\Donaciones.dat", "PREMIO" & loopc, "DefensaMagicaMinima")
            PremiosListD(loopc).ObjMaxDefMag = GetVar(App.Path & "\Dat\Donaciones.dat", "PREMIO" & loopc, "DefensaMagicaMaxima")
            PremiosListD(loopc).ObjDescripcion = GetVar(App.Path & "\Dat\Donaciones.dat", "PREMIO" & loopc, "Descripcion")
            PremiosListD(loopc).ObjFoto = GetVar(App.Path & "\Dat\Donaciones.dat", "PREMIO" & loopc, "Foto")
        Next loopc
    End Sub

Public Sub CargarSpawnList()

    On Error GoTo fallo

    Dim n As Integer, loopc As Integer
    n = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(n) As tCriaturasEntrenador

    For loopc = 1 To n
        SpawnList(loopc).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & loopc))
        SpawnList(loopc).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & loopc)
    Next loopc

    Exit Sub
fallo:
    Call LogError("CARGARSPAWNLIST" & Err.number & " D: " & Err.Description)

End Sub

Function EsDios(ByVal Name As String) As Boolean

    On Error GoTo fallo

    Dim NumWizs As Integer
    Dim WizNum As Integer
    NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))

    For WizNum = 1 To NumWizs

        If UCase$(Name) = UCase$(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum)) Then
            EsDios = True
            Exit Function

        End If

    Next WizNum

    EsDios = False

    Exit Function
fallo:
    Call LogError("ESDIOS" & Err.number & " D: " & Err.Description)

End Function

Function EsSemiDios(ByVal Name As String) As Boolean

    On Error GoTo fallo

    Dim NumWizs As Integer
    Dim WizNum As Integer
    NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))

    For WizNum = 1 To NumWizs

        If UCase$(Name) = UCase$(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum)) Then
            EsSemiDios = True
            Exit Function

        End If

    Next WizNum

    EsSemiDios = False

    Exit Function
fallo:
    Call LogError("ESSEMIDIOS" & Err.number & " D: " & Err.Description)

End Function

Function EsConsejero(ByVal Name As String) As Boolean

    On Error GoTo fallo

    Dim NumWizs As Integer
    Dim WizNum As Integer
    NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))

    For WizNum = 1 To NumWizs

        If UCase$(Name) = UCase$(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum)) Then
            EsConsejero = True
            Exit Function

        End If

    Next WizNum

    EsConsejero = False

    Exit Function
fallo:
    Call LogError("ESCONSEJERO" & Err.number & " D: " & Err.Description)

End Function

Public Function TxtDimension(ByVal Name As String) As Long

    On Error GoTo fallo

    Dim n As Integer, cad As String, Tam As Long
    n = FreeFile(1)
    Open Name For Input As #n
    Tam = 0

    Do While Not EOF(n)
        Tam = Tam + 1
        Line Input #n, cad
    Loop
    Close n
    TxtDimension = Tam

    Exit Function
fallo:
    Call LogError("TXTDIMENSION" & Err.number & " D: " & Err.Description)

End Function

Public Sub CargarForbidenWords()

    On Error GoTo fallo

    ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
    Dim n As Integer, i As Integer
    n = FreeFile(1)
    Open DatPath & "NombresInvalidos.txt" For Input As #n

    For i = 1 To UBound(ForbidenNames)
        Line Input #n, ForbidenNames(i)
    Next i

    Close n
    Exit Sub
fallo:
    Call LogError("CAGARFORBIDENWORDS" & Err.number & " D: " & Err.Description)

End Sub

Public Sub CargarHechizos()

    On Error GoTo errhandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."

    Dim Hechizo As Integer

    'pluto fusión
    Dim leer As New clsIniManager
    leer.Initialize DatPath & "Hechizos.dat"

    'obtiene el numero de hechizos
    NumeroHechizos = val(leer.GetValue("INIT", "NumeroHechizos"))
    'NumeroHechizos = val(GetVar(DatPath & "Hechizos.dat", "INIT", "NumeroHechizos"))
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo

    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumeroHechizos
    frmCargando.cargar.value = 0

    'Llena la lista
    For Hechizo = 1 To NumeroHechizos
        frmCargando.Label1(2).Caption = "Hechizo: (" & Hechizo & "/" & NumeroHechizos & ")"

        Hechizos(Hechizo).Nombre = leer.GetValue("Hechizo" & Hechizo, "Nombre")
        Hechizos(Hechizo).Desc = leer.GetValue("Hechizo" & Hechizo, "Desc")
        Hechizos(Hechizo).PalabrasMagicas = leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")

        Hechizos(Hechizo).HechizeroMsg = leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
        Hechizos(Hechizo).TargetMsg = leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
        Hechizos(Hechizo).PropioMsg = leer.GetValue("Hechizo" & Hechizo, "PropioMsg")

        Hechizos(Hechizo).Tipo = val(leer.GetValue("Hechizo" & Hechizo, "Tipo"))
        Hechizos(Hechizo).WAV = val(leer.GetValue("Hechizo" & Hechizo, "WAV"))
        Hechizos(Hechizo).FXgrh = val(leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))

        Hechizos(Hechizo).loops = val(leer.GetValue("Hechizo" & Hechizo, "Loops"))

        Hechizos(Hechizo).Resis = val(leer.GetValue("Hechizo" & Hechizo, "Resis"))

        Hechizos(Hechizo).SubeHP = val(leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
        Hechizos(Hechizo).MinHP = val(leer.GetValue("Hechizo" & Hechizo, "MinHP"))
        Hechizos(Hechizo).MaxHP = val(leer.GetValue("Hechizo" & Hechizo, "MaxHP"))

        Hechizos(Hechizo).SubeMana = val(leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
        Hechizos(Hechizo).MiMana = val(leer.GetValue("Hechizo" & Hechizo, "MinMana"))
        Hechizos(Hechizo).MaMana = val(leer.GetValue("Hechizo" & Hechizo, "MaxMana"))

        Hechizos(Hechizo).SubeSta = val(leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
        Hechizos(Hechizo).MinSta = val(leer.GetValue("Hechizo" & Hechizo, "MinSta"))
        Hechizos(Hechizo).MaxSta = val(leer.GetValue("Hechizo" & Hechizo, "MaxSta"))

        Hechizos(Hechizo).SubeHam = val(leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
        Hechizos(Hechizo).MinHam = val(leer.GetValue("Hechizo" & Hechizo, "MinHam"))
        Hechizos(Hechizo).MaxHam = val(leer.GetValue("Hechizo" & Hechizo, "MaxHam"))

        Hechizos(Hechizo).SubeSed = val(leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
        Hechizos(Hechizo).MinSed = val(leer.GetValue("Hechizo" & Hechizo, "MinSed"))
        Hechizos(Hechizo).MaxSed = val(leer.GetValue("Hechizo" & Hechizo, "MaxSed"))

        Hechizos(Hechizo).SubeAgilidad = val(leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
        Hechizos(Hechizo).MinAgilidad = val(leer.GetValue("Hechizo" & Hechizo, "MinAG"))
        Hechizos(Hechizo).MaxAgilidad = val(leer.GetValue("Hechizo" & Hechizo, "MaxAG"))

        Hechizos(Hechizo).SubeFuerza = val(leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
        Hechizos(Hechizo).MinFuerza = val(leer.GetValue("Hechizo" & Hechizo, "MinFU"))
        Hechizos(Hechizo).MaxFuerza = val(leer.GetValue("Hechizo" & Hechizo, "MaxFU"))

        Hechizos(Hechizo).SubeCarisma = val(leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
        Hechizos(Hechizo).MinCarisma = val(leer.GetValue("Hechizo" & Hechizo, "MinCA"))
        Hechizos(Hechizo).MaxCarisma = val(leer.GetValue("Hechizo" & Hechizo, "MaxCA"))

        Hechizos(Hechizo).Invisibilidad = val(leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
        Hechizos(Hechizo).Paraliza = val(leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
        Hechizos(Hechizo).Paralizaarea = val(leer.GetValue("Hechizo" & Hechizo, "Paralizaarea"))

        'Hechizos(Hechizo).Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
        Hechizos(Hechizo).RemoverParalisis = val(leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
        'Hechizos(Hechizo).RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
        Hechizos(Hechizo).RemueveInvisibilidadParcial = val(leer.GetValue("Hechizo" & Hechizo, _
                                                                          "RemueveInvisibilidadParcial"))

        Hechizos(Hechizo).CuraVeneno = val(leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
        Hechizos(Hechizo).Envenena = val(leer.GetValue("Hechizo" & Hechizo, "Envenena"))
        'pluto:2.15
        Hechizos(Hechizo).Protec = val(leer.GetValue("Hechizo" & Hechizo, "Protec"))

        Hechizos(Hechizo).Maldicion = val(leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
        Hechizos(Hechizo).RemoverMaldicion = val(leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
        Hechizos(Hechizo).Bendicion = val(leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
        Hechizos(Hechizo).Revivir = val(leer.GetValue("Hechizo" & Hechizo, "Revivir"))
        Hechizos(Hechizo).Morph = val(leer.GetValue("Hechizo" & Hechizo, "Morph"))

        Hechizos(Hechizo).Ceguera = val(leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
        Hechizos(Hechizo).Estupidez = val(leer.GetValue("Hechizo" & Hechizo, "Estupidez"))

        Hechizos(Hechizo).invoca = val(leer.GetValue("Hechizo" & Hechizo, "Invoca"))
        Hechizos(Hechizo).NumNpc = val(leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
        Hechizos(Hechizo).Cant = val(leer.GetValue("Hechizo" & Hechizo, "Cant"))
        'Hechizos(Hechizo).Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))

        Hechizos(Hechizo).MinNivel = val(leer.GetValue("Hechizo" & Hechizo, "MinNivel"))
        Hechizos(Hechizo).itemIndex = val(leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))

        Hechizos(Hechizo).MinSkill = val(leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
        Hechizos(Hechizo).ManaRequerido = val(leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
        Hechizos(Hechizo).Target = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Target"))

        frmCargando.cargar.value = frmCargando.cargar.value + 1
        'DoEvents
    Next

    'quitar esto
    Exit Sub

    '------------------------------------------------------------------------------------
    'Esto genera el hechizos.log para meterlo al cliente, el server no usa nada de lo de abajo.
    '------------------------------------------------------------------------------------
    Dim File As String
    Dim n As Byte
    Dim Object As Integer
    File = DatPath & "Hechizos.dat"
    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\Hechizo.log" For Append Shared As #nfile

    For Object = 1 To NumeroHechizos
        Debug.Print Object
        Print #nfile, "hechizos(" & Object & ").nombre=" & Chr(34) & Hechizos(Object).Nombre & Chr(34)
        Print #nfile, "hechizos(" & Object & ").desc=" & Chr(34) & Hechizos(Object).Desc & Chr(34)
        Print #nfile, "hechizos(" & Object & ").palabrasmagicas=" & Chr(34) & Hechizos(Object).PalabrasMagicas & Chr(34)
        Print #nfile, "hechizos(" & Object & ").hechizeromsg=" & Chr(34) & Hechizos(Object).HechizeroMsg & Chr(34)
        Print #nfile, "hechizos(" & Object & ").propiomsg=" & Chr(34) & Hechizos(Object).PropioMsg & Chr(34)
        Print #nfile, "hechizos(" & Object & ").targetmsg=" & Chr(34) & Hechizos(Object).TargetMsg & Chr(34)

        If Hechizos(Object).Bendicion > 0 Then Print #nfile, "hechizos(" & Object & ").bendicion =" & Hechizos( _
                                                             Object).Bendicion

        If Hechizos(Object).Cant > 0 Then Print #nfile, "hechizos(" & Object & ").cant =" & Hechizos(Object).Cant
        If Hechizos(Object).Ceguera > 0 Then Print #nfile, "hechizos(" & Object & ").ceguera =" & Hechizos( _
                                                           Object).Ceguera

        If Hechizos(Object).CuraVeneno > 0 Then Print #nfile, "hechizos(" & Object & ").curaveneno =" & Hechizos( _
                                                              Object).CuraVeneno

        If Hechizos(Object).Envenena > 0 Then Print #nfile, "hechizos(" & Object & ").envenena =" & Hechizos( _
                                                            Object).Envenena

        If Hechizos(Object).Estupidez > 0 Then Print #nfile, "hechizos(" & Object & ").estupidez =" & Hechizos( _
                                                             Object).Estupidez

        If Hechizos(Object).FXgrh > 0 Then Print #nfile, "hechizos(" & Object & ").fxgrh =" & Hechizos(Object).FXgrh
        If Hechizos(Object).Invisibilidad > 0 Then Print #nfile, "hechizos(" & Object & ").invisibilidad =" & _
                                                                 Hechizos(Object).Invisibilidad

        If Hechizos(Object).invoca > 0 Then Print #nfile, "hechizos(" & Object & ").invoca =" & Hechizos(Object).invoca
        If Hechizos(Object).itemIndex > 0 Then Print #nfile, "hechizos(" & Object & ").itemindex =" & Hechizos( _
                                                             Object).itemIndex

        If Hechizos(Object).loops > 0 Then Print #nfile, "hechizos(" & Object & ").loops =" & Hechizos(Object).loops
        If Hechizos(Object).Maldicion > 0 Then Print #nfile, "hechizos(" & Object & ").maldicion  =" & Hechizos( _
                                                             Object).Maldicion

        If Hechizos(Object).MaMana > 0 Then Print #nfile, "hechizos(" & Object & ").mamana=" & Hechizos(Object).MaMana
        If Hechizos(Object).ManaRequerido > 0 Then Print #nfile, "hechizos(" & Object & ").ManaRequerido =" & _
                                                                 Hechizos(Object).ManaRequerido

        If Hechizos(Object).MaxAgilidad > 0 Then Print #nfile, "hechizos(" & Object & ").maxagilidad =" & Hechizos( _
                                                               Object).MaxAgilidad

        If Hechizos(Object).MaxCarisma > 0 Then Print #nfile, "hechizos(" & Object & ").Maxcarisma =" & Hechizos( _
                                                              Object).MaxCarisma

        If Hechizos(Object).MaxFuerza > 0 Then Print #nfile, "hechizos(" & Object & ").Maxfuerza =" & Hechizos( _
                                                             Object).MaxFuerza

        If Hechizos(Object).MaxHam > 0 Then Print #nfile, "hechizos(" & Object & ").maxham =" & Hechizos(Object).MaxHam
        If Hechizos(Object).MaxHP > 0 Then Print #nfile, "hechizos(" & Object & ").Maxhp =" & Hechizos(Object).MaxHP
        If Hechizos(Object).MaxSed > 0 Then Print #nfile, "hechizos(" & Object & ").Maxsed =" & Hechizos(Object).MaxSed
        If Hechizos(Object).MaxSta > 0 Then Print #nfile, "hechizos(" & Object & ").Maxsta =" & Hechizos(Object).MaxSta
        If Hechizos(Object).MiMana > 0 Then Print #nfile, "hechizos(" & Object & ").Mimana =" & Hechizos(Object).MiMana
        If Hechizos(Object).MinAgilidad > 0 Then Print #nfile, "hechizos(" & Object & ").minagilidad =" & Hechizos( _
                                                               Object).MinAgilidad

        If Hechizos(Object).MinCarisma > 0 Then Print #nfile, "hechizos(" & Object & ").mincarisma =" & Hechizos( _
                                                              Object).MinCarisma

        If Hechizos(Object).MinFuerza > 0 Then Print #nfile, "hechizos(" & Object & ").Minfuerza =" & Hechizos( _
                                                             Object).MinFuerza

        If Hechizos(Object).MinHam > 0 Then Print #nfile, "hechizos(" & Object & ").Minham =" & Hechizos(Object).MinHam
        If Hechizos(Object).MinHP > 0 Then Print #nfile, "hechizos(" & Object & ").Minhp =" & Hechizos(Object).MinHP
        If Hechizos(Object).MinSed > 0 Then Print #nfile, "hechizos(" & Object & ").minsed =" & Hechizos(Object).MinSed
        If Hechizos(Object).MinSkill > 0 Then Print #nfile, "hechizos(" & Object & ").minskill =" & Hechizos( _
                                                            Object).MinSkill

        If Hechizos(Object).MinSta > 0 Then Print #nfile, "hechizos(" & Object & ").Minsta =" & Hechizos(Object).MinSta
        If Hechizos(Object).Morph > 0 Then Print #nfile, "hechizos(" & Object & ").morph =" & Hechizos(Object).Morph
        If Hechizos(Object).MinNivel > 0 Then Print #nfile, "hechizos(" & Object & ").MinNivel =" & Hechizos( _
                                                            Object).MinNivel

        If Hechizos(Object).NumNpc > 0 Then Print #nfile, "hechizos(" & Object & ").numnpc =" & Hechizos(Object).NumNpc
        If Hechizos(Object).Paraliza > 0 Then Print #nfile, "hechizos(" & Object & ").paraliza =" & Hechizos( _
                                                            Object).Paraliza

        If Hechizos(Object).Paralizaarea > 0 Then Print #nfile, "hechizos(" & Object & ").paralizaarea =" & Hechizos( _
                                                                Object).Paralizaarea

        If Hechizos(Object).Protec > 0 Then Print #nfile, "hechizos(" & Object & ").protec =" & Hechizos(Object).Protec
        If Hechizos(Object).RemoverMaldicion > 0 Then Print #nfile, "hechizos(" & Object & ").removermaldicion =" & _
                                                                    Hechizos(Object).RemoverMaldicion

        If Hechizos(Object).RemoverParalisis > 0 Then Print #nfile, "hechizos(" & Object & ").removerparalisis =" & _
                                                                    Hechizos(Object).RemoverParalisis

        If Hechizos(Object).Resis > 0 Then Print #nfile, "hechizos(" & Object & ").resis =" & Hechizos(Object).Resis
        If Hechizos(Object).Revivir > 0 Then Print #nfile, "hechizos(" & Object & ").revivir =" & Hechizos( _
                                                           Object).Revivir

        If Hechizos(Object).SubeAgilidad > 0 Then Print #nfile, "hechizos(" & Object & ").subeagilidad =" & Hechizos( _
                                                                Object).SubeAgilidad

        If Hechizos(Object).SubeCarisma > 0 Then Print #nfile, "hechizos(" & Object & ").subecarisma =" & Hechizos( _
                                                               Object).SubeCarisma

        If Hechizos(Object).SubeFuerza > 0 Then Print #nfile, "hechizos(" & Object & ").subefuerza=" & Hechizos( _
                                                              Object).SubeFuerza

        If Hechizos(Object).SubeHam > 0 Then Print #nfile, "hechizos(" & Object & ").subeham =" & Hechizos( _
                                                           Object).SubeHam

        If Hechizos(Object).SubeHP > 0 Then Print #nfile, "hechizos(" & Object & ").subehp =" & Hechizos(Object).SubeHP
        If Hechizos(Object).SubeMana > 0 Then Print #nfile, "hechizos(" & Object & ").subemana =" & Hechizos( _
                                                            Object).SubeMana

        If Hechizos(Object).SubeSed > 0 Then Print #nfile, "hechizos(" & Object & ").subesed =" & Hechizos( _
                                                           Object).SubeSed

        If Hechizos(Object).SubeSta > 0 Then Print #nfile, "hechizos(" & Object & ").subesta =" & Hechizos( _
                                                           Object).SubeSta

        If Hechizos(Object).Target > 0 Then Print #nfile, "hechizos(" & Object & ").target =" & Hechizos(Object).Target
        If Hechizos(Object).Tipo > 0 Then Print #nfile, "hechizos(" & Object & ").tipo =" & Hechizos(Object).Tipo
        If Hechizos(Object).WAV > 0 Then Print #nfile, "hechizos(" & Object & ").wav =" & Hechizos(Object).WAV

    Next
    Close #nfile
    Exit Sub
errhandler:
    MsgBox "Error cargando hechizos.dat"

End Sub

Sub LoadMotd()

    On Error GoTo fallo

    Dim i As Integer
    MaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))
    ReDim MOTD(1 To MaxLines) As String

    For i = 1 To MaxLines
        MOTD(i) = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
    Next i

    Exit Sub
fallo:
    Call LogError("LOADMOTD" & Err.number & " D: " & Err.Description)

End Sub

Public Sub DoBackUp()

'Call LogTarea("Sub DoBackUp")
    On Error GoTo fallo

    haciendoBK = True
    Call SendData2(ToAll, 0, 0, 19)

    Call SaveGuildsDB
    Call LimpiarMundo
    Call WorldSave

    Call SendData2(ToAll, 0, 0, 19)

    haciendoBK = False

    'Log

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time
    Close #nfile

    Exit Sub
fallo:
    Call LogError("DOBACKUP" & Err.number & " D: " & Err.Description)

End Sub

Public Sub grabaPJ()

    On Error GoTo fallo

    Dim Pj As Integer
    Dim Name As String
    haciendoBKPJ = True
    Call SendData(ToAll, 0, 0, "||%%%% POR FAVOR ESPERE, GRABANDO FICHAS DE PJS...%%%%" & "´" & _
                               FontTypeNames.FONTTYPE_INFO)
    Call SendData2(ToAll, 0, 0, 19)

    For Pj = 1 To LastUser
        Call SaveUser(Pj, CharPath & Left$(UCase$(UserList(Pj).Name), 1) & "\" & UCase$(UserList(Pj).Name) & ".chr")
    Next Pj

    Call SendData2(ToAll, 0, 0, 19)
    Call SendData(ToAll, 0, 0, "||%%%% FICHAS GRABADAS, PUEDEN CONTINUAR.GRACIAS. %%%%" & "´" & _
                               FontTypeNames.FONTTYPE_INFO)

    haciendoBKPJ = False

    'Log

    Dim nfile As Integer
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\logs\BackupPJ.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time
    Close #nfile

    Exit Sub
fallo:
    Call LogError("GRABAPJ" & Err.number & " D: " & Err.Description)

End Sub


Sub LoadArmasHerreria()

    On Error GoTo fallo

    Dim n As Integer, LC As Integer

    n = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

    ReDim Preserve ArmasHerrero(1 To n) As Integer

    For LC = 1 To n
        ArmasHerrero(LC) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & LC, "Index"))
        'pluto:6.0a
        ObjData(ArmasHerrero(LC)).ParaHerre = 1
    Next LC

    Exit Sub
fallo:
    Call LogError("LOADARMASHERRERIA" & Err.number & " D: " & Err.Description)

End Sub

Sub LoadArmadurasHerreria()

    On Error GoTo fallo

    Dim n As Integer, LC As Integer

    n = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

    ReDim Preserve ArmadurasHerrero(1 To n) As Integer

    For LC = 1 To n
        ArmadurasHerrero(LC) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & LC, "Index"))
        'pluto:6.0a
        ObjData(ArmadurasHerrero(LC)).ParaHerre = 1
    Next LC

    Exit Sub
fallo:
    Call LogError("LOADARMADURASHERRERIA" & Err.number & " D: " & Err.Description)

End Sub

Sub LoadPorcentajesMascotas()

    PMascotas(1).Tipo = "Unicornio"
    PMascotas(2).Tipo = "Caballo Negro"
    PMascotas(3).Tipo = "Tigre"
    PMascotas(4).Tipo = "Elefante"
    PMascotas(5).Tipo = "Dragón"
    PMascotas(6).Tipo = "Jabato"
    PMascotas(7).Tipo = "Kong"
    PMascotas(8).Tipo = "Hipogrifo"
    PMascotas(9).Tipo = "Rinosaurio"
    PMascotas(10).Tipo = "Cerbero"
    PMascotas(11).Tipo = "Wyvern"
    PMascotas(12).Tipo = "Avestruz"

    'unicornio
    PMascotas(1).AumentoMagia = 15
    PMascotas(1).ReduceMagia = 9
    PMascotas(1).AumentoEvasion = 6
    PMascotas(1).VidaporLevel = 35
    PMascotas(1).GolpeporLevel = 13
    PMascotas(1).TopeAtMagico = 15
    PMascotas(1).TopeDefMagico = 9
    PMascotas(1).TopeEvasion = 6
    'negro
    PMascotas(2).AumentoMagia = 15
    PMascotas(2).ReduceMagia = 3
    PMascotas(2).AumentoEvasion = 1
    PMascotas(2).VidaporLevel = 30
    PMascotas(2).GolpeporLevel = 13
    PMascotas(2).TopeAtMagico = 9
    PMascotas(2).TopeDefMagico = 15
    PMascotas(2).TopeEvasion = 6
    'tigre
    PMascotas(3).ReduceCuerpo = 2
    PMascotas(3).AumentoEvasion = 6
    PMascotas(3).AumentoFlecha = 4
    PMascotas(3).VidaporLevel = 35
    PMascotas(3).GolpeporLevel = 13
    PMascotas(3).TopeAtFlechas = 9
    PMascotas(3).TopeDefMagico = 9
    PMascotas(3).TopeEvasion = 12
    'elefante
    PMascotas(4).AumentoCuerpo = 6
    PMascotas(4).ReduceCuerpo = 1
    PMascotas(4).ReduceFlecha = 1
    PMascotas(4).VidaporLevel = 50
    PMascotas(4).GolpeporLevel = 13
    PMascotas(4).TopeAtCuerpo = 15
    PMascotas(4).TopeDefCuerpo = 9
    PMascotas(4).TopeEvasion = 6
    'dragon
    PMascotas(5).AumentoCuerpo = 6
    PMascotas(5).ReduceCuerpo = 6
    PMascotas(5).AumentoMagia = 6
    PMascotas(5).ReduceMagia = 6
    PMascotas(5).AumentoFlecha = 6
    PMascotas(5).ReduceFlecha = 6
    PMascotas(5).AumentoEvasion = 6
    PMascotas(5).VidaporLevel = 80
    PMascotas(5).GolpeporLevel = 28
    PMascotas(5).TopeAtMagico = 9
    PMascotas(5).TopeDefMagico = 9
    PMascotas(5).TopeEvasion = 9
    PMascotas(5).TopeAtCuerpo = 9
    PMascotas(5).TopeDefCuerpo = 9
    PMascotas(5).TopeAtFlechas = 9
    PMascotas(5).TopeDefFlechas = 9
    'jabalí pequeño
    PMascotas(6).AumentoCuerpo = 1
    PMascotas(6).ReduceCuerpo = 6
    PMascotas(6).ReduceFlecha = 0
    PMascotas(6).VidaporLevel = 7
    PMascotas(6).GolpeporLevel = 13
    PMascotas(6).TopeAtMagico = 16
    PMascotas(6).TopeDefMagico = 16
    PMascotas(6).TopeEvasion = 16
    PMascotas(6).TopeAtCuerpo = 16
    PMascotas(6).TopeDefCuerpo = 16
    PMascotas(6).TopeAtFlechas = 16
    PMascotas(6).TopeDefFlechas = 16
    'kong
    PMascotas(7).AumentoCuerpo = 6
    PMascotas(7).ReduceCuerpo = 6
    PMascotas(7).AumentoMagia = 6
    PMascotas(7).ReduceMagia = 6
    PMascotas(7).AumentoFlecha = 6
    PMascotas(7).ReduceFlecha = 6
    PMascotas(7).AumentoEvasion = 6
    PMascotas(7).VidaporLevel = 80
    PMascotas(7).GolpeporLevel = 35
    PMascotas(7).TopeAtMagico = 9
    PMascotas(7).TopeDefMagico = 9
    PMascotas(7).TopeEvasion = 9
    PMascotas(7).TopeAtCuerpo = 9
    PMascotas(7).TopeDefCuerpo = 9
    PMascotas(7).TopeAtFlechas = 9
    PMascotas(7).TopeDefFlechas = 9
    'Crom
    PMascotas(8).AumentoCuerpo = 6
    PMascotas(8).ReduceCuerpo = 6
    PMascotas(8).AumentoMagia = 6
    PMascotas(8).ReduceMagia = 6
    PMascotas(8).AumentoFlecha = 6
    PMascotas(8).ReduceFlecha = 6
    PMascotas(8).AumentoEvasion = 6
    PMascotas(8).VidaporLevel = 80
    PMascotas(8).GolpeporLevel = 35
    PMascotas(8).TopeAtMagico = 9
    PMascotas(8).TopeDefMagico = 9
    PMascotas(8).TopeEvasion = 9
    PMascotas(8).TopeAtCuerpo = 9
    PMascotas(8).TopeDefCuerpo = 9
    PMascotas(8).TopeAtFlechas = 9
    PMascotas(8).TopeDefFlechas = 9
    'rinosaurio
    PMascotas(9).AumentoCuerpo = 6
    PMascotas(9).ReduceCuerpo = 6
    PMascotas(9).ReduceFlecha = 1
    PMascotas(9).VidaporLevel = 55
    PMascotas(9).GolpeporLevel = 13
    PMascotas(9).TopeEvasion = 9
    PMascotas(9).TopeDefMagico = 15
    PMascotas(9).TopeAtCuerpo = 6
    'cerbero
    PMascotas(10).ReduceCuerpo = 6
    PMascotas(10).AumentoEvasion = 2
    PMascotas(10).AumentoFlecha = 6
    PMascotas(10).VidaporLevel = 45
    PMascotas(10).GolpeporLevel = 13
    PMascotas(10).TopeAtFlechas = 6
    PMascotas(10).TopeDefMagico = 12
    PMascotas(10).TopeDefCuerpo = 12
    'wyvern
    PMascotas(11).AumentoMagia = 6
    PMascotas(11).ReduceMagia = 4
    PMascotas(11).AumentoEvasion = 3
    PMascotas(11).VidaporLevel = 40
    PMascotas(11).GolpeporLevel = 13
    PMascotas(11).TopeDefFlechas = 9
    PMascotas(11).TopeAtMagico = 12
    PMascotas(11).TopeDefMagico = 9
    'avestruz
    PMascotas(12).ReduceCuerpo = 1
    PMascotas(12).AumentoEvasion = 2
    PMascotas(12).AumentoFlecha = 6
    PMascotas(12).VidaporLevel = 35
    PMascotas(12).GolpeporLevel = 13
    PMascotas(12).TopeAtFlechas = 15
    PMascotas(12).TopeDefFlechas = 9
    PMascotas(12).TopeEvasion = 6
    'tope niveles
    PMascotas(1).TopeLevel = 30
    PMascotas(2).TopeLevel = 30
    PMascotas(3).TopeLevel = 30
    PMascotas(4).TopeLevel = 30
    PMascotas(5).TopeLevel = 16
    PMascotas(6).TopeLevel = 16
    PMascotas(7).TopeLevel = 17
    PMascotas(8).TopeLevel = 17
    PMascotas(9).TopeLevel = 30
    PMascotas(10).TopeLevel = 30
    PMascotas(11).TopeLevel = 30
    PMascotas(12).TopeLevel = 30

    'pluto:6.0A cargamos exp mascotas
    Dim n As Byte
    Dim nn As Byte
    Dim aa As Integer
    Dim bb As Long
    Dim cc As Long

    For n = 1 To 30
        aa = aa + 400
        bb = bb + 1800
        cc = cc + 20

        For nn = 1 To 12

            If nn = 5 Or nn = 7 Or nn = 8 Then
                PMascotas(nn).exp(n) = PMascotas(nn).exp(n) + bb
            ElseIf nn = 6 Then
                PMascotas(nn).exp(n) = PMascotas(nn).exp(n) + cc
            Else
                PMascotas(nn).exp(n) = PMascotas(nn).exp(n) + aa

            End If

        Next nn
    Next n

End Sub

Sub LoadObjCarpintero()

    On Error GoTo fallo

    Dim n As Integer, LC As Integer

    n = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

    ReDim Preserve ObjCarpintero(1 To n) As Integer

    For LC = 1 To n
        ObjCarpintero(LC) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & LC, "Index"))
        'pluto:6.0a
        ObjData(ObjCarpintero(LC)).ParaCarpin = 1
    Next LC

    Exit Sub
fallo:
    Call LogError("LOADOBJCARPINTERO" & Err.number & " D: " & Err.Description)

End Sub

'[MerLiNz:6]
Sub LoadObjMagicosermitano()

    On Error GoTo fallo

    Dim n As Integer, LC As Integer

    n = val(GetVar(DatPath & "Objermitano.dat", "INIT", "NumObjs"))

    ReDim Preserve Objermitano(1 To n) As Integer

    For LC = 1 To n
        Objermitano(LC) = val(GetVar(DatPath & "Objermitano.dat", "Obj" & LC, "Index"))
        'pluto:6.0a
        ObjData(Objermitano(LC)).ParaErmi = 1
    Next LC

    Exit Sub
fallo:
    Call LogError("LOADOBJMAGICOERMITAÑO" & Err.number & " D: " & Err.Description)

    '[\END]
End Sub

'Pluto:hoy
Sub Loadtrivial()

    On Error GoTo perro

    Dim n As Integer
    Dim numtrivial As Integer
    Dim leer As New clsIniManager
    Dim obj As ObjData

    leer.Initialize DatPath & "Trivial.txt"

    'numtrivial = val(GetVar(DatPath & "Trivial.txt", "INIT", "NumTrivial"))
    numtrivial = val(leer.GetValue("INIT", "NumTrivial"))

    n = RandomNumber(1, numtrivial)
    'PreTrivial = GetVar(DatPath & "TRIVIAL.TXT", "T" & n, "tx")
    PreTrivial = leer.GetValue("T" & n, "tx")

    'ResTrivial = GetVar(DatPath & "TRIVIAL.TXT", "T" & n, "RES")
    ResTrivial = leer.GetValue("T" & n, "RES")

    Exit Sub

perro:
    LogError ("Trivial: Error en la pregunta numero: " & n & " : " & Err.Description)

End Sub

'Pluto:2.4
Sub Loadrecord()

    On Error GoTo perro

    NivCrimi = val(GetVar(IniPath & "RECORD.TXT", "INIT", "NivCrimi"))
    NivCiu = val(GetVar(IniPath & "RECORD.TXT", "INIT", "NivCiu"))
    MaxTorneo = val(GetVar(IniPath & "RECORD.TXT", "INIT", "MaxTorneo"))
    Moro = val(GetVar(IniPath & "RECORD.TXT", "INIT", "Moro"))
    NNivCrimi = GetVar(IniPath & "RECORD.TXT", "INIT", "NNivCrimi")
    NNivCiu = GetVar(IniPath & "RECORD.TXT", "INIT", "NNivCiu")
    NMaxTorneo = GetVar(IniPath & "RECORD.TXT", "INIT", "NMaxTorneo")
    NMoro = GetVar(IniPath & "RECORD.TXT", "INIT", "NMoro")
    'pluto:6.9
    'Clan1Torneo = GetVar(IniPath & "RECORD.TXT", "INIT", "Clan1Torneo")
    'Clan2Torneo = GetVar(IniPath & "RECORD.TXT", "INIT", "Clan2Torneo")
    'PClan1Torneo = val(GetVar(IniPath & "RECORD.TXT", "INIT", "PClan1Torneo"))
    'PClan2Torneo = val(GetVar(IniPath & "RECORD.TXT", "INIT", "PClan2Torneo"))
    Exit Sub
perro:
    LogError ("Records: Error en cargando Records: " & Err.Description)

End Sub

'Pluto:hoy
Sub LoadEgipto()

    On Error GoTo perro

    Dim n As Integer
    Dim numegipto As Integer
    numegipto = val(GetVar(DatPath & "egipto.txt", "INIT", "NumEgipto"))
    n = RandomNumber(1, numegipto)
    PreEgipto = GetVar(DatPath & "EGIPTO.TXT", "T" & n, "tx")
    ResEgipto = GetVar(DatPath & "EGIPTO.TXT", "T" & n, "RES")
    Exit Sub
perro:
    LogError ("Egipto: Error en la pregunta numero: " & n & " : " & Err.Description)

End Sub

Sub LoadOBJData()

    On Error GoTo errhandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Object As Integer
    'pluto fusion
    Dim leer As New clsIniManager
    leer.Initialize DatPath & "Obj.dat"

    'obtiene el numero de obj
    NumObjDatas = val(leer.GetValue("INIT", "NumObjs"))

    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumObjDatas
    frmCargando.cargar.value = 0

    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    Dim Calcu As Double

    'Llena la lista
    For Object = 1 To NumObjDatas
        Calcu = Object
        Calcu = Calcu * 100
        Calcu = Calcu / NumObjDatas
        frmCargando.Label1(2).Caption = "Objeto: (" & Object & "/" & NumObjDatas & ") " & Round(Calcu, 1) & "%"

        ObjData(Object).Name = leer.GetValue("OBJ" & Object, "Name")
        'ObjData(Object).Name = Leer.GetValue("OBJ" & Object, "Name")
        'pluto 2.17
        ObjData(Object).Magia = val(leer.GetValue("OBJ" & Object, "Magia"))

        'pluto:2.8.0
        ObjData(Object).Vendible = val(leer.GetValue("OBJ" & Object, "Vendible"))

        ObjData(Object).GrhIndex = val(leer.GetValue("OBJ" & Object, "GrhIndex"))

        ObjData(Object).OBJType = val(leer.GetValue("OBJ" & Object, "ObjType"))
        ObjData(Object).SubTipo = val(leer.GetValue("OBJ" & Object, "Subtipo"))
        'pluto:6.0A
        ObjData(Object).ArmaNpc = val(leer.GetValue("OBJ" & Object, "ArmaNpc"))

        ObjData(Object).Newbie = val(leer.GetValue("OBJ" & Object, "Newbie"))
        'pluto:2.3
        ObjData(Object).Peso = 0    ' val(Leer.GetValue("OBJ" & Object, "Peso"))

        If ObjData(Object).SubTipo = OBJTYPE_ESCUDO Then
            ObjData(Object).ShieldAnim = val(leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))  ' * 2
            '[MerLiNz:6]
            ObjData(Object).Gemas = val(leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = val(leer.GetValue("OBJ" & Object, "Diamantes"))

            '[\END]
        End If

        'pluto:6.2----------
        If ObjData(Object).OBJType = OBJTYPE_Anillo Then
            ObjData(Object).LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))  ' * 2
            ObjData(Object).Gemas = val(leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = val(leer.GetValue("OBJ" & Object, "Diamantes"))

        End If

        '--------------------

        If ObjData(Object).SubTipo = OBJTYPE_CASCO Then

            ObjData(Object).CascoAnim = val(leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
            '[MerLiNz:6]
            ObjData(Object).Gemas = val(leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = val(leer.GetValue("OBJ" & Object, "Diamantes"))
            '[\END]
            ObjData(Object).SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))  '* 2

        End If

        If ObjData(Object).SubTipo = OBJTYPE_ALAS Then
            ObjData(Object).AlasAnim = val(leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
            '[MerLiNz:6]
            ObjData(Object).Gemas = val(leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = val(leer.GetValue("OBJ" & Object, "Diamantes"))
            '[\END]
            ObjData(Object).SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))  '* 2

        End If

        '[GAU]
        If ObjData(Object).SubTipo = OBJTYPE_BOTA Then
            ObjData(Object).Botas = val(leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))  ' * 2

        End If

        '[GAU]
        ObjData(Object).Ropaje = val(leer.GetValue("OBJ" & Object, "NumRopaje"))
        ObjData(Object).HechizoIndex = val(leer.GetValue("OBJ" & Object, "HechizoIndex"))

        If ObjData(Object).OBJType = OBJTYPE_WEAPON Then
            ObjData(Object).WeaponAnim = val(leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = val(leer.GetValue("OBJ" & Object, "Apuñala"))
            ObjData(Object).Envenena = val(leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = val(leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))  ' * 2
            ObjData(Object).Real = val(leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(leer.GetValue("OBJ" & Object, "Caos"))
            ObjData(Object).proyectil = val(leer.GetValue("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = val(leer.GetValue("OBJ" & Object, "Municiones"))
            '[MerLiNz:6]
            ObjData(Object).Gemas = val(leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = val(leer.GetValue("OBJ" & Object, "Diamantes"))
            '[\END]
            ObjData(Object).SkArma = val(leer.GetValue("OBJ" & Object, "SKARMA"))
            ObjData(Object).SkArco = val(leer.GetValue("OBJ" & Object, "SKARCO"))

        End If

        If ObjData(Object).OBJType = OBJTYPE_ARMOUR Then
            ObjData(Object).LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))  ' * 2
            ObjData(Object).Real = val(leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(leer.GetValue("OBJ" & Object, "Caos"))
            '[MerLiNz:6]
            ObjData(Object).Gemas = val(leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = val(leer.GetValue("OBJ" & Object, "Diamantes"))
            'pluto:2.10
            ObjData(Object).ObjetoClan = leer.GetValue("OBJ" & Object, "ObjetoClan")

            '[\END]
        End If

        If ObjData(Object).OBJType = OBJTYPE_HERRAMIENTAS Then
            ObjData(Object).LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))  '* 2
            '[MerLiNz:6]
            ObjData(Object).Gemas = val(leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = val(leer.GetValue("OBJ" & Object, "Diamantes"))

            '[\END]
        End If

        If ObjData(Object).OBJType = OBJTYPE_INSTRUMENTOS Then
            ObjData(Object).Snd1 = val(leer.GetValue("OBJ" & Object, "SND1"))
            ObjData(Object).Snd2 = val(leer.GetValue("OBJ" & Object, "SND2"))
            ObjData(Object).Snd3 = val(leer.GetValue("OBJ" & Object, "SND3"))
            ObjData(Object).MinInt = val(leer.GetValue("OBJ" & Object, "MinInt"))

        End If

        ObjData(Object).LingoteIndex = val(leer.GetValue("OBJ" & Object, "LingoteIndex"))

        If ObjData(Object).OBJType = 31 Or ObjData(Object).OBJType = 23 Then
            ObjData(Object).MinSkill = val(leer.GetValue("OBJ" & Object, "MinSkill"))

        End If

        ObjData(Object).MineralIndex = val(leer.GetValue("OBJ" & Object, "MineralIndex"))

        ObjData(Object).MaxHP = val(leer.GetValue("OBJ" & Object, "MaxHP"))
        ObjData(Object).MinHP = val(leer.GetValue("OBJ" & Object, "MinHP"))

        ObjData(Object).Mujer = val(leer.GetValue("OBJ" & Object, "Mujer"))
        ObjData(Object).Hombre = val(leer.GetValue("OBJ" & Object, "Hombre"))

        ObjData(Object).MinHam = val(leer.GetValue("OBJ" & Object, "MinHam"))
        ObjData(Object).MinSed = val(leer.GetValue("OBJ" & Object, "MinAgu"))

        'pluto:7.0
        ObjData(Object).MinDef = val(leer.GetValue("OBJ" & Object, "MINDEF"))
        ObjData(Object).MaxDef = val(leer.GetValue("OBJ" & Object, "MAXDEF"))
        ObjData(Object).Defmagica = val(leer.GetValue("OBJ" & Object, "DEFMAGICA"))
        'nati:agrego DefCuerpo
        ObjData(Object).Defcuerpo = val(leer.GetValue("OBJ" & Object, "DEFCUERPO"))
        ObjData(Object).Drop = val(leer.GetValue("OBJ" & Object, "DROP"))

        'ObjData(Object).Defproyectil = val(Leer.GetValue("OBJ" & Object, "DEFPROYECTIL"))

        ObjData(Object).Respawn = val(leer.GetValue("OBJ" & Object, "ReSpawn"))

        ObjData(Object).RazaEnana = val(leer.GetValue("OBJ" & Object, "RazaEnana"))
        ObjData(Object).razaelfa = val(leer.GetValue("OBJ" & Object, "RazaElfa"))
        ObjData(Object).razavampiro = val(leer.GetValue("OBJ" & Object, "Razavampiro"))
        ObjData(Object).razaorca = val(leer.GetValue("OBJ" & Object, "Razaorca"))
        ObjData(Object).razahumana = val(leer.GetValue("OBJ" & Object, "Razahumana"))

        ObjData(Object).Valor = val(leer.GetValue("OBJ" & Object, "Valor"))
        ObjData(Object).nocaer = val(leer.GetValue("OBJ" & Object, "nocaer"))
        ObjData(Object).objetoespecial = val(leer.GetValue("OBJ" & Object, "objetoespecial"))

        ObjData(Object).Crucial = val(leer.GetValue("OBJ" & Object, "Crucial"))

        ObjData(Object).Cerrada = val(leer.GetValue("OBJ" & Object, "abierta"))

        If ObjData(Object).Cerrada = 1 Then
            ObjData(Object).Llave = val(leer.GetValue("OBJ" & Object, "Llave"))
            ObjData(Object).Clave = val(leer.GetValue("OBJ" & Object, "Clave"))

        End If

        If ObjData(Object).OBJType = OBJTYPE_PUERTAS Or ObjData(Object).OBJType = OBJTYPE_BOTELLAVACIA Or ObjData( _
           Object).OBJType = OBJTYPE_BOTELLALLENA Then
            ObjData(Object).IndexAbierta = val(leer.GetValue("OBJ" & Object, "IndexAbierta"))
            ObjData(Object).IndexCerrada = val(leer.GetValue("OBJ" & Object, "IndexCerrada"))
            ObjData(Object).IndexCerradaLlave = val(leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))

        End If

        'Puertas y llaves
        ObjData(Object).Clave = val(leer.GetValue("OBJ" & Object, "Clave"))

        ObjData(Object).Texto = leer.GetValue("OBJ" & Object, "Texto")
        ObjData(Object).GrhSecundario = val(leer.GetValue("OBJ" & Object, "VGrande"))

        ObjData(Object).Agarrable = val(leer.GetValue("OBJ" & Object, "Agarrable"))
        ObjData(Object).ForoID = leer.GetValue("OBJ" & Object, "ID")

        Dim i As Integer, tStr As String

        For i = 1 To NUMCLASES

            tStr = leer.GetValue("OBJ" & Object, "CP" & i)

            If tStr <> "" Then
                tStr = mid$(tStr, 1, Len(tStr) - 1)
                tStr = Right$(tStr, Len(tStr) - 1)
            End If

            ObjData(Object).ClaseProhibida(i) = tStr
        Next

        ObjData(Object).Resistencia = val(leer.GetValue("OBJ" & Object, "Resistencia"))

        'Pociones
        If ObjData(Object).OBJType = 11 Then
            ObjData(Object).TipoPocion = val(leer.GetValue("OBJ" & Object, "TipoPocion"))
            ObjData(Object).MaxModificador = val(leer.GetValue("OBJ" & Object, "MaxModificador"))
            ObjData(Object).MinModificador = val(leer.GetValue("OBJ" & Object, "MinModificador"))
            ObjData(Object).DuracionEfecto = val(leer.GetValue("OBJ" & Object, "DuracionEfecto"))

        End If

        ObjData(Object).SkCarpinteria = val(leer.GetValue("OBJ" & Object, "SkCarpinteria"))  '* 2

        If ObjData(Object).SkCarpinteria > 0 Then ObjData(Object).Madera = val(leer.GetValue("OBJ" & Object, "Madera"))

        If ObjData(Object).OBJType = OBJTYPE_BARCOS Then
            ObjData(Object).MaxHIT = val(leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(leer.GetValue("OBJ" & Object, "MinHIT"))

        End If

        If ObjData(Object).OBJType = OBJTYPE_FLECHAS Then
            ObjData(Object).MaxHIT = val(leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(leer.GetValue("OBJ" & Object, "MinHIT"))

        End If
        
        If ObjData(Object).OBJType = OBJTYPE_HUEVOS Then
            ObjData(Object).Doma = val(leer.GetValue("OBJ" & Object, "Doma"))

        End If

        'Bebidas
        ObjData(Object).MinSta = val(leer.GetValue("OBJ" & Object, "MinST"))
        ObjData(Object).razavampiro = val(leer.GetValue("OBJ" & Object, "razavampiro"))
        'pluto:6.0A----
        ObjData(Object).Cregalos = val(leer.GetValue("OBJ" & Object, "Cregalos"))
        ObjData(Object).Pregalo = val(leer.GetValue("OBJ" & Object, "Pregalo"))
        '--------------
        frmCargando.cargar.value = frmCargando.cargar.value + 1

        'pluto:6.0A
        If ObjData(Object).Pregalo > 0 Then

            Select Case ObjData(Object).Pregalo

            Case 1
                Reo1 = Reo1 + 1
                ObjRegalo1(Reo1) = Object

            Case 2
                Reo2 = Reo2 + 1
                ObjRegalo2(Reo2) = Object

            Case 3
                Reo3 = Reo3 + 1
                ObjRegalo3(Reo3) = Object

            End Select

        End If

    Next Object

    'quitar esto
    Exit Sub
    '------------------------------------------------------------------------------------
    'Esto genera el obj.log para meterlo al cliente, el server no usa nada de lo de abajo.
    '------------------------------------------------------------------------------------
    Dim File As String
    Dim n As Byte
    File = DatPath & "Obj.dat"
    Dim nfile As Integer
    Dim vec As Byte
    Dim vec2 As Integer
    vec = 1
    nfile = FreeFile    ' obtenemos un canal
    Open App.Path & "\Objeto.log" For Append Shared As #nfile
    Print #nfile, "Sub CargamosObjetos" & vec & "()"

    For Object = 1 To NumObjDatas
        vec2 = vec2 + 1
        Debug.Print Object

        If vec2 > 100 Then
            vec = vec + 1
            vec2 = 0
            Print #nfile, "end sub"
            Print #nfile, "sub CargamosObjetos" & vec & "()"

        End If

        Print #nfile, "ObjData(" & Object & ").name=" & Chr(34) & ObjData(Object).Name & Chr(34)

        If ObjData(Object).Agarrable > 0 Then Print #nfile, "ObjData(" & Object & ").agarrable =" & ObjData( _
                                                            Object).Agarrable

        If ObjData(Object).Apuñala > 0 Then Print #nfile, "ObjData(" & Object & ").apuñala=" & ObjData(Object).Apuñala
        If ObjData(Object).ArmaNpc > 0 Then Print #nfile, "ObjData(" & Object & ").armanpc=" & ObjData(Object).ArmaNpc
        If ObjData(Object).Botas > 0 Then Print #nfile, "ObjData(" & Object & ").botas=" & ObjData(Object).Botas
        If ObjData(Object).AlasAnim > 0 Then Print #nfile, "ObjData(" & Object & ").alasanim=" & ObjData( _
                                                           Object).AlasAnim

        If ObjData(Object).Caos > 0 Then Print #nfile, "ObjData(" & Object & ").caos=" & ObjData(Object).Caos
        If ObjData(Object).CascoAnim > 0 Then Print #nfile, "ObjData(" & Object & ").cascoanim=" & ObjData( _
                                                            Object).CascoAnim

        If ObjData(Object).Cerrada > 0 Then Print #nfile, "ObjData(" & Object & ").cerrada=" & ObjData(Object).Cerrada

        For n = 1 To 21

            If ObjData(Object).ClaseProhibida(n) <> "" Then Print #nfile, "ObjData(" & Object & ").claseprohibida(" & _
                                                                          n & ")=" & Chr(34) & ObjData(Object).ClaseProhibida(n) & Chr(34)
        Next

        If ObjData(Object).Clave > 0 Then Print #nfile, "ObjData(" & Object & ").clave=" & ObjData(Object).Clave
        If ObjData(Object).Crucial > 0 Then Print #nfile, "ObjData(" & Object & ").crucial=" & ObjData(Object).Crucial
        If ObjData(Object).Def > 0 Then Print #nfile, "ObjData(" & Object & ").def=" & ObjData(Object).Def
        If ObjData(Object).Diamantes > 0 Then Print #nfile, "ObjData(" & Object & ").diamantes=" & ObjData( _
                                                            Object).Diamantes

        If ObjData(Object).DuracionEfecto > 0 Then Print #nfile, "ObjData(" & Object & ").duracionefecto=" & ObjData( _
                                                                 Object).DuracionEfecto

        If ObjData(Object).Envenena > 0 Then Print #nfile, "ObjData(" & Object & ").envenena=" & ObjData( _
                                                           Object).Envenena

        If ObjData(Object).ForoID <> "" Then Print #nfile, "ObjData(" & Object & ").foroid=" & Chr(34) & ObjData( _
                                                           Object).ForoID & Chr(34)

        If ObjData(Object).Gemas > 0 Then Print #nfile, "ObjData(" & Object & ").gemas=" & ObjData(Object).Gemas
        If ObjData(Object).GrhIndex > 0 Then Print #nfile, "ObjData(" & Object & ").grhindex=" & ObjData( _
                                                           Object).GrhIndex

        If ObjData(Object).GrhSecundario > 0 Then Print #nfile, "ObjData(" & Object & ").grhsecundario=" & ObjData( _
                                                                Object).GrhSecundario

        If ObjData(Object).HechizoIndex > 0 Then Print #nfile, "ObjData(" & Object & ").hechizoindex=" & ObjData( _
                                                               Object).HechizoIndex

        If ObjData(Object).Hombre > 0 Then Print #nfile, "ObjData(" & Object & ").hombre=" & ObjData(Object).Hombre
        If ObjData(Object).IndexAbierta > 0 Then Print #nfile, "ObjData(" & Object & ").indexabierta=" & ObjData( _
                                                               Object).IndexAbierta

        If ObjData(Object).IndexCerrada > 0 Then Print #nfile, "ObjData(" & Object & ").indexcerrada=" & ObjData( _
                                                               Object).IndexCerrada

        If ObjData(Object).IndexCerradaLlave > 0 Then Print #nfile, "ObjData(" & Object & ").indexcerradallave=" & _
                                                                    ObjData(Object).IndexCerradaLlave

        If ObjData(Object).LingH > 0 Then Print #nfile, "ObjData(" & Object & ").lingh=" & ObjData(Object).LingH
        If ObjData(Object).LingO > 0 Then Print #nfile, "ObjData(" & Object & ").lingo=" & ObjData(Object).LingO
        If ObjData(Object).LingoteIndex > 0 Then Print #nfile, "ObjData(" & Object & ").lingoteindex=" & ObjData( _
                                                               Object).LingoteIndex

        If ObjData(Object).LingP > 0 Then Print #nfile, "ObjData(" & Object & ").lingp=" & ObjData(Object).LingP
        If ObjData(Object).Llave > 0 Then Print #nfile, "ObjData(" & Object & ").llave=" & ObjData(Object).Llave
        If ObjData(Object).Madera > 0 Then Print #nfile, "ObjData(" & Object & ").madera=" & ObjData(Object).Madera
        If ObjData(Object).Magia > 0 Then Print #nfile, "ObjData(" & Object & ").magia=" & ObjData(Object).Magia
        If ObjData(Object).MaxDef > 0 Then Print #nfile, "ObjData(" & Object & ").maxdef=" & ObjData(Object).MaxDef
        If ObjData(Object).MaxHIT > 0 Then Print #nfile, "ObjData(" & Object & ").maxhit=" & ObjData(Object).MaxHIT
        If ObjData(Object).MaxHP > 0 Then Print #nfile, "ObjData(" & Object & ").maxhp=" & ObjData(Object).MaxHP
        If ObjData(Object).MaxItems > 0 Then Print #nfile, "ObjData(" & Object & ").maxitems=" & ObjData( _
                                                           Object).MaxItems

        If ObjData(Object).MaxModificador > 0 Then Print #nfile, "ObjData(" & Object & ").maxmodificador=" & ObjData( _
                                                                 Object).MaxModificador

        If ObjData(Object).MinDef > 0 Then Print #nfile, "ObjData(" & Object & ").mindef=" & ObjData(Object).MinDef

        'pluto:7.0
        If ObjData(Object).Defmagica > 0 Then Print #nfile, "ObjData(" & Object & ").defmagica =" & ObjData( _
                                                            Object).Defmagica

        'nati: Agrego defCuerpo
        If ObjData(Object).Defcuerpo > 0 Then Print #nfile, "ObjData(" & Object & ").defcuerpo =" & ObjData( _
                                                            Object).Defcuerpo
        'If ObjData(Object).Defproyectil > 0 Then Print #nfile, "ObjData(" & Object & ").defproyectil =" & ObjData(Object).Defproyectil

        If ObjData(Object).MineralIndex > 0 Then Print #nfile, "ObjData(" & Object & ").mineralindex=" & ObjData( _
                                                               Object).MineralIndex

        If ObjData(Object).MinHam > 0 Then Print #nfile, "ObjData(" & Object & ").minham=" & ObjData(Object).MinHam
        If ObjData(Object).MinHIT > 0 Then Print #nfile, "ObjData(" & Object & ").minhit=" & ObjData(Object).MinHIT
        If ObjData(Object).MinHP > 0 Then Print #nfile, "ObjData(" & Object & ").minhp=" & ObjData(Object).MinHP
        If ObjData(Object).MinInt > 0 Then Print #nfile, "ObjData(" & Object & ").minint=" & ObjData(Object).MinInt
        If ObjData(Object).MinModificador > 0 Then Print #nfile, "ObjData(" & Object & ").minmodificador=" & ObjData( _
                                                                 Object).MinModificador

        If ObjData(Object).MinSed > 0 Then Print #nfile, "ObjData(" & Object & ").minsed=" & ObjData(Object).MinSed
        If ObjData(Object).MinSkill > 0 Then Print #nfile, "ObjData(" & Object & ").minskill=" & ObjData( _
                                                           Object).MinSkill

        If ObjData(Object).MinSta > 0 Then Print #nfile, "ObjData(" & Object & ").minsta=" & ObjData(Object).MinSta
        If ObjData(Object).Mujer > 0 Then Print #nfile, "ObjData(" & Object & ").mujer=" & ObjData(Object).Mujer
        If ObjData(Object).Municion > 0 Then Print #nfile, "ObjData(" & Object & ").municion=" & ObjData( _
                                                           Object).Municion

        If ObjData(Object).Newbie > 0 Then Print #nfile, "ObjData(" & Object & ").Newbie=" & ObjData(Object).Newbie
        If ObjData(Object).nocaer > 0 Then Print #nfile, "ObjData(" & Object & ").nocaer=" & ObjData(Object).nocaer
        If ObjData(Object).ObjetoClan <> "" Then Print #nfile, "ObjData(" & Object & ").objetoclan=" & Chr(34) & _
                                                               ObjData(Object).ObjetoClan & Chr(34)

        If ObjData(Object).objetoespecial > 0 Then Print #nfile, "ObjData(" & Object & ").objetoespecial=" & ObjData( _
                                                                 Object).objetoespecial

        If ObjData(Object).OBJType > 0 Then Print #nfile, "ObjData(" & Object & ").objtype=" & ObjData(Object).OBJType
        If ObjData(Object).Peso > 0 Then Print #nfile, "ObjData(" & Object & ").peso=" & ObjData(Object).Peso
        If ObjData(Object).proyectil > 0 Then Print #nfile, "ObjData(" & Object & ").proyectil=" & ObjData( _
                                                            Object).proyectil

        If ObjData(Object).razaelfa > 0 Then Print #nfile, "ObjData(" & Object & ").razaelfa=" & ObjData( _
                                                           Object).razaelfa

        If ObjData(Object).RazaEnana > 0 Then Print #nfile, "ObjData(" & Object & ").razaenana=" & ObjData( _
                                                            Object).RazaEnana

        If ObjData(Object).razahumana > 0 Then Print #nfile, "ObjData(" & Object & ").razahumana=" & ObjData( _
                                                             Object).razahumana

        If ObjData(Object).razaorca > 0 Then Print #nfile, "ObjData(" & Object & ").razaorca=" & ObjData( _
                                                           Object).razaorca

        If ObjData(Object).razavampiro > 0 Then Print #nfile, "ObjData(" & Object & ").razavampiro=" & ObjData( _
                                                              Object).razavampiro

        If ObjData(Object).Real > 0 Then Print #nfile, "ObjData(" & Object & ").real=" & ObjData(Object).Real
        If ObjData(Object).Resistencia > 0 Then Print #nfile, "ObjData(" & Object & ").resistencia=" & ObjData( _
                                                              Object).Resistencia

        If ObjData(Object).Respawn > 0 Then Print #nfile, "ObjData(" & Object & ").respawn=" & ObjData(Object).Respawn
        If ObjData(Object).Ropaje > 0 Then Print #nfile, "ObjData(" & Object & ").ropaje=" & ObjData(Object).Ropaje
        If ObjData(Object).ShieldAnim > 0 Then Print #nfile, "ObjData(" & Object & ").shieldanim=" & ObjData( _
                                                             Object).ShieldAnim

        If ObjData(Object).SkArco > 0 Then Print #nfile, "ObjData(" & Object & ").skarco=" & ObjData(Object).SkArco
        If ObjData(Object).SkArma > 0 Then Print #nfile, "ObjData(" & Object & ").skarma=" & ObjData(Object).SkArma
        If ObjData(Object).SkCarpinteria > 0 Then Print #nfile, "ObjData(" & Object & ").skcarpinteria=" & ObjData( _
                                                                Object).SkCarpinteria

        If ObjData(Object).SkHerreria > 0 Then Print #nfile, "ObjData(" & Object & ").skherreria=" & ObjData( _
                                                             Object).SkHerreria

        If ObjData(Object).Snd1 > 0 Then Print #nfile, "ObjData(" & Object & ").snd1=" & ObjData(Object).Snd1
        If ObjData(Object).Snd2 > 0 Then Print #nfile, "ObjData(" & Object & ").snd2=" & ObjData(Object).Snd2
        If ObjData(Object).Snd3 > 0 Then Print #nfile, "ObjData(" & Object & ").snd3=" & ObjData(Object).Snd3
        If ObjData(Object).SubTipo > 0 Then Print #nfile, "ObjData(" & Object & ").subtipo=" & ObjData(Object).SubTipo
        If ObjData(Object).Texto <> "" Then Print #nfile, "ObjData(" & Object & ").texto=" & Chr(34) & ObjData( _
                                                          Object).Texto & Chr(34)

        If ObjData(Object).TipoPocion > 0 Then Print #nfile, "ObjData(" & Object & ").tipopocion=" & ObjData( _
                                                             Object).TipoPocion

        If ObjData(Object).Valor > 0 Then Print #nfile, "ObjData(" & Object & ").valor=" & ObjData(Object).Valor
        If ObjData(Object).Vendible > 0 Then Print #nfile, "ObjData(" & Object & ").vendible=" & ObjData( _
                                                           Object).Vendible

        If ObjData(Object).WeaponAnim > 0 Then Print #nfile, "ObjData(" & Object & ").weaponanim=" & ObjData( _
                                                             Object).WeaponAnim

        If ObjData(Object).Pregalo > 0 Then Print #nfile, "ObjData(" & Object & ").pregalo=" & ObjData(Object).Pregalo
        If ObjData(Object).Cregalos > 0 Then Print #nfile, "ObjData(" & Object & ").cregalos=" & ObjData( _
                                                           Object).Cregalos

        'pluto:7.0
        If ObjData(Object).Drop > 0 Then Print #nfile, "ObjData(" & Object & ").drop=" & ObjData(Object).Drop

        DoEvents
    Next
    Close #nfile

    Exit Sub

errhandler:
    MsgBox "error cargando objetos"

End Sub

'pluto:2.3
Sub LoadUserMontura(Userindex As Integer, UserFile As String)
'on error GoTo fallo
'Dim LoopC As Integer
'Dim Leer As New clsLeerInis
'Leer.Initialize userfile
'For LoopC = 1 To MAXMONTURA
'UserList(UserIndex).Montura.Nivel(LoopC) = val(Leer.GetValue("MONTURA", "NIVEL" & LoopC))
'UserList(UserIndex).Montura.exp(LoopC) = val(Leer.GetValue("MONTURA", "EXP" & LoopC))
'UserList(UserIndex).Montura.Elu(LoopC) = val(Leer.GetValue("MONTURA", "ELU" & LoopC))
'UserList(UserIndex).Montura.Vida(LoopC) = val(Leer.GetValue("MONTURA", "VIDA" & LoopC))
'UserList(UserIndex).Montura.Golpe(LoopC) = val(Leer.GetValue("MONTURA", "GOLPE" & LoopC))
'UserList(UserIndex).Montura.Nombre(LoopC) = Leer.GetValue("MONTURA", "NOMBRE" & LoopC)

'Next

'Exit Sub
'fallo:
'Call LogError("LOADUSERMONTURA" & Err.Number & " D: " & Err.Description)

End Sub

Sub LoadUserStats(Userindex As Integer, UserFile As String)
'on error GoTo fallo
'Dim LoopC As Integer

'For LoopC = 1 To NUMATRIBUTOS
' UserList(UserIndex).Stats.UserAtributos(LoopC) = Leer.GetValue( "ATRIBUTOS", "AT" & LoopC)
'UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) = UserList(UserIndex).Stats.UserAtributos(LoopC)
'Next

'For LoopC = 1 To NUMSKILLS
' UserList(UserIndex).Stats.UserSkills(LoopC) = val(Leer.GetValue( "SKILLS", "SK" & LoopC))
'Next

'For LoopC = 1 To MAXUSERHECHIZOS
' UserList(UserIndex).Stats.UserHechizos(LoopC) = val(Leer.GetValue( "Hechizos", "H" & LoopC))
'Next
'pluto:2-3-04
'UserList(UserIndex).Stats.Puntos = val(Leer.GetValue( "STATS", "PUNTOS"))

'UserList(UserIndex).Stats.GLD = val(Leer.GetValue( "STATS", "GLD"))
'UserList(UserIndex).Remort = val(Leer.GetValue( "STATS", "REMORT"))
'UserList(UserIndex).Stats.Banco = val(Leer.GetValue( "STATS", "BANCO"))

'UserList(UserIndex).Stats.MET = val(Leer.GetValue( "STATS", "MET"))
'UserList(UserIndex).Stats.MaxHP = val(Leer.GetValue( "STATS", "MaxHP"))
'UserList(UserIndex).Stats.MinHP = val(Leer.GetValue( "STATS", "MinHP"))

'UserList(UserIndex).Stats.FIT = val(Leer.GetValue( "STATS", "FIT"))
'UserList(UserIndex).Stats.MinSta = val(Leer.GetValue( "STATS", "MinSTA"))
'UserList(UserIndex).Stats.MaxSta = val(Leer.GetValue( "STATS", "MaxSTA"))

'UserList(UserIndex).Stats.MaxMAN = val(Leer.GetValue( "STATS", "MaxMAN"))
'UserList(UserIndex).Stats.MinMAN = val(Leer.GetValue( "STATS", "MinMAN"))

'UserList(UserIndex).Stats.MaxHIT = val(Leer.GetValue( "STATS", "MaxHIT"))
'UserList(UserIndex).Stats.MinHIT = val(Leer.GetValue( "STATS", "MinHIT"))

'UserList(UserIndex).Stats.MaxAGU = val(Leer.GetValue( "STATS", "MaxAGU"))
'UserList(UserIndex).Stats.MinAGU = val(Leer.GetValue( "STATS", "MinAGU"))

'UserList(UserIndex).Stats.MaxHam = val(Leer.GetValue( "STATS", "MaxHAM"))
'UserList(UserIndex).Stats.MinHam = val(Leer.GetValue( "STATS", "MinHAM"))

'UserList(UserIndex).Stats.SkillPts = val(Leer.GetValue( "STATS", "SkillPtsLibres"))

'UserList(UserIndex).Stats.exp = val(Leer.GetValue( "STATS", "EXP"))
'UserList(UserIndex).Stats.Elu = val(Leer.GetValue( "STATS", "ELU"))
'UserList(UserIndex).Stats.ELV = val(Leer.GetValue( "STATS", "ELV"))
'pluto:2.4.5
'UserList(UserIndex).Stats.PClan = val(Leer.GetValue( "STATS", "PCLAN"))
'UserList(UserIndex).Stats.GTorneo = val(Leer.GetValue( "STATS", "GTORNEO"))

'UserList(UserIndex).Stats.UsuariosMatados = val(Leer.GetValue( "MUERTES", "UserMuertes"))
'UserList(UserIndex).Stats.CriminalesMatados = val(Leer.GetValue( "MUERTES", "CrimMuertes"))
'UserList(UserIndex).Stats.NPCsMuertos = val(Leer.GetValue( "MUERTES", "NpcsMuertes"))
'Exit Sub
'fallo:
'Call LogError("LOADUSERSTATS" & Err.Number & " D: " & Err.Description)

End Sub

Sub LoadUserReputacion(Userindex As Integer, UserFile As String)
'on error GoTo fallo
'UserList(UserIndex).Reputacion.AsesinoRep = val(Leer.GetValue( "REP", "Asesino"))
'UserList(UserIndex).Reputacion.BandidoRep = val(Leer.GetValue( "REP", "Dandido"))
'UserList(UserIndex).Reputacion.BurguesRep = val(Leer.GetValue( "REP", "Burguesia"))
'UserList(UserIndex).Reputacion.LadronesRep = val(Leer.GetValue( "REP", "Ladrones"))
'UserList(UserIndex).Reputacion.NobleRep = val(Leer.GetValue( "REP", "Nobles"))
'UserList(UserIndex).Reputacion.PlebeRep = val(Leer.GetValue( "REP", "Plebe"))
'UserList(UserIndex).Reputacion.Promedio = val(Leer.GetValue( "REP", "Promedio"))
'pluto:2-3-04
'If UserList(UserIndex).Faccion.FuerzasCaos > 0 And UserList(UserIndex).Reputacion.Promedio >= 0 Then Call ExpulsarCaos(UserIndex)
'Exit Sub
'fallo:
'Call LogError("LOADUSERREPUTACION" & Err.Number & " D: " & Err.Description)

End Sub

Sub LoadUserInit(Userindex As Integer, UserFile As String, Name As String)

    On Error GoTo fallo

    Dim loopc As Integer
    Dim ln As String
    Dim Ln2 As String
    'pluto:2.24

    Dim leer As New clsIniManager
    leer.Initialize UserFile

    UserList(Userindex).Faccion.ArmadaReal = val(leer.GetValue("FACCIONES", "EjercitoReal"))
    UserList(Userindex).Faccion.FuerzasCaos = val(leer.GetValue("FACCIONES", "EjercitoCaos"))
    UserList(Userindex).Faccion.CiudadanosMatados = val(leer.GetValue("FACCIONES", "CiudMatados"))
    UserList(Userindex).Faccion.NeutralesMatados = val(leer.GetValue("FACCIONES", "NeutMatados"))
    UserList(Userindex).Faccion.CriminalesMatados = val(leer.GetValue("FACCIONES", "CrimMatados"))
    UserList(Userindex).Faccion.RecibioArmaduraCaos = val(leer.GetValue("FACCIONES", "rArCaos"))
    UserList(Userindex).Faccion.RecibioArmaduraReal = val(leer.GetValue("FACCIONES", "rArReal"))
    UserList(Userindex).Faccion.RecibioArmaduraLegion = val(leer.GetValue("FACCIONES", "rArLegion"))
    UserList(Userindex).Faccion.RecibioExpInicialCaos = val(leer.GetValue("FACCIONES", "rExCaos"))
    UserList(Userindex).Faccion.RecibioExpInicialReal = val(leer.GetValue("FACCIONES", "rExReal"))
    UserList(Userindex).Faccion.RecompensasCaos = val(leer.GetValue("FACCIONES", "recCaos"))
    UserList(Userindex).Faccion.RecompensasReal = val(leer.GetValue("FACCIONES", "recReal"))
    UserList(Userindex).flags.LiderAlianza = val(leer.GetValue("FLAGS", "LiderAlianza"))
    UserList(Userindex).flags.LiderHorda = val(leer.GetValue("FLAGS", "LiderHorda"))
    UserList(Userindex).flags.Muerto = val(leer.GetValue("FLAGS", "Muerto"))
    UserList(Userindex).flags.Escondido = val(leer.GetValue("FLAGS", "Escondido"))
    UserList(Userindex).flags.Hambre = val(leer.GetValue("FLAGS", "Hambre"))
    UserList(Userindex).flags.Sed = val(leer.GetValue("FLAGS", "Sed"))
    UserList(Userindex).flags.Desnudo = val(leer.GetValue("FLAGS", "Desnudo"))
    UserList(Userindex).flags.Envenenado = val(leer.GetValue("FLAGS", "Envenenado"))
    UserList(Userindex).flags.Morph = val(leer.GetValue("FLAGS", "Morph"))
    UserList(Userindex).flags.Paralizado = val(leer.GetValue("FLAGS", "Paralizado"))
    UserList(Userindex).flags.Angel = val(leer.GetValue("FLAGS", "Angel"))
    UserList(Userindex).flags.Demonio = val(leer.GetValue("FLAGS", "Demonio"))
    'pluto:6.5
    UserList(Userindex).flags.Minotauro = val(leer.GetValue("FLAGS", "Minotauro"))
    UserList(Userindex).flags.MinutosOnline = val(leer.GetValue("FLAGS", "MinOn"))
    'pluto:7.0
    UserList(Userindex).flags.Creditos = val(leer.GetValue("FLAGS", "Creditos"))
    UserList(Userindex).flags.DragCredito1 = val(leer.GetValue("FLAGS", "DragC1"))
    UserList(Userindex).flags.DragCredito2 = val(leer.GetValue("FLAGS", "DragC2"))
    UserList(Userindex).flags.DragCredito3 = val(leer.GetValue("FLAGS", "DragC3"))
    UserList(Userindex).flags.DragCredito4 = val(leer.GetValue("FLAGS", "DragC4"))
    UserList(Userindex).flags.DragCredito5 = val(leer.GetValue("FLAGS", "DragC5"))
    'pluto:6.9
    UserList(Userindex).flags.DragCredito6 = val(leer.GetValue("FLAGS", "DragC6"))

    UserList(Userindex).flags.Elixir = val(leer.GetValue("FLAGS", "Elixir"))
    '---------------------

    UserList(Userindex).flags.Navegando = val(leer.GetValue("FLAGS", "Navegando"))
    UserList(Userindex).flags.Montura = val(leer.GetValue("FLAGS", "Montura"))
    UserList(Userindex).flags.ClaseMontura = val(leer.GetValue("FLAGS", "ClaseMontura"))
    UserList(Userindex).Counters.Pena = val(leer.GetValue("COUNTERS", "Pena"))
    UserList(Userindex).EmailActual = leer.GetValue("CONTACTO", "EmailActual")
    UserList(Userindex).Email = leer.GetValue("CONTACTO", "Email")
    UserList(Userindex).Remorted = leer.GetValue("INIT", "RAZAREMORT")
    'pluto:6.0A
    UserList(Userindex).BD = val(leer.GetValue("INIT", "BD"))

    UserList(Userindex).Genero = leer.GetValue("INIT", "Genero")
    UserList(Userindex).clase = leer.GetValue("INIT", "Clase")
    UserList(Userindex).raza = leer.GetValue("INIT", "Raza")
    UserList(Userindex).Hogar = leer.GetValue("INIT", "Hogar")
    UserList(Userindex).Char.Heading = val(leer.GetValue("INIT", "Heading"))
    UserList(Userindex).Esposa = Trim$(leer.GetValue("INIT", "Esposa"))
    UserList(Userindex).Paquete = 0
    'pluto:2.24-------------------------------
    'Dim filexx As String

    'If UserList(UserIndex).Esposa = "0" Then
    'filexx = "C:\Esposas\Charfile\" & Left$(UCase$(Name), 1) & "\" & UCase$(Name) & ".chr"
    'UserList(UserIndex).Esposa = GetVar(filexx, "INIT", "Esposa")
    'End If
    '-----------------------------------------

    UserList(Userindex).Nhijos = val(leer.GetValue("INIT", "Nhijos"))

    For loopc = 1 To 5
        UserList(Userindex).Hijo(loopc) = Trim$(leer.GetValue("INIT", "Hijo" & loopc))
    Next

    UserList(Userindex).Amor = val(leer.GetValue("INIT", "Amor"))
    UserList(Userindex).Embarazada = val(leer.GetValue("INIT", "Embarazada"))
    UserList(Userindex).Bebe = val(leer.GetValue("INIT", "Bebe"))
    UserList(Userindex).NombreDelBebe = Trim$(leer.GetValue("INIT", "NombreDelBebe"))
    UserList(Userindex).Padre = Trim$(leer.GetValue("INIT", "Padre"))
    UserList(Userindex).Madre = Trim$(leer.GetValue("INIT", "Madre"))
    UserList(Userindex).OrigChar.Head = val(leer.GetValue("INIT", "Head"))
    UserList(Userindex).OrigChar.Body = val(leer.GetValue("INIT", "Body"))
    UserList(Userindex).OrigChar.WeaponAnim = val(leer.GetValue("INIT", "Arma"))
    UserList(Userindex).OrigChar.ShieldAnim = val(leer.GetValue("INIT", "Escudo"))
    UserList(Userindex).OrigChar.CascoAnim = val(leer.GetValue("INIT", "Casco"))
    UserList(Userindex).OrigChar.Botas = val(leer.GetValue("INIT", "Botas"))
    UserList(Userindex).OrigChar.AlasAnim = val(leer.GetValue("INIT", "Alas"))
    UserList(Userindex).UltimoLogeo = val(leer.GetValue("INIT", "UltimoLogeo"))
    UserList(Userindex).UltimaDenuncia = val(leer.GetValue("INIT", "UltimaDenuncia"))
    UserList(Userindex).PrimeraDenuncia = val(leer.GetValue("INIT", "PrimeraDenuncia"))
    'UserList(UserIndex).Faccion.ArmadaReal = val(Leer.GetValue( "FACCIONES", "EjercitoReal"))
    'UserList(UserIndex).Faccion.FuerzasCaos = val(Leer.GetValue( "FACCIONES", "EjercitoCaos"))
    'UserList(UserIndex).Faccion.CiudadanosMatados = val(Leer.GetValue( "FACCIONES", "CiudMatados"))
    'UserList(UserIndex).Faccion.CriminalesMatados = val(Leer.GetValue( "FACCIONES", "CrimMatados"))
    'UserList(UserIndex).Faccion.RecibioArmaduraCaos = val(Leer.GetValue( "FACCIONES", "rArCaos"))
    'UserList(UserIndex).Faccion.RecibioArmaduraReal = val(Leer.GetValue( "FACCIONES", "rArReal"))
    'pluto:2.3
    'UserList(UserIndex).Faccion.RecibioArmaduraLegion = val(Leer.GetValue( "FACCIONES", "rArLegion"))
    'UserList(UserIndex).Faccion.RecibioExpInicialCaos = val(Leer.GetValue( "FACCIONES", "rExCaos"))
    'UserList(UserIndex).Faccion.RecibioExpInicialReal = val(Leer.GetValue( "FACCIONES", "rExReal"))

    'UserList(UserIndex).Faccion.RecompensasCaos = val(Leer.GetValue( "FACCIONES", "recCaos"))
    'UserList(UserIndex).Faccion.RecompensasReal = val(Leer.GetValue( "FACCIONES", "recReal"))

    'UserList(UserIndex).flags.Muerto = val(Leer.GetValue( "FLAGS", "Muerto"))
    'UserList(UserIndex).flags.Escondido = val(Leer.GetValue( "FLAGS", "Escondido"))

    'UserList(UserIndex).flags.Hambre = val(Leer.GetValue( "FLAGS", "Hambre"))
    'UserList(UserIndex).flags.Sed = val(Leer.GetValue( "FLAGS", "Sed"))
    'UserList(UserIndex).flags.Desnudo = val(Leer.GetValue( "FLAGS", "Desnudo"))

    '[Tite]Party
    UserList(Userindex).flags.party = False
    UserList(Userindex).flags.partyNum = 0
    UserList(Userindex).flags.invitado = ""
    '[\Tite]
    'UserList(UserIndex).flags.Envenenado = val(Leer.GetValue( "FLAGS", "Envenenado"))
    'UserList(UserIndex).flags.Morph = val(Leer.GetValue( "FLAGS", "Morph"))
    'UserList(UserIndex).flags.Paralizado = val(Leer.GetValue( "FLAGS", "Paralizado"))
    'UserList(UserIndex).flags.Angel = val(Leer.GetValue( "FLAGS", "Angel"))
    'UserList(UserIndex).flags.Demonio = val(Leer.GetValue( "FLAGS", "Demonio"))

    'UserList(UserIndex).flags.Navegando = val(Leer.GetValue( "FLAGS", "Navegando"))
    'pluto:2.3
    'UserList(UserIndex).flags.Montura = val(Leer.GetValue( "FLAGS", "Montura"))
    'UserList(UserIndex).flags.ClaseMontura = val(Leer.GetValue( "FLAGS", "ClaseMontura"))

    'UserList(UserIndex).Counters.Pena = val(Leer.GetValue( "COUNTERS", "Pena"))
    'pluto:2.10
    'UserList(UserIndex).EmailActual = Leer.GetValue( "CONTACTO", "EmailActual")

    'UserList(UserIndex).Email = Leer.GetValue( "CONTACTO", "Email")
    'UserList(UserIndex).Remorted = Leer.GetValue( "INIT", "RAZAREMORT")
    'UserList(UserIndex).Genero = Leer.GetValue( "INIT", "Genero")
    'UserList(UserIndex).clase = Leer.GetValue( "INIT", "Clase")
    'UserList(UserIndex).raza = Leer.GetValue( "INIT", "Raza")
    'UserList(UserIndex).Hogar = Leer.GetValue( "INIT", "Hogar")
    'UserList(UserIndex).Char.Heading = val(Leer.GetValue( "INIT", "Heading"))
    'pluto:2.14--------
    'UserList(UserIndex).Esposa = Leer.GetValue( "INIT", "Esposa")
    'UserList(UserIndex).Nhijos = val(Leer.GetValue( "INIT", "Nhijos"))
    'pluto:2.15
    'Dim X As Byte
    'For X = 1 To 5
    'UserList(UserIndex).Hijo(X) = Leer.GetValue( "INIT", "Hijo" & X)
    'Next
    'UserList(UserIndex).Amor = val(Leer.GetValue( "INIT", "Amor"))
    'UserList(UserIndex).Embarazada = val(Leer.GetValue( "INIT", "Embarazada"))
    'UserList(UserIndex).Bebe = val(Leer.GetValue( "INIT", "Bebe"))
    'UserList(UserIndex).NombreDelBebe = Leer.GetValue( "INIT", "NombreDelBebe")
    'UserList(UserIndex).Padre = Leer.GetValue( "INIT", "Padre")
    'UserList(UserIndex).Madre = Leer.GetValue( "INIT", "Madre")
    '-------------------

    'UserList(UserIndex).OrigChar.Head = val(Leer.GetValue( "INIT", "Head"))
    'UserList(UserIndex).OrigChar.Body = val(Leer.GetValue( "INIT", "Body"))
    'UserList(UserIndex).OrigChar.WeaponAnim = val(Leer.GetValue( "INIT", "Arma"))
    'UserList(UserIndex).OrigChar.ShieldAnim = val(Leer.GetValue( "INIT", "Escudo"))
    'UserList(UserIndex).OrigChar.CascoAnim = val(Leer.GetValue( "INIT", "Casco"))
    '[GAU]
    'UserList(UserIndex).OrigChar.Botas = val(Leer.GetValue( "INIT", "Botas"))
    '[GAU]
    UserList(Userindex).OrigChar.Heading = SOUTH

    If UserList(Userindex).flags.Muerto = 0 Then
        UserList(Userindex).Char = UserList(Userindex).OrigChar
    Else

        If Not Criminal(Userindex) Then UserList(Userindex).Char.Body = iCuerpoMuerto Else UserList( _
           Userindex).Char.Body = iCuerpoMuerto2

        If Not Criminal(Userindex) Then UserList(Userindex).Char.Head = iCabezaMuerto Else UserList( _
           Userindex).Char.Head = iCabezaMuerto2
        UserList(Userindex).Char.WeaponAnim = NingunArma
        UserList(Userindex).Char.ShieldAnim = NingunEscudo
        UserList(Userindex).Char.CascoAnim = NingunCasco
        '[GAU]
        UserList(Userindex).Char.Botas = NingunBota
        UserList(Userindex).Char.AlasAnim = NingunAla

        '[GAU]
    End If

    UserList(Userindex).Desc = Trim$(leer.GetValue("INIT", "Desc"))
    'UserList(UserIndex).Desc = Leer.GetValue("INIT", "Desc")

    UserList(Userindex).Pos.Map = val(ReadField(1, leer.GetValue("INIT", "Position"), 45))
    UserList(Userindex).Pos.X = val(ReadField(2, leer.GetValue("INIT", "Position"), 45))
    UserList(Userindex).Pos.Y = val(ReadField(3, leer.GetValue("INIT", "Position"), 45))
    'UserList(Userindex).Faccion.RecompensasReal = val(leer.GetValue("FACCIONES", "recReal"))

    'Delzak
    'If UserList(UserIndex).Pos.Map <> 0 Then Call BuscaPosicionValida(UserIndex)

    'UserList(UserIndex).Invent.NroItems = Leer.GetValue( "Inventory", "CantidadItems")
    UserList(Userindex).Invent.NroItems = leer.GetValue("Inventory", "CantidadItems")
    Dim loopd As Integer

    '[KEVIN]--------------------------------------------------------------------

    '***********************************************************************************
    'pluto:7.0 quito todo esto lo paso a cuentas

    'UserList(UserIndex).BancoInvent.NroItems = val(Leer.GetValue("BancoInventory", "CantidadItems"))

    'Lista de objetos del banco
    'For loopd = 1 To MAX_BANCOINVENTORY_SLOTS

    '   ln2 = Leer.GetValue("BancoInventory", "Obj" & loopd)

    '  UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex = val(ReadField(1, ln2, 45))
    ' UserList(UserIndex).BancoInvent.Object(loopd).Amount = val(ReadField(2, ln2, 45))
    'Next loopd
    '------------------------------------------------------------------------------------

    '[/KEVIN]*****************************************************************************

    'Lista de objetos
    For loopc = 1 To MAX_INVENTORY_SLOTS
        'ln = Leer.GetValue( "Inventory", "Obj" & LoopC)
        ln = leer.GetValue("Inventory", "Obj" & loopc)

        UserList(Userindex).Invent.Object(loopc).ObjIndex = val(ReadField(1, ln, 45))
        UserList(Userindex).Invent.Object(loopc).Amount = val(ReadField(2, ln, 45))
        UserList(Userindex).Invent.Object(loopc).Equipped = val(ReadField(3, ln, 45))
    Next loopc

    'Obtiene el indice-objeto del arma
    'UserList(UserIndex).Invent.WeaponEqpSlot = val(Leer.GetValue( "Inventory", "WeaponEqpSlot"))
    UserList(Userindex).Invent.WeaponEqpSlot = val(leer.GetValue("Inventory", "WeaponEqpSlot"))

    If UserList(Userindex).Invent.WeaponEqpSlot > 0 Then
        UserList(Userindex).Invent.WeaponEqpObjIndex = UserList(Userindex).Invent.Object(UserList( _
                                                                                         Userindex).Invent.WeaponEqpSlot).ObjIndex

    End If

    'Obtiene el indice-objeto del anillo
    'UserList(UserIndex).Invent.AnilloEqpSlot = val(Leer.GetValue( "Inventory", "AnilloEqpSlot"))
    UserList(Userindex).Invent.AnilloEqpSlot = val(leer.GetValue("Inventory", "AnilloEqpSlot"))

    If UserList(Userindex).Invent.AnilloEqpSlot > 0 Then
        UserList(Userindex).Invent.AnilloEqpObjIndex = UserList(Userindex).Invent.Object(UserList( _
                                                                                         Userindex).Invent.AnilloEqpSlot).ObjIndex

    End If

    'Obtiene el indice-objeto del armadura
    'UserList(UserIndex).Invent.ArmourEqpSlot = val(Leer.GetValue( "Inventory", "ArmourEqpSlot"))
    UserList(Userindex).Invent.ArmourEqpSlot = val(leer.GetValue("Inventory", "ArmourEqpSlot"))

    If UserList(Userindex).Invent.ArmourEqpSlot > 0 Then
        UserList(Userindex).Invent.ArmourEqpObjIndex = UserList(Userindex).Invent.Object(UserList( _
                                                                                         Userindex).Invent.ArmourEqpSlot).ObjIndex
        UserList(Userindex).flags.Desnudo = 0
    Else
        UserList(Userindex).flags.Desnudo = 1

    End If

    'Obtiene el indice-objeto del escudo
    'UserList(UserIndex).Invent.EscudoEqpSlot = val(Leer.GetValue( "Inventory", "EscudoEqpSlot"))
    UserList(Userindex).Invent.EscudoEqpSlot = val(leer.GetValue("Inventory", "EscudoEqpSlot"))

    If UserList(Userindex).Invent.EscudoEqpSlot > 0 Then
        UserList(Userindex).Invent.EscudoEqpObjIndex = UserList(Userindex).Invent.Object(UserList( _
                                                                                         Userindex).Invent.EscudoEqpSlot).ObjIndex

    End If

    'Obtiene el indice-objeto del casco
    'UserList(UserIndex).Invent.CascoEqpSlot = val(Leer.GetValue( "Inventory", "CascoEqpSlot"))
    UserList(Userindex).Invent.CascoEqpSlot = val(leer.GetValue("Inventory", "CascoEqpSlot"))

    If UserList(Userindex).Invent.CascoEqpSlot > 0 Then
        UserList(Userindex).Invent.CascoEqpObjIndex = UserList(Userindex).Invent.Object(UserList( _
                                                                                        Userindex).Invent.CascoEqpSlot).ObjIndex

    End If

    UserList(Userindex).Invent.AlaEqpSlot = val(leer.GetValue("Inventory", "AlaEqpSlot"))

    If UserList(Userindex).Invent.AlaEqpSlot > 0 Then
        UserList(Userindex).Invent.AlaEqpObjIndex = UserList(Userindex).Invent.Object(UserList( _
                                                                                      Userindex).Invent.AlaEqpSlot).ObjIndex

    End If

    '[GAU]
    'Obtiene el indice-objeto de las botas
    'UserList(UserIndex).Invent.BotaEqpSlot = val(Leer.GetValue( "Inventory", "BotaEqpSlot"))
    UserList(Userindex).Invent.BotaEqpSlot = val(leer.GetValue("Inventory", "BotaEqpSlot"))

    If UserList(Userindex).Invent.BotaEqpSlot > 0 Then
        UserList(Userindex).Invent.BotaEqpObjIndex = UserList(Userindex).Invent.Object(UserList( _
                                                                                       Userindex).Invent.BotaEqpSlot).ObjIndex

    End If

    '[GAU]
    'Obtiene el indice-objeto barco
    'UserList(UserIndex).Invent.BarcoSlot = val(Leer.GetValue( "Inventory", "BarcoSlot"))
    UserList(Userindex).Invent.BarcoSlot = val(leer.GetValue("Inventory", "BarcoSlot"))

    If UserList(Userindex).Invent.BarcoSlot > 0 Then
        UserList(Userindex).Invent.BarcoObjIndex = UserList(Userindex).Invent.Object(UserList( _
                                                                                     Userindex).Invent.BarcoSlot).ObjIndex

    End If

    'Obtiene el indice-objeto barco
    'UserList(UserIndex).Invent.MunicionEqpSlot = val(Leer.GetValue( "Inventory", "MunicionSlot"))
    UserList(Userindex).Invent.MunicionEqpSlot = val(leer.GetValue("Inventory", "MunicionSlot"))

    If UserList(Userindex).Invent.MunicionEqpSlot > 0 Then
        UserList(Userindex).Invent.MunicionEqpObjIndex = UserList(Userindex).Invent.Object(UserList( _
                                                                                           Userindex).Invent.MunicionEqpSlot).ObjIndex

    End If

    'UserList(UserIndex).NroMacotas = val(Leer.GetValue( "Mascotas", "NroMascotas"))
    UserList(Userindex).NroMacotas = val(leer.GetValue("Mascotas", "NroMascotas"))

    If UserList(Userindex).NroMacotas < 0 Then UserList(Userindex).NroMacotas = 0

    'Lista de objetos
    For loopc = 1 To MAXMASCOTAS
        ' UserList(UserIndex).MascotasType(LoopC) = val(Leer.GetValue( "Mascotas", "Mas" & LoopC))
        UserList(Userindex).MascotasType(loopc) = val(leer.GetValue("Mascotas", "Mas" & loopc))
    Next loopc

    'UserList(UserIndex).GuildInfo.FundoClan = val(Leer.GetValue( "Guild", "FundoClan"))
    'UserList(UserIndex).GuildInfo.EsGuildLeader = val(Leer.GetValue( "Guild", "EsGuildLeader"))
    'UserList(UserIndex).GuildInfo.Echadas = val(Leer.GetValue( "Guild", "Echadas"))
    'UserList(UserIndex).GuildInfo.Solicitudes = val(Leer.GetValue( "Guild", "Solicitudes"))
    'UserList(UserIndex).GuildInfo.SolicitudesRechazadas = val(Leer.GetValue( "Guild", "SolicitudesRechazadas"))
    'UserList(UserIndex).GuildInfo.VecesFueGuildLeader = val(Leer.GetValue( "Guild", "VecesFueGuildLeader"))
    'UserList(UserIndex).GuildInfo.YaVoto = val(Leer.GetValue( "Guild", "YaVoto"))
    'UserList(UserIndex).GuildInfo.ClanesParticipo = val(Leer.GetValue( "Guild", "ClanesParticipo"))
    'UserList(UserIndex).GuildInfo.GuildPoints = val(Leer.GetValue( "Guild", "GuildPts"))
    'UserList(UserIndex).GuildInfo.ClanFundado = Leer.GetValue( "Guild", "ClanFundado")
    'UserList(UserIndex).GuildInfo.GuildName = Leer.GetValue( "Guild", "GuildName")

    UserList(Userindex).GuildInfo.FundoClan = val(leer.GetValue("Guild", "FundoClan"))
    UserList(Userindex).GuildInfo.EsGuildLeader = val(leer.GetValue("Guild", "EsGuildLeader"))
    UserList(Userindex).GuildInfo.Echadas = val(leer.GetValue("Guild", "Echadas"))
    UserList(Userindex).GuildInfo.Solicitudes = val(leer.GetValue("Guild", "Solicitudes"))
    UserList(Userindex).GuildInfo.SolicitudesRechazadas = val(leer.GetValue("Guild", "SolicitudesRechazadas"))
    UserList(Userindex).GuildInfo.VecesFueGuildLeader = val(leer.GetValue("Guild", "VecesFueGuildLeader"))
    UserList(Userindex).GuildInfo.YaVoto = val(leer.GetValue("Guild", "Yavoto"))
    UserList(Userindex).GuildInfo.ClanesParticipo = val(leer.GetValue("Guild", "ClanesParticipo"))
    UserList(Userindex).GuildInfo.GuildPoints = val(leer.GetValue("Guild", "GuildPts"))
    UserList(Userindex).GuildInfo.ClanFundado = Trim$(leer.GetValue("Guild", "ClanFundado"))
    UserList(Userindex).GuildInfo.GuildName = Trim$(leer.GetValue("Guild", "GuildName"))

    'loaduserstats-------------------------------
    For loopc = 1 To NUMATRIBUTOS
        UserList(Userindex).Stats.UserAtributos(loopc) = leer.GetValue("ATRIBUTOS", "AT" & loopc)
        UserList(Userindex).Stats.UserAtributosBackUP(loopc) = UserList(Userindex).Stats.UserAtributos(loopc)
    Next
    'pluto:7.0
    UserList(Userindex).UserDañoProyetilesRaza = val(leer.GetValue("PORC", "P1"))
    UserList(Userindex).UserDañoArmasRaza = val(leer.GetValue("PORC", "P2"))
    UserList(Userindex).UserDañoMagiasRaza = val(leer.GetValue("PORC", "P3"))
    UserList(Userindex).UserDefensaMagiasRaza = val(leer.GetValue("PORC", "P4"))
    UserList(Userindex).UserEvasiónRaza = val(leer.GetValue("PORC", "P5"))
    UserList(Userindex).UserDefensaEscudos = val(leer.GetValue("PORC", "P6"))

    If UserList(Userindex).UserDañoProyetilesRaza + UserList(Userindex).UserDañoArmasRaza + UserList( _
       Userindex).UserDañoMagiasRaza + UserList(Userindex).UserDefensaMagiasRaza + UserList( _
       Userindex).UserEvasiónRaza + UserList(Userindex).UserDefensaEscudos > 15 Then
        UserList(Userindex).UserDañoArmasRaza = 5
        UserList(Userindex).UserDañoMagiasRaza = 5
        UserList(Userindex).UserDefensaMagiasRaza = 5

    End If

    For loopc = 1 To NUMSKILLS
        UserList(Userindex).Stats.UserSkills(loopc) = val(leer.GetValue("SKILLS", "SK" & loopc))
    Next

    For loopc = 1 To MAXUSERHECHIZOS
        UserList(Userindex).Stats.UserHechizos(loopc) = val(leer.GetValue("Hechizos", "H" & loopc))
    Next
    'pluto:2-3-04
    UserList(Userindex).Stats.Puntos = val(leer.GetValue("STATS", "PUNTOS"))

    UserList(Userindex).Stats.GLD = val(leer.GetValue("STATS", "GLD"))
    UserList(Userindex).Remort = val(leer.GetValue("STATS", "REMORT"))
    UserList(Userindex).Stats.Banco = val(leer.GetValue("STATS", "BANCO"))

    UserList(Userindex).Stats.MET = val(leer.GetValue("STATS", "MET"))
    UserList(Userindex).Stats.MaxHP = val(leer.GetValue("STATS", "MaxHP"))
    UserList(Userindex).Stats.MinHP = val(leer.GetValue("STATS", "MinHP"))

    UserList(Userindex).Stats.FIT = val(leer.GetValue("STATS", "FIT"))
    UserList(Userindex).Stats.MinSta = val(leer.GetValue("STATS", "MinSTA"))
    UserList(Userindex).Stats.MaxSta = val(leer.GetValue("STATS", "MaxSTA"))

    UserList(Userindex).Stats.MaxMAN = val(leer.GetValue("STATS", "MaxMAN"))
    UserList(Userindex).Stats.MinMAN = val(leer.GetValue("STATS", "MinMAN"))

    UserList(Userindex).Stats.MaxHIT = val(leer.GetValue("STATS", "MaxHIT"))
    UserList(Userindex).Stats.MinHIT = val(leer.GetValue("STATS", "MinHIT"))

    UserList(Userindex).Stats.MaxAGU = val(leer.GetValue("STATS", "MaxAGU"))
    UserList(Userindex).Stats.MinAGU = val(leer.GetValue("STATS", "MinAGU"))

    UserList(Userindex).Stats.MaxHam = val(leer.GetValue("STATS", "MaxHAM"))
    UserList(Userindex).Stats.MinHam = val(leer.GetValue("STATS", "MinHAM"))

    UserList(Userindex).Stats.SkillPts = val(leer.GetValue("STATS", "SkillPtsLibres"))

    UserList(Userindex).Stats.exp = val(leer.GetValue("STATS", "EXP"))
    UserList(Userindex).Stats.Elu = val(leer.GetValue("STATS", "ELU"))
    UserList(Userindex).Stats.Elo = val(leer.GetValue("STATS", "ELO"))
    UserList(Userindex).Stats.ELV = val(leer.GetValue("STATS", "ELV"))
    UserList(Userindex).Stats.LibrosUsados = val(leer.GetValue("STATS", "LIBROSUSADOS"))
    UserList(Userindex).Stats.Fama = val(leer.GetValue("STATS", "FAMA"))
    'pluto:2.4.5
    UserList(Userindex).Stats.PClan = val(leer.GetValue("STATS", "PCLAN"))
    UserList(Userindex).Stats.GTorneo = val(leer.GetValue("STATS", "GTORNEO"))

    UserList(Userindex).Stats.UsuariosMatados = val(leer.GetValue("MUERTES", "UserMuertes"))
    UserList(Userindex).Stats.CriminalesMatados = val(leer.GetValue("MUERTES", "CrimMuertes"))
    UserList(Userindex).Stats.NPCsMuertos = val(leer.GetValue("MUERTES", "NpcsMuertes"))
    '--------------------------------------------

    'Delzak-----------------------------------------
    '...............................................
    '              SISTEMA PREMIOS
    '...............................................
    '--Modificado por Pluto:7.0---------------------

    'Stats de premios por matar NPCs
    For loopc = 1 To 34
        UserList(Userindex).Stats.PremioNPC(loopc) = val(leer.GetValue("PREMIOS", "L" & loopc))
    Next
    '--------------------------------------------

    'PLUTO 6.0A  loadusermonturas ---------------------------

    UserList(Userindex).Nmonturas = val(leer.GetValue("MONTURAS", "NroMonturas"))

    Dim n As Byte

    For n = 1 To 3

        If val(leer.GetValue("MONTURA" & n, "TIPO")) > 0 Then
            loopc = val(leer.GetValue("MONTURA" & n, "TIPO"))

            UserList(Userindex).Montura.Tipo(loopc) = val(leer.GetValue("MONTURA" & n, "TIPO"))
            UserList(Userindex).Montura.Nivel(loopc) = val(leer.GetValue("MONTURA" & n, "NIVEL"))
            UserList(Userindex).Montura.exp(loopc) = val(leer.GetValue("MONTURA" & n, "EXP"))
            UserList(Userindex).Montura.Elu(loopc) = val(leer.GetValue("MONTURA" & n, "ELU"))
            UserList(Userindex).Montura.Vida(loopc) = val(leer.GetValue("MONTURA" & n, "VIDA"))
            UserList(Userindex).Montura.Golpe(loopc) = val(leer.GetValue("MONTURA" & n, "GOLPE"))
            UserList(Userindex).Montura.Nombre(loopc) = Trim$(leer.GetValue("MONTURA" & n, "NOMBRE"))
            UserList(Userindex).Montura.AtCuerpo(loopc) = val(leer.GetValue("MONTURA" & n, "ATCUERPO"))
            UserList(Userindex).Montura.Defcuerpo(loopc) = val(leer.GetValue("MONTURA" & n, "DEFCUERPO"))
            UserList(Userindex).Montura.AtFlechas(loopc) = val(leer.GetValue("MONTURA" & n, "ATFLECHAS"))
            UserList(Userindex).Montura.DefFlechas(loopc) = val(leer.GetValue("MONTURA" & n, "DEFFLECHAS"))
            UserList(Userindex).Montura.AtMagico(loopc) = val(leer.GetValue("MONTURA" & n, "ATMAGICO"))
            UserList(Userindex).Montura.DefMagico(loopc) = val(leer.GetValue("MONTURA" & n, "DEFMAGICO"))
            UserList(Userindex).Montura.Evasion(loopc) = val(leer.GetValue("MONTURA" & n, "EVASION"))
            UserList(Userindex).Montura.Libres(loopc) = val(leer.GetValue("MONTURA" & n, "LIBRES"))
            UserList(Userindex).Montura.index(loopc) = n

        End If

    Next n

    '---------------------------------------------

    'loaduserreputacion---------------------------
    UserList(Userindex).Reputacion.AsesinoRep = val(leer.GetValue("REP", "Asesino"))
    UserList(Userindex).Reputacion.BandidoRep = val(leer.GetValue("REP", "Dandido"))
    UserList(Userindex).Reputacion.BurguesRep = val(leer.GetValue("REP", "Burguesia"))
    UserList(Userindex).Reputacion.LadronesRep = val(leer.GetValue("REP", "Ladrones"))
    UserList(Userindex).Reputacion.NobleRep = val(leer.GetValue("REP", "Nobles"))
    UserList(Userindex).Reputacion.PlebeRep = val(leer.GetValue("REP", "Plebe"))
    UserList(Userindex).Reputacion.Promedio = val(leer.GetValue("REP", "Promedio"))

    'pluto:2-3-04
    'If UserList(Userindex).Faccion.FuerzasCaos > 0 And UserList(Userindex).Reputacion.Promedio >= 0 Then Call _
       ExpulsarCaos(Userindex)
    '------------------------------------------------------

    Call LoadQuestStats(Userindex, leer)

    Exit Sub
fallo:
    Call LogError("LOADUSERINIT" & Err.number & " D: " & Err.Description)

End Sub

Function GetVar(File As String, Main As String, Var As String) As String

    On Error GoTo fallo

    Dim sSpaces As String    ' This will hold the input that the program will retrieve
    Dim szReturn As String    ' This will be the defaul value if the string is not found

    szReturn = ""

    sSpaces = Space(5000)    ' This tells the computer how long the longest string can be

    GetPrivateProfileString Main, Var, szReturn, sSpaces, Len(sSpaces), File

    GetVar = RTrim(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    Exit Function
fallo:
    Call LogError("GETVAR" & Err.number & " D: " & Err.Description)

End Function

Sub LoadSini()

    On Error GoTo fallo

    Dim Temporal As Long
    Dim Temporal1 As Long

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."

    BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

    ServerIp = GetVar(IniPath & "Server.ini", "INIT", "ServerIp")
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = (mid(ServerIp, 1, Temporal - 1) And &H7F) * 16777216
    ServerIp = mid(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + mid(ServerIp, 1, Temporal - 1) * 65536
    ServerIp = mid(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + mid(ServerIp, 1, Temporal - 1) * 256
    ServerIp = mid(ServerIp, Temporal + 1, Len(ServerIp))

    Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
    HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
    AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
    IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
    'Lee la version correcta del cliente
    ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")
    'pluto:6.9
    TOPELANZAR = val(GetVar(IniPath & "Server.ini", "INIT", "AvisoLanzar"))
    TOPEFLECHA = val(GetVar(IniPath & "Server.ini", "INIT", "AvisoFlecha"))

    ArmaduraImperial1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial1"))
    ArmaduraImperial2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial2"))
    ArmaduraImperial3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial3"))
    TunicaMagoImperial = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperial"))
    TunicaMagoImperialEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperialEnanos"))

    ArmaduraCaos1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos1"))
    ArmaduraCaos2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos2"))
    ArmaduraCaos3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos3"))
    TunicaMagoCaos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaos"))
    TunicaMagoCaosEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaosEnanos"))
    'ropa legion
    ArmaduraLegion1 = val(GetVar(IniPath & "Server.ini", "INIT", "Armaduralegion1"))
    ArmaduraLegion2 = val(GetVar(IniPath & "Server.ini", "INIT", "Armaduralegion2"))
    ArmaduraLegion3 = val(GetVar(IniPath & "Server.ini", "INIT", "Armaduralegion3"))
    TunicaMagoLegion = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagolegion"))
    TunicaMagoLegionEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagolegionEnanos"))
    'castillos clanes
    castillo1 = GetVar(IniPath & "castillos.txt", "INIT", "Castillo1")
    castillo2 = GetVar(IniPath & "castillos.txt", "INIT", "Castillo2")
    castillo3 = GetVar(IniPath & "castillos.txt", "INIT", "Castillo3")
    castillo4 = GetVar(IniPath & "castillos.txt", "INIT", "Castillo4")
    fortaleza = GetVar(IniPath & "castillos.txt", "INIT", "fortaleza")
    'ciudades dueños
    'DueñoNix = val(GetVar(IniPath & "ciudades.txt", "INIT", "NIX"))
    'DueñoCaos = val(GetVar(IniPath & "ciudades.txt", "INIT", "CAOS"))
    'DueñoUlla = val(GetVar(IniPath & "ciudades.txt", "INIT", "ULLA"))
    'DueñoBander = val(GetVar(IniPath & "ciudades.txt", "INIT", "BANDER"))
    'DueñoDescanso = val(GetVar(IniPath & "ciudades.txt", "INIT", "DESCANSO"))
    'DueñoQuest = val(GetVar(IniPath & "ciudades.txt", "INIT", "QUEST"))
    'DueñoArghal = val(GetVar(IniPath & "ciudades.txt", "INIT", "ARGHAL"))
    'DueñoLaurana = val(GetVar(IniPath & "ciudades.txt", "INIT", "LAURANA"))
    'DueñoLindos = val(GetVar(IniPath & "ciudades.txt", "INIT", "LINDOS"))

    hora1 = GetVar(IniPath & "castillos.txt", "INIT", "hora1")
    hora2 = GetVar(IniPath & "castillos.txt", "INIT", "hora2")
    hora3 = GetVar(IniPath & "castillos.txt", "INIT", "hora3")
    hora4 = GetVar(IniPath & "castillos.txt", "INIT", "hora4")
    hora5 = GetVar(IniPath & "castillos.txt", "INIT", "hora5")
    date1 = GetVar(IniPath & "castillos.txt", "INIT", "date1")
    date2 = GetVar(IniPath & "castillos.txt", "INIT", "date2")
    date3 = GetVar(IniPath & "castillos.txt", "INIT", "date3")
    date4 = GetVar(IniPath & "castillos.txt", "INIT", "date4")
    date5 = GetVar(IniPath & "castillos.txt", "INIT", "date5")

    ClientsCommandsQueue = val(GetVar(IniPath & "Server.ini", "INIT", "ClientsCommandsQueue"))

    If ClientsCommandsQueue <> 0 Then
        frmMain.CmdExec.Enabled = True
    Else
        frmMain.CmdExec.Enabled = False

    End If

    'Start pos
    StartPos.Map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
    StartPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
    StartPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))

    'Intervalos
    SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
    FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar

    StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
    FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar

    SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
    FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar

    StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
    FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar

    IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
    FrmInterv.txtIntervaloSed.Text = IntervaloSed

    IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
    FrmInterv.txtIntervaloHambre.Text = IntervaloHambre

    IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
    FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno

    IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
    FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado

    IntervaloParalisisPJ = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalisisPJ"))
    'FrmInterv.txtIntervaloParalisisPJ.Text = IntervaloParalisisPJ
    IntervaloMorphPJ = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMorphPJ"))
    Intervaloceguera = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "Intervaloceguera"))
    'FrmInterv.txtIntervaloceguera.Text = Intervaloceguera

    IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
    FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible

    IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
    FrmInterv.txtIntervaloFrio.Text = IntervaloFrio

    IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
    FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx

    IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
    FrmInterv.txtInvocacion.Text = IntervaloInvocacion

    IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
    FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion

    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&

    IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
    FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear

    frmMain.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
    FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval

    frmMain.npcataca.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
    FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval

    IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
    FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar
    'pluto:2.8.0
    IntervaloUserPuedeFlechas = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechas"))
    FrmInterv.TxtFlechas.Text = IntervaloUserPuedeFlechas

    IntervaloRegeneraVampiro = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloRegeneraVampiro"))
    FrmInterv.txtVampire.Text = IntervaloRegeneraVampiro

    'pluto:2.10
    IntervaloUserPuedeTomar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeTomar"))

    IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
    FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar

    frmMain.tLluvia.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
    FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval

    frmMain.CmdExec.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTimerExec"))
    FrmInterv.txtCmdExec.Text = frmMain.CmdExec.Interval

    MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))

    If MinutosWs < 60 Then MinutosWs = 180

    IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))

    'Ressurect pos
    ResPos.Map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
    ResPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
    ResPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))

    recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))

    'Max users
    MaxUsers = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))

    'pluto:2.17
    TimeEmbarazo = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "TimeEmbarazo"))
    TimeAborto = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "TimeAborto"))
    ProbEmbarazo = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "ProbEmbarazo"))
    'pluto:6.0A
    NumeroGranPoder = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "NumeroGranPoder"))

    ReDim UserList(1 To MaxUsers) As User
    ReDim Cuentas(1 To MaxUsers)
    Call IniciaCuentas

    Nix.Map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
    Nix.X = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
    Nix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")

    Ullathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
    Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
    Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")

    Banderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
    Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
    Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")

    Lindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
    Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
    Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")

    ciudadcaos.Map = GetVar(DatPath & "Ciudades.dat", "CAOS", "Mapa")
    ciudadcaos.X = GetVar(DatPath & "Ciudades.dat", "CAOS", "X")
    ciudadcaos.Y = GetVar(DatPath & "Ciudades.dat", "CAOS", "Y")
    'pluto:2.17
    Pobladohumano.Map = GetVar(DatPath & "Ciudades.dat", "humano", "Mapa")
    Pobladohumano.X = GetVar(DatPath & "Ciudades.dat", "humano", "X")
    Pobladohumano.Y = GetVar(DatPath & "Ciudades.dat", "humano", "Y")
    Pobladoorco.Map = GetVar(DatPath & "Ciudades.dat", "orco", "Mapa")
    Pobladoorco.X = GetVar(DatPath & "Ciudades.dat", "orco", "X")
    Pobladoorco.Y = GetVar(DatPath & "Ciudades.dat", "orco", "Y")
    Pobladoenano.Map = GetVar(DatPath & "Ciudades.dat", "enano", "Mapa")
    Pobladoenano.X = GetVar(DatPath & "Ciudades.dat", "enano", "X")
    Pobladoenano.Y = GetVar(DatPath & "Ciudades.dat", "enano", "Y")
    Pobladoelfo.Map = GetVar(DatPath & "Ciudades.dat", "elfo", "Mapa")
    Pobladoelfo.X = GetVar(DatPath & "Ciudades.dat", "elfo", "X")
    Pobladoelfo.Y = GetVar(DatPath & "Ciudades.dat", "elfo", "Y")
    Pobladovampiro.Map = GetVar(DatPath & "Ciudades.dat", "vampiro", "Mapa")
    Pobladovampiro.X = GetVar(DatPath & "Ciudades.dat", "vampiro", "X")
    Pobladovampiro.Y = GetVar(DatPath & "Ciudades.dat", "vampiro", "Y")
    '-------------------------------

    'pluto:2.24------------------------------------
    WeB = GetVar(IniPath & "Server.ini", "INIT", "WebAodraG")
    DifServer = val(GetVar(IniPath & "Server.ini", "INIT", "DificultadServer"))
    DifOro = val(GetVar(IniPath & "Server.ini", "INIT", "DificultadOro"))
    BaseDatos = val(GetVar(IniPath & "Server.ini", "INIT", "BaseDatos"))
    ServerPrimario = val(GetVar(IniPath & "Server.ini", "INIT", "ServerPrimario"))
    NumeroObjEvento = val(GetVar(IniPath & "Server.ini", "EVENTOS", "NumeroObjEvento"))
    CantEntregarObjEvento = val(GetVar(IniPath & "Server.ini", "EVENTOS", "CantEntregarObjEvento"))
    CantObjRecompensa = val(GetVar(IniPath & "Server.ini", "EVENTOS", "CantObjRecompensa"))
    ObjRecompensaEventos(1) = val(GetVar(IniPath & "Server.ini", "EVENTOS", "ObjRecompensaEventos1"))
    ObjRecompensaEventos(2) = val(GetVar(IniPath & "Server.ini", "EVENTOS", "ObjRecompensaEventos2"))
    ObjRecompensaEventos(3) = val(GetVar(IniPath & "Server.ini", "EVENTOS", "ObjRecompensaEventos3"))
    ObjRecompensaEventos(4) = val(GetVar(IniPath & "Server.ini", "EVENTOS", "ObjRecompensaEventos4"))
    '------------------------------------------------

    'Call SQLConnect("localhost", "aodrag", "root", "")
    Call BDDConnect
    'Call BDDResetGMsos
    Call BDDSetUsersOnline

    Call BDDSetCastillos

    Exit Sub
fallo:
    Call LogError("LOADSINI" & Err.number & " D: " & Err.Description)

End Sub

Sub WriteVar(File As String, Main As String, Var As String, value As String)

'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************
    On Error GoTo fallo

    writeprivateprofilestring Main, Var, value, File
    Exit Sub
fallo:
    Call LogError("WRITEVAR" & Err.number & " D: " & Err.Description)

End Sub

Sub SaveUser(Userindex As Integer, UserFile As String)

    On Error GoTo errhandler

    'pluto:6.2------------------------------------------
    'Posicion de comienzo
    'Dim x As Integer
    'Dim Y As Integer
    'Dim Map As Integer

    'Select Case UserList(UserIndex).Pos.Map

    'Case MAPATORNEO 'torneos
    'If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
    'UserList(UserIndex).Pos.Map = 296
    'UserList(UserIndex).Pos.X = 71
    'UserList(UserIndex).Pos.Y = 64
    'Case MapaTorneo2 'torneos
    'If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
    'UserList(UserIndex).Pos.Map = 296
    'UserList(UserIndex).Pos.X = 71
    'UserList(UserIndex).Pos.Y = 64
    'Case 291 To 295 'torneos
    'If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
    'UserList(UserIndex).Pos.Map = 296
    'UserList(UserIndex).Pos.X = 71
    'UserList(UserIndex).Pos.Y = 64
    'Case 277 'fabrica lingotes
    'If UserList(UserIndex).Pos.X = 36 And UserList(UserIndex).Pos.Y = 70 Then UserList(UserIndex).Pos = Nix

    'Case 186 'fortaleza
    'If fortaleza <> UserList(UserIndex).GuildInfo.GuildName Then
    'If Not Criminal(UserIndex) Then UserList(UserIndex).Pos = Banderbill Else UserList(UserIndex).Pos = ciudadcaos
    'End If

    'Case 166 To 169 'castillos
    'UserList(UserIndex).Pos.X = 26 + RandomNumber(1, 9)
    'UserList(UserIndex).Pos.Y = 85 + RandomNumber(1, 5)

    'Case 191 To 192 'dragfutbol o bloqueo
    'UserList(UserIndex).Pos = Nix

    'End Select
    '---------------------------------

    If FileExist(UserFile, vbNormal) Then
        If UserList(Userindex).flags.Muerto = 1 Then UserList(Userindex).Char.Head = val(GetVar(UserFile, "INIT", _
                                                                                                "Head"))

        '       Kill UserFile
    End If

    'pluto:6.5 quito esto lo llevo a closeuser
    'If UserList(UserIndex).flags.Montura = 1 Then
    'Dim obj As ObjData
    'Call UsaMontura(UserIndex, obj)
    'End If
    Dim loopc As Integer

    Call WriteVar(UserFile, "FLAGS", "Muerto", val(UserList(Userindex).flags.Muerto))
    Call WriteVar(UserFile, "FLAGS", "LiderAlianza", val(UserList(Userindex).flags.LiderAlianza))
    Call WriteVar(UserFile, "FLAGS", "LiderHorda", val(UserList(Userindex).flags.LiderHorda))
    Call WriteVar(UserFile, "FLAGS", "Escondido", val(UserList(Userindex).flags.Escondido))
    Call WriteVar(UserFile, "FLAGS", "Hambre", val(UserList(Userindex).flags.Hambre))
    Call WriteVar(UserFile, "FLAGS", "Sed", val(UserList(Userindex).flags.Sed))
    Call WriteVar(UserFile, "FLAGS", "Desnudo", val(UserList(Userindex).flags.Desnudo))
    Call WriteVar(UserFile, "FLAGS", "Ban", val(UserList(Userindex).flags.ban))
    Call WriteVar(UserFile, "FLAGS", "Navegando", val(UserList(Userindex).flags.Navegando))
    'pluto:6.0A---------------
    Call WriteVar(UserFile, "FLAGS", "Minotauro", val(UserList(Userindex).flags.Minotauro))
    Call WriteVar(UserFile, "FLAGS", "MinOn", val(UserList(Userindex).flags.MinutosOnline))
    'pluto:7.0
    Call WriteVar(UserFile, "FLAGS", "Creditos", val(UserList(Userindex).flags.Creditos))

    Call WriteVar(UserFile, "FLAGS", "DragC1", val(UserList(Userindex).flags.DragCredito1))
    Call WriteVar(UserFile, "FLAGS", "DragC2", val(UserList(Userindex).flags.DragCredito2))
    Call WriteVar(UserFile, "FLAGS", "DragC3", val(UserList(Userindex).flags.DragCredito3))
    Call WriteVar(UserFile, "FLAGS", "DragC4", val(UserList(Userindex).flags.DragCredito4))
    Call WriteVar(UserFile, "FLAGS", "DragC5", val(UserList(Userindex).flags.DragCredito5))
    Call WriteVar(UserFile, "FLAGS", "DragC6", val(UserList(Userindex).flags.DragCredito6))

    Call WriteVar(UserFile, "FLAGS", "Elixir", val(UserList(Userindex).flags.Elixir))
    '--------------------------

    'pluto:2.3
    'Call WriteVar(UserFile, "FLAGS", "Montura", val(UserList(UserIndex).Flags.Montura))
    'Call WriteVar(UserFile, "FLAGS", "ClaseMontura", val(UserList(UserIndex).Flags.ClaseMontura))
    'pluto:2.4.1
    Call WriteVar(UserFile, "FLAGS", "Montura", 0)
    Call WriteVar(UserFile, "FLAGS", "ClaseMontura", 0)

    Call WriteVar(UserFile, "FLAGS", "Envenenado", val(UserList(Userindex).flags.Envenenado))
    Call WriteVar(UserFile, "FLAGS", "Paralizado", val(UserList(Userindex).flags.Paralizado))
    Call WriteVar(UserFile, "FLAGS", "Morph", val(UserList(Userindex).flags.Morph))

    Call WriteVar(UserFile, "FLAGS", "Angel", val(UserList(Userindex).flags.Angel))
    Call WriteVar(UserFile, "FLAGS", "Demonio", val(UserList(Userindex).flags.Demonio))

    Call WriteVar(UserFile, "COUNTERS", "Pena", val(UserList(Userindex).Counters.Pena))

    Call WriteVar(UserFile, "FACCIONES", "EjercitoReal", val(UserList(Userindex).Faccion.ArmadaReal))
    Call WriteVar(UserFile, "FACCIONES", "EjercitoCaos", val(UserList(Userindex).Faccion.FuerzasCaos))
    Call WriteVar(UserFile, "FACCIONES", "CiudMatados", val(UserList(Userindex).Faccion.CiudadanosMatados))
    Call WriteVar(UserFile, "FACCIONES", "NeutMatados", val(UserList(Userindex).Faccion.NeutralesMatados))
    Call WriteVar(UserFile, "FACCIONES", "CrimMatados", val(UserList(Userindex).Faccion.CriminalesMatados))
    Call WriteVar(UserFile, "FACCIONES", "rArCaos", val(UserList(Userindex).Faccion.RecibioArmaduraCaos))
    Call WriteVar(UserFile, "FACCIONES", "rArReal", val(UserList(Userindex).Faccion.RecibioArmaduraReal))
    'pluto:2.3
    Call WriteVar(UserFile, "FACCIONES", "rArLegion", val(UserList(Userindex).Faccion.RecibioArmaduraLegion))
    Call WriteVar(UserFile, "FACCIONES", "rExCaos", val(UserList(Userindex).Faccion.RecibioExpInicialCaos))
    Call WriteVar(UserFile, "FACCIONES", "rExReal", val(UserList(Userindex).Faccion.RecibioExpInicialReal))
    Call WriteVar(UserFile, "FACCIONES", "recCaos", val(UserList(Userindex).Faccion.RecompensasCaos))
    Call WriteVar(UserFile, "FACCIONES", "recReal", val(UserList(Userindex).Faccion.RecompensasReal))

    Call WriteVar(UserFile, "GUILD", "EsGuildLeader", val(UserList(Userindex).GuildInfo.EsGuildLeader))
    Call WriteVar(UserFile, "GUILD", "Echadas", val(UserList(Userindex).GuildInfo.Echadas))
    Call WriteVar(UserFile, "GUILD", "Solicitudes", val(UserList(Userindex).GuildInfo.Solicitudes))
    Call WriteVar(UserFile, "GUILD", "SolicitudesRechazadas", val(UserList(Userindex).GuildInfo.SolicitudesRechazadas))
    Call WriteVar(UserFile, "GUILD", "VecesFueGuildLeader", val(UserList(Userindex).GuildInfo.VecesFueGuildLeader))
    Call WriteVar(UserFile, "GUILD", "YaVoto", val(UserList(Userindex).GuildInfo.YaVoto))
    Call WriteVar(UserFile, "GUILD", "FundoClan", val(UserList(Userindex).GuildInfo.FundoClan))
    'pluto:2.4.5
    Call WriteVar(UserFile, "STATS", "PClan", val(UserList(Userindex).Stats.PClan))
    Call WriteVar(UserFile, "STATS", "GTorneo", val(UserList(Userindex).Stats.GTorneo))

    Call WriteVar(UserFile, "GUILD", "GuildName", UserList(Userindex).GuildInfo.GuildName)
    Call WriteVar(UserFile, "GUILD", "ClanFundado", UserList(Userindex).GuildInfo.ClanFundado)
    Call WriteVar(UserFile, "GUILD", "ClanesParticipo", str(UserList(Userindex).GuildInfo.ClanesParticipo))
    Call WriteVar(UserFile, "GUILD", "GuildPts", str(UserList(Userindex).GuildInfo.GuildPoints))

    '¿Fueron modificados los atributos del usuario?
    If Not UserList(Userindex).flags.TomoPocion Then

        For loopc = 1 To UBound(UserList(Userindex).Stats.UserAtributos)
            Call WriteVar(UserFile, "ATRIBUTOS", "AT" & loopc, val(UserList(Userindex).Stats.UserAtributos(loopc)))
        Next
    Else

        For loopc = 1 To UBound(UserList(Userindex).Stats.UserAtributos)
            UserList(Userindex).Stats.UserAtributos(loopc) = UserList(Userindex).Stats.UserAtributosBackUP(loopc)
            Call WriteVar(UserFile, "ATRIBUTOS", "AT" & loopc, val(UserList(Userindex).Stats.UserAtributos(loopc)))
        Next

    End If

    'pluto:7.0
    Call WriteVar(UserFile, "PORC", "P1", str(UserList(Userindex).UserDañoProyetilesRaza))
    Call WriteVar(UserFile, "PORC", "P2", str(UserList(Userindex).UserDañoArmasRaza))
    Call WriteVar(UserFile, "PORC", "P3", str(UserList(Userindex).UserDañoMagiasRaza))
    Call WriteVar(UserFile, "PORC", "P4", str(UserList(Userindex).UserDefensaMagiasRaza))
    Call WriteVar(UserFile, "PORC", "P5", str(UserList(Userindex).UserEvasiónRaza))
    Call WriteVar(UserFile, "PORC", "P6", str(UserList(Userindex).UserDefensaEscudos))

    For loopc = 1 To UBound(UserList(Userindex).Stats.UserSkills)
        Call WriteVar(UserFile, "SKILLS", "SK" & loopc, val(UserList(Userindex).Stats.UserSkills(loopc)))
    Next

    Call WriteVar(UserFile, "CONTACTO", "Email", UserList(Userindex).Email)
    'pluto:2.10
    Call WriteVar(UserFile, "CONTACTO", "EmailActual", Cuentas(Userindex).mail)
    

    Call WriteVar(UserFile, "INIT", "Genero", UserList(Userindex).Genero)
    Call WriteVar(UserFile, "INIT", "Raza", UserList(Userindex).raza)
    Call WriteVar(UserFile, "INIT", "Hogar", UserList(Userindex).Hogar)
    Call WriteVar(UserFile, "INIT", "Clase", UserList(Userindex).clase)
    Call WriteVar(UserFile, "INIT", "Desc", UserList(Userindex).Desc)
    Call WriteVar(UserFile, "INIT", "Heading", str(UserList(Userindex).Char.Heading))
    Call WriteVar(UserFile, "INIT", "Head", str(UserList(Userindex).OrigChar.Head))

    If UserList(Userindex).flags.Muerto = 0 Then
        Call WriteVar(UserFile, "INIT", "Body", str(UserList(Userindex).Char.Body))

    End If

    If UserList(Userindex).flags.Morph > 0 Then
        Call WriteVar(UserFile, "INIT", "Body", str(UserList(Userindex).flags.Morph))

    End If

    If UserList(Userindex).flags.Angel > 0 Then
        Call WriteVar(UserFile, "INIT", "Body", str(UserList(Userindex).flags.Angel))

    End If

    If UserList(Userindex).flags.Demonio > 0 Then
        Call WriteVar(UserFile, "INIT", "Body", str(UserList(Userindex).flags.Demonio))

    End If

    Call WriteVar(UserFile, "INIT", "Arma", str(UserList(Userindex).Char.WeaponAnim))
    Call WriteVar(UserFile, "INIT", "Escudo", str(UserList(Userindex).Char.ShieldAnim))
    Call WriteVar(UserFile, "INIT", "Casco", str(UserList(Userindex).Char.CascoAnim))
    '[GAU]
    Call WriteVar(UserFile, "INIT", "Botas", str(UserList(Userindex).Char.Botas))
    Call WriteVar(UserFile, "INIT", "Alas", str(UserList(Userindex).Char.AlasAnim))
    '[GAU]
    Call WriteVar(UserFile, "INIT", "RAZAREMORT", UserList(Userindex).Remorted)
    Call WriteVar(UserFile, "INIT", "BD", val(UserList(Userindex).BD))

    Call WriteVar(UserFile, "INIT", "LastIP", UserList(Userindex).ip)
    'pluto:2.14
    Call WriteVar(UserFile, "INIT", "LastSerie", UserList(Userindex).Serie)
    Call WriteVar(UserFile, "INIT", "LastMac", UserList(Userindex).MacPluto)
    Call WriteVar(UserFile, "INIT", "UltimoLogeo", UserList(Userindex).UltimoLogeo)
    Call WriteVar(UserFile, "INIT", "UltimaDenuncia", UserList(Userindex).UltimaDenuncia)
    Call WriteVar(UserFile, "INIT", "PrimeraDenuncia", UserList(Userindex).PrimeraDenuncia)

    'Debug.Print userfile

    'pluto:6.5---------
    'If UserList(UserIndex).Pos.Map = 170 Or UserList(UserIndex).Pos.Map = 34 Then
    'If UserList(UserIndex).Pos.X > 16 And UserList(UserIndex).Pos.X < 31 And UserList(UserIndex).Pos.Y > 42 And UserList(UserIndex).Pos.Y < 48 Then
    'UserList(UserIndex).Pos.X = 36
    'UserList(UserIndex).Pos.Y = 36
    'End If
    'End If
    '------------------
    Call WriteVar(UserFile, "INIT", "Position", UserList(Userindex).Pos.Map & "-" & UserList(Userindex).Pos.X & "-" & _
                                                UserList(Userindex).Pos.Y)

    ' pluto:2.15 -------------------
    Call WriteVar(UserFile, "INIT", "Esposa", UserList(Userindex).Esposa)
    Call WriteVar(UserFile, "INIT", "Nhijos", val(UserList(Userindex).Nhijos))
    Dim X As Byte

    For X = 1 To 5
        Call WriteVar(UserFile, "INIT", "Hijo" & X, UserList(Userindex).Hijo(X))
    Next
    Call WriteVar(UserFile, "INIT", "Amor", val(UserList(Userindex).Amor))
    Call WriteVar(UserFile, "INIT", "Embarazada", val(UserList(Userindex).Embarazada))
    Call WriteVar(UserFile, "INIT", "Bebe", val(UserList(Userindex).Bebe))
    Call WriteVar(UserFile, "INIT", "NombreDelBebe", UserList(Userindex).NombreDelBebe)
    Call WriteVar(UserFile, "INIT", "Padre", UserList(Userindex).Padre)
    Call WriteVar(UserFile, "INIT", "Madre", UserList(Userindex).Madre)
    '-----------------------------------

    'PLUTO:2-3-04
    Call WriteVar(UserFile, "STATS", "PUNTOS", str(UserList(Userindex).Stats.Puntos))

    Call WriteVar(UserFile, "STATS", "GLD", str(UserList(Userindex).Stats.GLD))
    Call WriteVar(UserFile, "STATS", "REMORT", str(UserList(Userindex).Remort))
    Call WriteVar(UserFile, "STATS", "BANCO", str(UserList(Userindex).Stats.Banco))

    Call WriteVar(UserFile, "STATS", "MET", str(UserList(Userindex).Stats.MET))
    Call WriteVar(UserFile, "STATS", "MaxHP", str(UserList(Userindex).Stats.MaxHP))
    Call WriteVar(UserFile, "STATS", "MinHP", str(UserList(Userindex).Stats.MinHP))

    Call WriteVar(UserFile, "STATS", "FIT", str(UserList(Userindex).Stats.FIT))
    Call WriteVar(UserFile, "STATS", "MaxSTA", str(UserList(Userindex).Stats.MaxSta))
    Call WriteVar(UserFile, "STATS", "MinSTA", str(UserList(Userindex).Stats.MinSta))

    Call WriteVar(UserFile, "STATS", "MaxMAN", str(UserList(Userindex).Stats.MaxMAN))
    Call WriteVar(UserFile, "STATS", "MinMAN", str(UserList(Userindex).Stats.MinMAN))

    Call WriteVar(UserFile, "STATS", "MaxHIT", str(UserList(Userindex).Stats.MaxHIT))
    Call WriteVar(UserFile, "STATS", "MinHIT", str(UserList(Userindex).Stats.MinHIT))

    Call WriteVar(UserFile, "STATS", "MaxAGU", str(UserList(Userindex).Stats.MaxAGU))
    Call WriteVar(UserFile, "STATS", "MinAGU", str(UserList(Userindex).Stats.MinAGU))

    Call WriteVar(UserFile, "STATS", "MaxHAM", str(UserList(Userindex).Stats.MaxHam))
    Call WriteVar(UserFile, "STATS", "MinHAM", str(UserList(Userindex).Stats.MinHam))

    Call WriteVar(UserFile, "STATS", "SkillPtsLibres", str(UserList(Userindex).Stats.SkillPts))

    Call WriteVar(UserFile, "STATS", "EXP", str(UserList(Userindex).Stats.exp))
    Call WriteVar(UserFile, "STATS", "ELV", str(UserList(Userindex).Stats.ELV))
    Call WriteVar(UserFile, "STATS", "ELU", str(UserList(Userindex).Stats.Elu))
    Call WriteVar(UserFile, "STATS", "ELO", str(UserList(Userindex).Stats.Elo))
    'pluto:6.0A
    Call WriteVar(UserFile, "STATS", "LIBROSUSADOS", str(UserList(Userindex).Stats.LibrosUsados))
    Call WriteVar(UserFile, "STATS", "FAMA", str(UserList(Userindex).Stats.Fama))

    Call WriteVar(UserFile, "MUERTES", "UserMuertes", val(UserList(Userindex).Stats.UsuariosMatados))
    Call WriteVar(UserFile, "MUERTES", "CrimMuertes", val(UserList(Userindex).Stats.CriminalesMatados))
    Call WriteVar(UserFile, "MUERTES", "NpcsMuertes", val(UserList(Userindex).Stats.NPCsMuertos))

    '[KEVIN]----------------------------------------------------------------------------
    '*******************************************************************************************

    'pluto:7.0 quito esto que pasa a sistema cuentas
    'Call WriteVar(userfile, "BancoInventory", "CantidadItems", val(UserList(UserIndex).BancoInvent.NroItems))
    'Dim loopd As Integer
    'pluto:7.0
    'For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
    '   Call WriteVar(userfile, "BancoInventory", "Obj" & loopd, UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex & "-" & UserList(UserIndex).BancoInvent.Object(loopd).Amount)
    'Next loopd
    '*******************************************************************************************
    '[/KEVIN]-----------

    'Save Inv
    Call WriteVar(UserFile, "Inventory", "CantidadItems", val(UserList(Userindex).Invent.NroItems))

    For loopc = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(UserFile, "Inventory", "Obj" & loopc, UserList(Userindex).Invent.Object(loopc).ObjIndex & "-" & _
                                                            UserList(Userindex).Invent.Object(loopc).Amount & "-" & UserList(Userindex).Invent.Object( _
                                                            loopc).Equipped)
    Next

    Call WriteVar(UserFile, "Inventory", "WeaponEqpSlot", str(UserList(Userindex).Invent.WeaponEqpSlot))
    Call WriteVar(UserFile, "Inventory", "ArmourEqpSlot", str(UserList(Userindex).Invent.ArmourEqpSlot))
    Call WriteVar(UserFile, "Inventory", "CascoEqpSlot", str(UserList(Userindex).Invent.CascoEqpSlot))
    Call WriteVar(UserFile, "Inventory", "EscudoEqpSlot", str(UserList(Userindex).Invent.EscudoEqpSlot))
    Call WriteVar(UserFile, "Inventory", "BarcoSlot", str(UserList(Userindex).Invent.BarcoSlot))
    Call WriteVar(UserFile, "Inventory", "MunicionSlot", str(UserList(Userindex).Invent.MunicionEqpSlot))
    'pluto:2.4.1
    Call WriteVar(UserFile, "Inventory", "AnilloEqpSlot", str(UserList(Userindex).Invent.AnilloEqpSlot))

    '[GAU]
    Call WriteVar(UserFile, "Inventory", "BotaEqpSlot", str(UserList(Userindex).Invent.BotaEqpSlot))
    Call WriteVar(UserFile, "Inventory", "AlaEqpSlot", str(UserList(Userindex).Invent.AlaEqpSlot))
    '[GAU]

    'Reputacion
    Call WriteVar(UserFile, "REP", "Asesino", val(UserList(Userindex).Reputacion.AsesinoRep))
    Call WriteVar(UserFile, "REP", "Bandido", val(UserList(Userindex).Reputacion.BandidoRep))
    Call WriteVar(UserFile, "REP", "Burguesia", val(UserList(Userindex).Reputacion.BurguesRep))
    Call WriteVar(UserFile, "REP", "Ladrones", val(UserList(Userindex).Reputacion.LadronesRep))
    Call WriteVar(UserFile, "REP", "Nobles", val(UserList(Userindex).Reputacion.NobleRep))
    Call WriteVar(UserFile, "REP", "Plebe", val(UserList(Userindex).Reputacion.PlebeRep))

    Dim l As Long
    l = (-UserList(Userindex).Reputacion.AsesinoRep) + (-UserList(Userindex).Reputacion.BandidoRep) + UserList( _
        Userindex).Reputacion.BurguesRep + (-UserList(Userindex).Reputacion.LadronesRep) + UserList( _
        Userindex).Reputacion.NobleRep + UserList(Userindex).Reputacion.PlebeRep
    l = l / 6
    Call WriteVar(UserFile, "REP", "Promedio", val(l))

    Dim cad As String

    For loopc = 1 To MAXUSERHECHIZOS
        cad = UserList(Userindex).Stats.UserHechizos(loopc)
        Call WriteVar(UserFile, "HECHIZOS", "H" & loopc, cad)
    Next

    Call SaveQuestStats(Userindex, UserFile)

    For loopc = 1 To MAXMASCOTAS

        ' Mascota valida?
        If UserList(Userindex).MascotasIndex(loopc) > 0 Then

            ' Nos aseguramos que la criatura no fue invocada
            If Npclist(UserList(Userindex).MascotasIndex(loopc)).Contadores.TiempoExistencia = 0 Then
                cad = UserList(Userindex).MascotasType(loopc)
            Else    'Si fue invocada no la guardamos
                cad = "0"
                UserList(Userindex).NroMacotas = UserList(Userindex).NroMacotas - 1

            End If

            Call WriteVar(UserFile, "MASCOTAS", "MAS" & loopc, 0)

        End If
        
        If UserList(Userindex).NroMacotas < 0 Then UserList(Userindex).NroMacotas = 0

    Next

    Call WriteVar(UserFile, "MASCOTAS", "NroMascotas", 0)

    'pluto:6.0A -guardamos mascotas
    Call WriteVar(UserFile, "MONTURAS", "NroMonturas", val(UserList(Userindex).Nmonturas))

    loopc = 0
    Dim n As Byte

    For n = 1 To 12

        loopc = UserList(Userindex).Montura.index(n)

        If loopc > 0 Then
            Call WriteVar(UserFile, "MONTURA" & loopc, "TIPO", val(UserList(Userindex).Montura.Tipo(n)))

            Call WriteVar(UserFile, "MONTURA" & loopc, "NIVEL", val(UserList(Userindex).Montura.Nivel(n)))
            Call WriteVar(UserFile, "MONTURA" & loopc, "EXP", val(UserList(Userindex).Montura.exp(n)))
            Call WriteVar(UserFile, "MONTURA" & loopc, "ELU", val(UserList(Userindex).Montura.Elu(n)))
            Call WriteVar(UserFile, "MONTURA" & loopc, "VIDA", val(UserList(Userindex).Montura.Vida(n)))
            Call WriteVar(UserFile, "MONTURA" & loopc, "GOLPE", val(UserList(Userindex).Montura.Golpe(n)))
            Call WriteVar(UserFile, "MONTURA" & loopc, "NOMBRE", UserList(Userindex).Montura.Nombre(n))

            Call WriteVar(UserFile, "MONTURA" & loopc, "ATCUERPO", val(UserList(Userindex).Montura.AtCuerpo(n)))
            Call WriteVar(UserFile, "MONTURA" & loopc, "DEFCUERPO", val(UserList(Userindex).Montura.Defcuerpo(n)))
            Call WriteVar(UserFile, "MONTURA" & loopc, "ATFLECHAS", val(UserList(Userindex).Montura.AtFlechas(n)))
            Call WriteVar(UserFile, "MONTURA" & loopc, "DEFFLECHAS", val(UserList(Userindex).Montura.DefFlechas(n)))
            Call WriteVar(UserFile, "MONTURA" & loopc, "ATMAGICO", val(UserList(Userindex).Montura.AtMagico(n)))
            Call WriteVar(UserFile, "MONTURA" & loopc, "DEFMAGICO", val(UserList(Userindex).Montura.DefMagico(n)))
            Call WriteVar(UserFile, "MONTURA" & loopc, "EVASION", val(UserList(Userindex).Montura.Evasion(n)))
            Call WriteVar(UserFile, "MONTURA" & loopc, "LIBRES", val(UserList(Userindex).Montura.Libres(n)))

        End If

    Next

    'Delzak sistema premios
    For n = 1 To 34
        Call WriteVar(UserFile, "PREMIOS", "L" & n, val(UserList(Userindex).Stats.PremioNPC(n)))
    Next

    Exit Sub
errhandler:
    Call LogError("Error en SaveUser")

End Sub

Function Criminal(ByVal Userindex As Integer) As Boolean

    On Error GoTo fallo

    'Dim a As Integer
    'If UserList(UserIndex).Reputacion.Promedio < 0 Then a = 1 Else a = 0
    'Dim l As Long
    'l = (-UserList(Userindex).Reputacion.AsesinoRep) + (-UserList(Userindex).Reputacion.BandidoRep) + UserList( _
     '   Userindex).Reputacion.BurguesRep + (-UserList(Userindex).Reputacion.LadronesRep) + UserList( _
      '  Userindex).Reputacion.NobleRep + UserList(Userindex).Reputacion.PlebeRep
    'l = l / 6
    'Criminal = (l < 0)
    'UserList(Userindex).Reputacion.Promedio = l
    If UserList(Userindex).Faccion.FuerzasCaos = 1 Then
    Criminal = True
    Else
    Criminal = False
    End If
    'If a = 0 And Criminal = True Then UserCrimi = UserCrimi + 1: UserCiu = UserCiu - 1
    'If a = 1 And Criminal = False Then UserCiu = UserCiu + 1: UserCrimi = UserCrimi - 1
    Exit Function
fallo:
    Call LogError("CRIMINAL " & Err.number & " D: " & Err.Description)

End Function

Sub BackUPnPc(NpcIndex As Integer)

    On Error GoTo fallo

    'Call LogTarea("Sub BackUPnPc NpcIndex:" & NpcIndex)

    Dim NpcNumero As Integer
    Dim npcfile As String
    Dim loopc As Integer

    NpcNumero = Npclist(NpcIndex).numero

    If NpcNumero > 499 Then
        npcfile = DatPath & "bkNPCs-HOSTILES.dat"
    Else
        npcfile = DatPath & "bkNPCs.dat"

    End If

    'General
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).Name)

    Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).Desc)
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(Npclist(NpcIndex).Char.Head))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(Npclist(NpcIndex).Char.Body))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.Heading))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(Npclist(NpcIndex).Movement))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(Npclist(NpcIndex).Attackable))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(Npclist(NpcIndex).Comercia))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(Npclist(NpcIndex).TipoItems))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(Npclist(NpcIndex).GiveEXP))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(Npclist(NpcIndex).GiveGLD))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Inflacion", val(Npclist(NpcIndex).Inflacion))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(Npclist(NpcIndex).InvReSpawn))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(Npclist(NpcIndex).NPCtype))
    'pluto:6.0A
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Arquero", val(Npclist(NpcIndex).Arquero))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Anima", val(Npclist(NpcIndex).Anima))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Raid", val(Npclist(NpcIndex).Raid))
    'pluto:7.0
    Call WriteVar(npcfile, "NPC" & NpcNumero, "LogroTipo", val(Npclist(NpcIndex).LogroTipo))

    'Stats
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.Def))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHIT))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHP))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHIT))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHP))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.UsuariosMatados))

    'Flags
    Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).flags.Respawn))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(Npclist(NpcIndex).flags.BackUp))
    Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).flags.Domable))

    'Inventario
    Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))

    If Npclist(NpcIndex).Invent.NroItems > 0 Then

        For loopc = 1 To MAX_INVENTORY_SLOTS
            Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & loopc, Npclist(NpcIndex).Invent.Object(loopc).ObjIndex _
                                                                     & "-" & Npclist(NpcIndex).Invent.Object(loopc).Amount)
        Next

    End If

    Exit Sub
fallo:
    Call LogError("BACKUPNPC" & Err.number & " D: " & Err.Description)

End Sub

Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)

'eze
    On Local Error Resume Next
    'eze

    On Error GoTo fallo

    'Call LogTarea("Sub CargarNpcBackUp NpcIndex:" & NpcIndex & " NpcNumber:" & NpcNumber)

    'Status
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"

    Dim npcfile As String

    If NpcNumber > 499 Then
        npcfile = DatPath & "bkNPCs-HOSTILES.dat"
    Else
        npcfile = DatPath & "bkNPCs.dat"

    End If

    Npclist(NpcIndex).numero = NpcNumber
    'pluto:2.17
    Npclist(NpcIndex).Anima = val(GetVar(npcfile, "NPC" & NpcNumber, "Anima"))
    Npclist(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
    Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
    Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
    Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))
    'pluto:6.0A
    Npclist(NpcIndex).Arquero = val(GetVar(npcfile, "NPC" & NpcNumber, "Arquero"))
    Npclist(NpcIndex).Raid = val(GetVar(npcfile, "NPC" & NpcNumber, "Raid"))
    'pluto:7.0
    Npclist(NpcIndex).LogroTipo = val(GetVar(npcfile, "NPC" & NpcNumber, "LogroTipo"))

    Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
    'EZE
    Npclist(NpcIndex).Char.ShieldAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "EscudoAnim"))
    Npclist(NpcIndex).Char.WeaponAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "ArmaAnim"))
    Npclist(NpcIndex).Char.CascoAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "CascoAnim"))
    'EZE
    Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
    Npclist(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))
    Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
    Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
    Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
    Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))

    Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))
    Npclist(NpcIndex).QuestNumber = val(GetVar(npcfile, "NPC" & NpcNumber, "QuestNumber"))

    Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

    '@Nati: NPCS vida a 1
    'Npclist(NpcIndex).Stats.MaxHP = 1
    'Npclist(NpcIndex).Stats.MinHP = 1
    Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
    Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
    Npclist(NpcIndex).Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
    Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
    Npclist(NpcIndex).Stats.Def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
    Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
    Npclist(NpcIndex).Stats.ImpactRate = val(GetVar(npcfile, "NPC" & NpcNumber, "ImpactRate"))
    'Npclist(NpcIndex).Premio = val(GetVar(npcfile, "NPC" & NpcNumber, "Premio")) 'Delzak sistema premios

    Dim loopc As Integer
    Dim ln As String
    Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))

    If Npclist(NpcIndex).Invent.NroItems > 0 Then

        For loopc = 1 To MAX_INVENTORY_SLOTS
            ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & loopc)
            Npclist(NpcIndex).Invent.Object(loopc).ObjIndex = val(ReadField(1, ln, 45))
            Npclist(NpcIndex).Invent.Object(loopc).Amount = val(ReadField(2, ln, 45))

        Next loopc

    Else

        For loopc = 1 To MAX_INVENTORY_SLOTS
            Npclist(NpcIndex).Invent.Object(loopc).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(loopc).Amount = 0
        Next loopc

    End If

    Npclist(NpcIndex).Inflacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Inflacion"))

    Npclist(NpcIndex).flags.NPCActive = True
    Npclist(NpcIndex).flags.UseAINow = False
    Npclist(NpcIndex).flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
    Npclist(NpcIndex).flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
    Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
    Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "PosOrig"))

    'Tipo de items con los que comercia
    Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))
    Exit Sub
fallo:
    Call LogError("CARGARNPCBACKUP" & Err.number & " D: " & Err.Description)

End Sub

Sub LogBan(ByVal BannedIndex As Integer, _
           ByVal Userindex As Integer, _
           ByVal moTivo As String)

    On Error GoTo fallo

    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "BannedBy", UserList( _
                                                                                                 Userindex).Name)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "Reason", moTivo)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "Fecha", Date)

    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, UserList(BannedIndex).Name
    Close #mifile
    Exit Sub
fallo:
    Call LogError("LOGBAN" & Err.number & " D: " & Err.Description)

End Sub

Private Sub BuscaPosicionValida(Userindex As Integer)

'Delzak (28-8-10)

    Dim leer As New clsIniManager
    Dim Mapa As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim MapaOK As Integer
    Dim XOK As Integer
    Dim YOK As Integer
    Dim dn As Integer
    Dim M As Integer
    Dim User As Integer
    Dim iNDiCe As Integer
    Dim QueSumo As Boolean    '0 para x, 1 para y
    Dim PosicionValida As Boolean
    Dim ControlBordes As Boolean

    Mapa = UserList(Userindex).Pos.Map
    X = UserList(Userindex).Pos.X
    Y = UserList(Userindex).Pos.Y
    MapaOK = Mapa
    XOK = X
    YOK = Y
    QueSumo = False
    iNDiCe = 1
    ControlBordes = True

    'Busco un hueco donde no haya nadie y que no este bloqueado (OPTIMIZADO 14-9-10)

    For dn = 1 To 6400    '80x80

        PosicionValida = True

        'Compruebo que no haya nadie en la posicion que quiero logear
        For User = 1 To LastUser

            If UserList(User).Pos.Map = MapaOK And UserList(User).Pos.X = XOK And UserList(User).Pos.Y = YOK Then _
               PosicionValida = False
        Next

        'Compruebo que no este bloqueado
        If PosicionValida = True Then

            If MapData(MapaOK, XOK, YOK).Blocked = 1 Then PosicionValida = False

        End If

        'Si la posicion es valida, salgo del bucle
        If PosicionValida = True And ControlBordes = True Then Exit For

        'Si no es valida, busco una trazando un espiral

        If QueSumo = False Then

            XOK = XOK + iNDiCe

        Else

            YOK = YOK + iNDiCe

            iNDiCe = iNDiCe * (-1)

            If iNDiCe < 0 Then iNDiCe = iNDiCe - 1
            If iNDiCe > 0 Then iNDiCe = iNDiCe + 1

        End If

        If QueSumo = True Then QueSumo = False
        If QueSumo = False Then QueSumo = True

        'Controlo que no me salga del borde
        If XOK < 4 Or XOK > 85 Or YOK < 4 Or YOK > 85 Then ControlBordes = False Else ControlBordes = True

        'Si termina el bucle y no he encontrado alternativa, que le den por culo
        If dn = 6400 Then
            MapaOK = Mapa
            XOK = X
            YOK = Y

        End If

    Next

    'Bloqueo la posicion donde voy a aparecer para que no me de por culo nadie
    MapData(MapaOK, XOK, YOK).Blocked = 1

    'Cargo mi posicion
    UserList(Userindex).Pos.Map = MapaOK
    UserList(Userindex).Pos.X = XOK
    UserList(Userindex).Pos.Y = YOK

    'Desbloqueo la posicion
    MapData(MapaOK, XOK, YOK).Blocked = 0

End Sub
