Attribute VB_Name = "Ventanas"
'----------------------------------------------------------------------------------
'                          �Heracles 2005
'----------------------------------------------------------------------------------
Option Explicit

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength _
                          Lib "user32" _
                              Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText _
                          Lib "user32" _
                              Alias "GetWindowTextA" (ByVal hwnd As Long, _
                                                      ByVal lpString As String, _
                                                      ByVal cch As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

' GetWindow() Constants
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&

Private Declare Function GetWindow _
                          Lib "user32" (ByVal hwnd As Long, _
                                        ByVal wFlag As Long) As Long

'
Private Declare Function FindWindow _
                          Lib "user32" _
                              Alias "FindWindowA" (ByVal lpClassName As String, _
                                                   ByVal lpWindowName As String) As Long
Private Declare Function SendMessage _
                          Lib "user32" _
                              Alias "SendMessageA" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long

Private Const SC_MINIMIZE = &HF020&
Private Const SC_CLOSE = &HF060&
Private Const WM_SYSCOMMAND = &H112
Private Const WM_CLOSE = &H10

Private Declare Function GetClassName _
                          Lib "user32" _
                              Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                     ByVal lpClassName As String, _
                                                     ByVal nMaxCount As Long) As Long

Public Sub CloseApp(ByVal xx As Long)
'Cerrar la ventana indicada, mediante el men� del sistema (o de windows)
'Esto funcionar� si la aplicaci�n tiene men� de sistema
'(aunque lo he probado con una utilidad sin controlBox y la cierra bien)
'
'Si se especifica ClassName, se cerrar�n la ventana si es de ese ClassName
'
'Dim hWnd As Long

'No cerrar la ventana "Progman"

    Call SendMessage(xx, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&)

End Sub

Private Function WindowTitle(ByVal hwnd As Long) As String
'Devuelve el t�tulo de una ventana, seg�n el hWnd indicado
'
    Dim sTitulo As String
    Dim lenTitulo As Long
    Dim ret As Long

    'Leer la longitud del t�tulo de la ventana
    lenTitulo = GetWindowTextLength(hwnd)

    If lenTitulo > 0 Then
        lenTitulo = lenTitulo + 1
        sTitulo = String$(lenTitulo, 0)
        'Leer el t�tulo de la ventana
        ret = GetWindowText(hwnd, sTitulo, lenTitulo)
        WindowTitle = Left$(sTitulo, ret)

    End If

End Function

Public Function EWindos()

    Dim sClase As String
    Dim aqui9 As String
    Dim sTitulo As String
    Dim hwnd As Long
    'Dim col As Collection
    Dim Bu As Integer
    Dim n As Integer
    Dim aa As String
    'Set col = New Collection
    Dim Stitula As String
    Dim ab As Long
    Stitula = ""
    aqui9 = ""
    Bu = 0
    Static VentP(100) As String
    Static Venta(100) As String

    'Primera ventana
    hwnd = GetWindow(GetDesktopWindow(), GW_CHILD)

    'Recorrer el resto de las ventanas
    Do While hwnd <> 0&
        'Si la ventana es visible
        sTitulo = WindowTitle(hwnd)

        If IsWindowVisible(hwnd) Then

            'Leer el caption de la ventana
            sTitulo = WindowTitle(hwnd)
            sClase = ClassName(sTitulo)

            If Len(sTitulo) Then

                For n = 1 To Len(sTitulo)
                    aa = mid$(sTitulo, n, 1)

                    If aa = "," Then Mid$(sTitulo, n, Len(sTitulo)) = " "
                Next n

                'A�adimos el t�tulo

                Bu = Bu + 1
                ' aqui9 = aqui9 & sTitulo & "," & sClase & "," & hWnd & ","
                VentP(Bu) = sTitulo

                'pluto:6.2
                If InStr(1, VentP(Bu), "Cheat Engine") Then
                    Dim Nus As Integer
                    Nus = Val(GetVar(App.Path & "\Init\Update.ini", "FICHERO", "z"))
                    Nus = Nus + 1
                    Call WriteVar(App.Path & "\Init\Update.ini", "FICHERO", "z", Val(Nus))
                    MsgBox _
                            "Cheat Engine Detectado!!. Ha quedado registrado este intento de usar un Cheat, si vemos que vuelves intentarlo ser�n baneados todos tus personajes. �� ESTAS AVISADO !!"
                    End

                End If

            End If

        End If

        'Siguiente ventana
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop

    'Set EWindows = col
    If Venta(1) = "" Then GoTo fue
    Dim GUAU As String
    Dim Igu As Byte
    Dim nn As Byte

    For n = 1 To 100
        For nn = 1 To 100

            If VentP(n) = Venta(nn) Then Igu = 1: Exit For
        Next nn

        'If VentP(n) = "Introduce los datos de tu cuenta" Then GoTo nu
        'If VentP(n) = "Cliente AoDraG" Then GoTo nu
        'If VentP(n) = "AoDraG Online" Then GoTo nu
        'If Left(VentP(n), 11) = "Servidor Ao" Then GoTo nu
        If Igu = 0 Then GUAU = GUAU + VentP(n)
nu:
        Igu = 0
    Next n

    If GUAU <> "" Then

        LogInicial = LogInicial & GUAU & ","

    End If

    'SendData ("TA1" & x & "," & Bu & "," & aqui9)

    'MsgBox (Stitula)
    'Form2.Show
fue:

    For n = 1 To 100
        Venta(n) = VentP(n)
    Next

End Function

Public Function EnumTopWindows(X As Integer)
'Enumera las ventanas que tienen t�tulo y son visibles
'Devuelve un array del tipo Variant con los nombres de las ventanas
'y su hWnd
'Por tanto la forma de acceder a este array ser�a:
'   Set col = EnumTopWindows
'   numItems = col.Count
'   For i = 1 To numItems Step 2
'       With List2
'           .AddItem col.Item(i)
'           .ItemData(.NewIndex) = col.Item(i + 1)
'       End With
'   Next
'
'Opcionalemente se puede especificar como par�metro un ListBox o ComboBox
'y los datos se a�adir�n a ese control
    Dim sClase As String
    Dim aqui9 As String
    Dim sTitulo As String
    Dim hwnd As Long
    Dim col As Collection
    Dim Bu As Integer
    Dim n As Integer
    Dim aa As String
    Set col = New Collection
    Dim Stitula As String
    Dim ab As Long
    Stitula = ""
    aqui9 = ""
    Bu = 0
    'If Not unListBox Is Nothing Then
    'unListBox.Clear
    'End If

    'Primera ventana
    hwnd = GetWindow(GetDesktopWindow(), GW_CHILD)

    'Recorrer el resto de las ventanas
    Do While hwnd <> 0&
        'Si la ventana es visible
        sTitulo = WindowTitle(hwnd)

        'If Not IsWindowVisible(hWnd) And sTitulo <> "" Then
        'ab = ab + 1
        'Form2.List1.AddItem ab & " - " & sTitulo
        'Stitula = Stitula + sTitulo
        ' End If
        If IsWindowVisible(hwnd) Then

            'Leer el caption de la ventana
            sTitulo = WindowTitle(hwnd)
            sClase = ClassName(sTitulo)

            If Len(sTitulo) Then

                For n = 1 To Len(sTitulo)
                    aa = mid$(sTitulo, n, 1)

                    If aa = "," Then Mid$(sTitulo, n, Len(sTitulo)) = " "
                Next n

                'A�adimos el t�tulo
                'col.Add sTitulo
                Bu = Bu + 1
                aqui9 = aqui9 & sTitulo & "," & sClase & "," & hwnd & ","

                'y el hWnd por si fuese �til
                ' col.Add hWnd
                'Si se especifica el ListBox
                'If Not unListBox Is Nothing Then
                ' With unListBox
                '.AddItem sTitulo
                '.ItemData(.NewIndex) = hWnd
                'End With
                ' End If
            End If

        End If

        'Siguiente ventana
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop
    Set EnumTopWindows = col

    SendData ("TA1" & X & "," & Bu & "," & aqui9 & ":::::LOGINICIAL:::::," & LogInicial)

    'MsgBox (Stitula)
    'Form2.Show
End Function

Public Function ClassName(ByVal Title As String) As String
'Devuelve el ClassName de una ventana, indicando el t�tulo de la misma
    Dim hwnd As Long
    Dim sClassName As String
    Dim nMaxCount As Long

    hwnd = FindWindow(sClassName, Title)

    nMaxCount = 256
    sClassName = Space$(nMaxCount)
    nMaxCount = GetClassName(hwnd, sClassName, nMaxCount)
    ClassName = Left$(sClassName, nMaxCount)

End Function

Public Sub MinimizeAll(Optional ClassName As String)
'Minimizar todas las ventanas
'
' Dim col As Collection
'Dim numItems As Long
' Dim i As Long
' Dim sTitulo As String
'Dim hWnd As Long

'Set col = New Collection

'Set col = Me.EnumTopWindows
'numItems = col.Count
'For i = 1 To numItems Step 2
'sTitulo = col.Item(i)
' hWnd = FindWindow(ClassName, sTitulo)
'hWnd = col.Item(i + 1)
'Call SendMessage(hWnd, WM_SYSCOMMAND, SC_MINIMIZE, ByVal 0&)
'Next

'Set col = Nothing
End Sub

'############################################################
'#                                                          #
'#                             Delzak)                      #
'#  Vamos a copiar a un archivo una rama dada del registro  #
'#    y despues vamos a leer una entrada en ese archivo     #
'#                                                          #
'############################################################

'Copiamos el archivo
Public Sub Backup_Reg()

    Dim Rama As String
    Dim pathArchivoReg As String
    pathArchivoReg = App.Path & "\INIT\Librerias.ini"

    'pluto:6.9
    If vWin = "Windows Vista" Then
        Rama = "HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache"
    Else
        Rama = "HKEY_CURRENT_USER\Software\Microsoft\Windows\ShellNoRoam\MUICache"

    End If

    'pluto:6.9----
    If FileExist(pathArchivoReg, vbArchive) Then
        Kill pathArchivoReg

    End If

    '--------------------
    ' si no existe el .reg lo crea, por que si no el _
      api GetShortPathName devuelve una cadena nula cuando un archivo no existe

    If Len(Dir(pathArchivoReg)) = 0 Then
        Open pathArchivoReg For Output As #1
        Close

    End If

    ' exporta la rama con el par�metro /e en la ruta indicada
    Shell "regedit /e " & Chr(34) & pathArchivoReg & Chr(34) & " " & Chr(34) & Rama & Chr(34)

    'Me.MousePointer = 0

    If Err.Number <> 0 Then
        Call LogError("RegistroCool: Err.Number <> 0")

        If Len(Dir(pathArchivoReg)) <> 0 Then
            Kill pathArchivoReg

        End If

    End If

End Sub

Public Function BorraEntrada()

    On Error Resume Next

    If vWin = "Windows Vista" Then
        Shell _
                "reg delete HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache /va /f"
    Else
        Shell "reg delete HKEY_CURRENT_USER\Software\Microsoft\Windows\ShellNoRoam\MUICache /va /f"

    End If

End Function

'Delzak)
'Buscamos dentro del archivo, la entrada WPE PRO y una vez leido, borramos el archivo y el registro
Public Function Buscawpe(ConLoG As Boolean)

    On Error Resume Next

    'Dim FilePath As String
    Dim data As String
    Dim Todo As String
    Dim MiObjeto As Object
    Dim Archum As Integer
    Dim filepath As String
    Archum = FreeFile()
    'Name App.Path & "\Init\Librerias.reg" As App.Path & "\Init\Librerias.ini"
    'Debug.Print ARCHNUM
    filepath = App.Path & "\Init\Librerias.ini"
    Wpe = False
    Bengi = False
    Open filepath For Input As #Archum

    '#     'carga el contenido del archivo en la variable
    '#     Contenido = Input$(LOF(F), #F)
    Do Until EOF(Archum)
        Line Input #1, data
        Todo = Todo & data & ","

        If data <> "" Then
            If InStr(data, "WPE") > 0 Then
                Wpe = True

            End If

            If InStr(data, "ENGINE") > 0 Then
                Bengi = True

            End If

        End If

        'data = Right$(data, 8)
        'data = Left$(data, 7)
        'wpe = False
        'If data = "WPE PRO" Then wpe = True
    Loop

    Close #1

    WpeLen = Len(Todo)
    Dim veces As Byte
    Dim n As Byte
    Dim Trozo As String

    If ConLoG = True Then
        TodoListado = Todo
        'veces = Int(Len(Todo) / 500) + 1
        'For n = 1 To veces
        'Trozo = Mid$(Todo, ((n - 1) * 500) + 1, 500)
        'SendData ("PSS" & Trozo)
        'Sleep 500
        'Else
        'Next
        frmMain.Wpetimer.Enabled = True

    End If

    'Borramos el archivo
    Kill App.Path & "\Init\Librerias.ini"
    Todo = ""

    'Dim clavita As String
    'clavita = "HKEY_CURRENT_USER\Software\Microsoft\Windows\ShellNoRoam\MUICache"
    'Borramos la clave del registro
    'Set MiObjeto = CreateObject("Wscript.Shell")
    'Shell "reg delete HKEY_CURRENT_USER\Software\Microsoft\Windows\ShellNoRoam\MUICache /va /f"
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\ShellNoRoam\MUICache
    'MiObjeto.regdelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\ShellNoRoam\MUICache" / va

    'Set MiObjeto = Nothing

End Function

