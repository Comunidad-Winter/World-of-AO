Attribute VB_Name = "CargaObjetos"
Option Explicit

'BORRAR
Public DataTrabajo()  As Integer
Public NumTrabajo     As Integer

Sub CargamosHechizos()

    On Error GoTo Errorfeo

    Hechizos(1).Nombre = "Curar veneno"
    Hechizos(1).Desc = "Cura el envenenamiento"
    Hechizos(1).PalabrasMagicas = "NIHIL VED"
    Hechizos(1).HechizeroMsg = "Has curado a"
    Hechizos(1).PropioMsg = "Te has curado del envenenamiento."
    Hechizos(1).TargetMsg = "te ha curado del envenenamiento."
    Hechizos(1).CuraVeneno = 1
    Hechizos(1).FXgrh = 107
    Hechizos(1).Loops = 1
    Hechizos(1).ManaRequerido = 10
    Hechizos(1).MinSkill = 8
    Hechizos(1).Target = 1
    Hechizos(1).Tipo = 2
    Hechizos(1).WAV = 16
    Hechizos(2).Nombre = "Dardo M�gico"
    Hechizos(2).Desc = "Causa 5 a 10 puntos de da�o a la v�ctima."
    Hechizos(2).PalabrasMagicas = "OHL VOR PEK"
    Hechizos(2).HechizeroMsg = "Has lanzado dardo m�gico sobre"
    Hechizos(2).PropioMsg = ""
    Hechizos(2).TargetMsg = "lanzo dardo m�gico sobre t�."
    Hechizos(2).FXgrh = 122
    Hechizos(2).Loops = 1
    Hechizos(2).ManaRequerido = 12
    Hechizos(2).MaxHP = 10
    Hechizos(2).MinHP = 5
    Hechizos(2).MinSkill = 6
    Hechizos(2).Resis = 1
    Hechizos(2).SubeHP = 2
    Hechizos(2).Target = 3
    Hechizos(2).Tipo = 1
    Hechizos(2).WAV = 16
    Hechizos(3).Nombre = "Curar Heridas Leves"
    Hechizos(3).Desc = "Curar heridas leves, restaura entre 1 y 5 puntos de salud."
    Hechizos(3).PalabrasMagicas = "CORP SANC"
    Hechizos(3).HechizeroMsg = "Has sanado a"
    Hechizos(3).PropioMsg = "Te has curado algunas heridas."
    Hechizos(3).TargetMsg = "te ha curado algunas heridas."
    Hechizos(3).FXgrh = 9
    Hechizos(3).Loops = 1
    Hechizos(3).ManaRequerido = 10
    Hechizos(3).MaxHP = 5
    Hechizos(3).MinHP = 1
    Hechizos(3).MinSkill = 10
    Hechizos(3).SubeHP = 1
    Hechizos(3).Target = 3
    Hechizos(3).Tipo = 1
    Hechizos(3).WAV = 17
    Hechizos(4).Nombre = "Envenenar"
    Hechizos(4).Desc = "Envenenamiento. Provoca la muerte si no se contraresta el veneno."
    Hechizos(4).PalabrasMagicas = "SERP XON IN"
    Hechizos(4).HechizeroMsg = "Has envenenado a"
    Hechizos(4).PropioMsg = "Te has envenenado."
    Hechizos(4).TargetMsg = "te ha envenenado."
    Hechizos(4).Envenena = 5
    Hechizos(4).FXgrh = 30
    Hechizos(4).Loops = 3
    Hechizos(4).ManaRequerido = 20
    Hechizos(4).MinSkill = 20
    Hechizos(4).Resis = 1
    Hechizos(4).Target = 3
    Hechizos(4).Tipo = 2
    Hechizos(4).WAV = 16
    Hechizos(5).Nombre = "Curar heridas graves"
    Hechizos(5).Desc = "Curar heridas graves, restaura entre 15 y 25 puntos de salud."
    Hechizos(5).PalabrasMagicas = "EN CORP SANCTIS"
    Hechizos(5).HechizeroMsg = "Has sanado a"
    Hechizos(5).PropioMsg = "Te has curado algunas heridas."
    Hechizos(5).TargetMsg = "te ha curado algunas heridas."
    Hechizos(5).FXgrh = 9
    Hechizos(5).Loops = 1
    Hechizos(5).ManaRequerido = 40
    Hechizos(5).MaxHP = 25
    Hechizos(5).MinHP = 15
    Hechizos(5).MinSkill = 38
    Hechizos(5).SubeHP = 1
    Hechizos(5).Target = 3
    Hechizos(5).Tipo = 1
    Hechizos(5).WAV = 18
    Hechizos(6).Nombre = "Flecha m�gica"
    Hechizos(6).Desc = "Causa 6 a 12 puntos de da�o a la victima."
    Hechizos(6).PalabrasMagicas = "VAX PER"
    Hechizos(6).HechizeroMsg = "Has lanzado flecha magica sobre "
    Hechizos(6).PropioMsg = "Has lanzado flecha magica sobre t�."
    Hechizos(6).TargetMsg = "lanzo flecha magica sobre t�."
    Hechizos(6).ManaRequerido = 20
    Hechizos(6).MaxHP = 15
    Hechizos(6).MinHP = 10
    Hechizos(6).MinSkill = 12
    Hechizos(6).Resis = 1
    Hechizos(6).SubeHP = 2
    Hechizos(6).Target = 3
    Hechizos(6).Tipo = 1
    Hechizos(6).WAV = 19
    Hechizos(7).Nombre = "Flecha el�ctrica"
    Hechizos(7).Desc = "Causa 15 a 25 puntos de da�o a la victima."
    Hechizos(7).PalabrasMagicas = "SUN VAP"
    Hechizos(7).HechizeroMsg = "Has lanzado flecha electrica sobre"
    Hechizos(7).PropioMsg = "Has lanzado flecha electrica sobre t�."
    Hechizos(7).TargetMsg = "lanzo flecha electrica sobre t�."
    Hechizos(7).FXgrh = 11
    Hechizos(7).ManaRequerido = 35
    Hechizos(7).MaxHP = 25
    Hechizos(7).MinHP = 15
    Hechizos(7).MinSkill = 22
    Hechizos(7).Resis = 1
    Hechizos(7).SubeHP = 2
    Hechizos(7).Target = 3
    Hechizos(7).Tipo = 1
    Hechizos(7).WAV = 16
    Hechizos(8).Nombre = "Misil m�gico"
    Hechizos(8).Desc = "Causa 30 a 37 puntos de da�o a la victima."
    Hechizos(8).PalabrasMagicas = "VAX IN TAR"
    Hechizos(8).HechizeroMsg = "Has lanzado misil magico sobre"
    Hechizos(8).PropioMsg = "Has lanzado misil magico sobre t�."
    Hechizos(8).TargetMsg = "lanzo misil magico sobre t�."
    Hechizos(8).FXgrh = 10
    Hechizos(8).Loops = 1
    Hechizos(8).ManaRequerido = 50
    Hechizos(8).MaxHP = 37
    Hechizos(8).MinHP = 30
    Hechizos(8).MinSkill = 38
    Hechizos(8).Resis = 1
    Hechizos(8).SubeHP = 2
    Hechizos(8).Target = 3
    Hechizos(8).Tipo = 1
    Hechizos(8).WAV = 16
    Hechizos(9).Nombre = "Paralizar"
    Hechizos(9).Desc = "Paraliza por un momento a la v�ctima."
    Hechizos(9).PalabrasMagicas = "HOAX VORP"
    Hechizos(9).HechizeroMsg = "Has paralizado a"
    Hechizos(9).PropioMsg = "Te has paralizado."
    Hechizos(9).TargetMsg = "te ha paralizado."
    Hechizos(9).FXgrh = 129
    Hechizos(9).Loops = 1
    Hechizos(9).ManaRequerido = 450
    Hechizos(9).MinSkill = 60
    Hechizos(9).Paraliza = 1
    Hechizos(9).Target = 3
    Hechizos(9).Tipo = 2
    Hechizos(9).WAV = 16
    Hechizos(10).Nombre = "Remover Par�lisis"
    Hechizos(10).Desc = "Remueve la par�lisis."
    Hechizos(10).PalabrasMagicas = "AN HOAX VORP"
    Hechizos(10).HechizeroMsg = "Le has removido la par�lisis a"
    Hechizos(10).PropioMsg = "Te has removido la par�lisis."
    Hechizos(10).TargetMsg = "te ha removido la par�lisis."
    Hechizos(10).FXgrh = 123
    Hechizos(10).ManaRequerido = 300
    Hechizos(10).MinSkill = 45
    Hechizos(10).RemoverParalisis = 1
    Hechizos(10).Target = 3
    Hechizos(10).Tipo = 2
    Hechizos(10).WAV = 16
    Hechizos(11).Nombre = "Resucitar"
    Hechizos(11).Desc = "Resucitar un usuario muerto."
    Hechizos(11).PalabrasMagicas = "AHIL KN� X�R"
    Hechizos(11).HechizeroMsg = "Has resucitado a"
    Hechizos(11).PropioMsg = "Te has resucitado."
    Hechizos(11).TargetMsg = "te ha resucitado."
    Hechizos(11).FXgrh = 72
    Hechizos(11).Loops = 1
    Hechizos(11).ManaRequerido = 420
    Hechizos(11).MinSkill = 75
    Hechizos(11).Revivir = 1
    Hechizos(11).Target = 1
    Hechizos(11).Tipo = 2
    Hechizos(11).WAV = 20
    Hechizos(12).Nombre = "Provocar hambre"
    Hechizos(12).Desc = "Provocar hambre, provoca la perdida de entre 20 y 50 pts de comida."
    Hechizos(12).PalabrasMagicas = "�L AEX"
    Hechizos(12).HechizeroMsg = "Le has lanzado hambre a"
    Hechizos(12).PropioMsg = "Te has lanzado el hechizo hambre."
    Hechizos(12).TargetMsg = "te ha lanzado el hechizo hambre."
    Hechizos(12).FXgrh = 28
    Hechizos(12).Loops = 1
    Hechizos(12).ManaRequerido = 20
    Hechizos(12).MaxHam = 50
    Hechizos(12).MinHam = 20
    Hechizos(12).MinSkill = 5
    Hechizos(12).Resis = 1
    Hechizos(12).SubeHam = 2
    Hechizos(12).Target = 1
    Hechizos(12).Tipo = 1
    Hechizos(12).WAV = 16
    Hechizos(13).Nombre = "Terrible hambre de Ig�r"
    Hechizos(13).Desc = _
    "Terrible hambre de Ig�r, provoca dejar a la v�ctima con s�lo 1 punto en hambre. Este encantamiento"
    Hechizos(13).PalabrasMagicas = "�X'�L AEX"
    Hechizos(13).HechizeroMsg = "Le has lanzado Hambre de Igor a"
    Hechizos(13).PropioMsg = "Te has lanzado el hechizo Hambre de Igor."
    Hechizos(13).TargetMsg = "te ha lanzado el hechizo Hambre de Igor."
    Hechizos(13).FXgrh = 28
    Hechizos(13).Loops = 1
    Hechizos(13).ManaRequerido = 55
    Hechizos(13).MaxHam = 100
    Hechizos(13).MinHam = 100
    Hechizos(13).MinSkill = 35
    Hechizos(13).Resis = 1
    Hechizos(13).SubeHam = 2
    Hechizos(13).Target = 1
    Hechizos(13).Tipo = 1
    Hechizos(13).WAV = 16
    Hechizos(14).Nombre = "Invisibilidad"
    Hechizos(14).Desc = "Vuelve invisible al target, efecto no permanente."
    Hechizos(14).PalabrasMagicas = "ROHL UX MAIO"
    Hechizos(14).HechizeroMsg = "Has lanzado invisibilidad sobre"
    Hechizos(14).PropioMsg = "Te has lanzado invisibilidad."
    Hechizos(14).TargetMsg = "lanzo invisibilidad sobre t�."
    Hechizos(14).Invisibilidad = 1
    Hechizos(14).ManaRequerido = 500
    Hechizos(14).MinSkill = 87
    Hechizos(14).Target = 1
    Hechizos(14).Tipo = 2
    Hechizos(14).WAV = 16
    Hechizos(15).Nombre = "Tormenta de fuego"
    Hechizos(15).Desc = "Causa 35 a 55 puntos de da�o a la victima."
    Hechizos(15).PalabrasMagicas = "EN VAX ON TAR"
    Hechizos(15).HechizeroMsg = "Has lanzado Tormenta de fuego sobre"
    Hechizos(15).PropioMsg = "Has lanzado Tormenta de fuego sobre t�."
    Hechizos(15).TargetMsg = "lanzo Tormenta de fuego sobre t�."
    Hechizos(15).FXgrh = 7
    Hechizos(15).Loops = 1
    Hechizos(15).ManaRequerido = 250
    Hechizos(15).MaxHP = 55
    Hechizos(15).MinHP = 35
    Hechizos(15).MinSkill = 75
    Hechizos(15).Resis = 1
    Hechizos(15).SubeHP = 2
    Hechizos(15).Target = 3
    Hechizos(15).Tipo = 1
    Hechizos(15).WAV = 27
    Hechizos(16).Nombre = "Llamado a la naturaleza"
    Hechizos(16).Desc = "Implora ayuda a la madre naturaleza, tres lobos acudiran en tu ayuda."
    Hechizos(16).PalabrasMagicas = "Nature et worg"
    Hechizos(16).HechizeroMsg = "Has lanzado Llamado a la naturaleza"
    Hechizos(16).PropioMsg = ""
    Hechizos(16).TargetMsg = ""
    Hechizos(16).Cant = 3
    Hechizos(16).invoca = 1
    Hechizos(16).ManaRequerido = 120
    Hechizos(16).MinSkill = 40
    Hechizos(16).NumNpc = 545
    Hechizos(16).Target = 4
    Hechizos(16).Tipo = 4
    Hechizos(16).WAV = 17
    Hechizos(17).Nombre = "Invokar Zombies"
    Hechizos(17).Desc = "Invoca la ayuda del los muertos, tres zombies acudiran en tu ayuda."
    Hechizos(17).PalabrasMagicas = "Mo� c�mus"
    Hechizos(17).HechizeroMsg = "Has invocado tres Zombies"
    Hechizos(17).PropioMsg = ""
    Hechizos(17).TargetMsg = ""
    Hechizos(17).Cant = 3
    Hechizos(17).invoca = 1
    Hechizos(17).ManaRequerido = 220
    Hechizos(17).MinSkill = 70
    Hechizos(17).NumNpc = 546
    Hechizos(17).Target = 4
    Hechizos(17).Tipo = 4
    Hechizos(17).WAV = 17
    Hechizos(18).Nombre = "Celeridad"
    Hechizos(18).Desc = "Aumenta la agilidad del usuario que recibe el spell"
    Hechizos(18).PalabrasMagicas = "YUP A'INC"
    Hechizos(18).HechizeroMsg = "Has lanzado celeridad sobre "
    Hechizos(18).PropioMsg = "Has lanzado celeridad sobre t�."
    Hechizos(18).TargetMsg = "ha lanzado celeridad sobre t�."
    Hechizos(18).FXgrh = 37
    Hechizos(18).Loops = 1
    Hechizos(18).ManaRequerido = 50
    Hechizos(18).MaxAgilidad = 5
    Hechizos(18).MinAgilidad = 2
    Hechizos(18).MinSkill = 35
    Hechizos(18).SubeAgilidad = 1
    Hechizos(18).Target = 1
    Hechizos(18).Tipo = 1
    Hechizos(18).WAV = 17
    Hechizos(19).Nombre = "Torpeza"
    Hechizos(19).Desc = "Reduce la agilidad del usuario que recibe el spell"
    Hechizos(19).PalabrasMagicas = "ASYNC YUP A'INC"
    Hechizos(19).HechizeroMsg = "Has lanzado torpeza sobre "
    Hechizos(19).PropioMsg = "Has lanzado torpeza sobre t�."
    Hechizos(19).TargetMsg = "ha lanzado torpeza sobre t�."
    Hechizos(19).FXgrh = 28
    Hechizos(19).Loops = 1
    Hechizos(19).ManaRequerido = 50
    Hechizos(19).MaxAgilidad = 5
    Hechizos(19).MinAgilidad = 2
    Hechizos(19).MinSkill = 20
    Hechizos(19).SubeAgilidad = 2
    Hechizos(19).Target = 1
    Hechizos(19).Tipo = 1
    Hechizos(19).WAV = 17
    Hechizos(20).Nombre = "Fuerza"
    Hechizos(20).Desc = "Aumenta la fuerza del objetivo"
    Hechizos(20).PalabrasMagicas = "Ar A'kron"
    Hechizos(20).HechizeroMsg = "Has lanzado fuerza sobre "
    Hechizos(20).PropioMsg = "Has lanzado fuerza sobre t�."
    Hechizos(20).TargetMsg = "ha lanzado fuerza sobre t�."
    Hechizos(20).FXgrh = 83
    Hechizos(20).Loops = 3
    Hechizos(20).ManaRequerido = 50
    Hechizos(20).MaxFuerza = 5
    Hechizos(20).MinFuerza = 2
    Hechizos(20).MinSkill = 35
    Hechizos(20).SubeFuerza = 1
    Hechizos(20).Target = 1
    Hechizos(20).Tipo = 1
    Hechizos(20).WAV = 17
    Hechizos(21).Nombre = "Debilidad"
    Hechizos(21).Desc = "Reduce la fuerza del usuario que recibe el spell"
    Hechizos(21).PalabrasMagicas = "Xoom varp"
    Hechizos(21).HechizeroMsg = "Has lanzado Debilidad sobre "
    Hechizos(21).PropioMsg = "Has lanzado Debilidad sobre t�."
    Hechizos(21).TargetMsg = "ha lanzado Debilidad sobre t�."
    Hechizos(21).FXgrh = 28
    Hechizos(21).Loops = 1
    Hechizos(21).ManaRequerido = 45
    Hechizos(21).MaxFuerza = 5
    Hechizos(21).MinFuerza = 2
    Hechizos(21).MinSkill = 35
    Hechizos(21).SubeFuerza = 2
    Hechizos(21).Target = 1
    Hechizos(21).Tipo = 1
    Hechizos(21).WAV = 17
    Hechizos(22).Nombre = "Fuerza II"
    Hechizos(22).Desc = "Aumenta la fuerza del objetivo"
    Hechizos(22).PalabrasMagicas = "Ar A'kron II"
    Hechizos(22).HechizeroMsg = "Has lanzado fuerza II sobre "
    Hechizos(22).PropioMsg = "Has lanzado fuerza II sobre t�."
    Hechizos(22).TargetMsg = "ha lanzado fuerza II sobre t�."
    Hechizos(22).FXgrh = 83
    Hechizos(22).Loops = 3
    Hechizos(22).ManaRequerido = 75
    Hechizos(22).MaxFuerza = 15
    Hechizos(22).MinFuerza = 10
    Hechizos(22).MinSkill = 60
    Hechizos(22).SubeFuerza = 1
    Hechizos(22).Target = 1
    Hechizos(22).Tipo = 1
    Hechizos(22).WAV = 17
    Hechizos(23).Nombre = "Descarga electrica"
    Hechizos(23).Desc = "Causa 45 a 65 puntos de da�o a la victima."
    Hechizos(23).PalabrasMagicas = "T'HY KOOOL"
    Hechizos(23).HechizeroMsg = "Has lanzado Descarga electrica sobre"
    Hechizos(23).PropioMsg = "Has lanzado Descarga electrica sobre t�."
    Hechizos(23).TargetMsg = "lanzo Descarga electrica t�."
    Hechizos(23).FXgrh = 121
    Hechizos(23).Loops = 1
    Hechizos(23).ManaRequerido = 350
    Hechizos(23).MaxHP = 65
    Hechizos(23).MinHP = 45
    Hechizos(23).MinSkill = 75
    Hechizos(23).Resis = 1
    Hechizos(23).SubeHP = 2
    Hechizos(23).Target = 3
    Hechizos(23).Tipo = 1
    Hechizos(23).WAV = 108
    Hechizos(24).Nombre = "Par�lisis De Odin"
    Hechizos(24).Desc = "Paraliza por un momento a la victima."
    Hechizos(24).PalabrasMagicas = "HOAX VORP AR ZONE"
    Hechizos(24).HechizeroMsg = "Has paralizado a"
    Hechizos(24).PropioMsg = "Te has paralizado."
    Hechizos(24).TargetMsg = "te ha paralizado."
    Hechizos(24).FXgrh = 8
    Hechizos(24).Loops = 1
    Hechizos(24).ManaRequerido = 2000
    Hechizos(24).MinSkill = 170
    Hechizos(24).MinNivel = 40
    Hechizos(24).Paralizaarea = 1
    Hechizos(24).Target = 2
    Hechizos(24).Tipo = 2
    Hechizos(24).WAV = 16
    Hechizos(25).Nombre = "Apocalipsis"
    Hechizos(25).Desc = "Causa 75 a 80 puntos de da�o a la victima."
    Hechizos(25).PalabrasMagicas = "Rahma Na�arak O'al"
    Hechizos(25).HechizeroMsg = "Has lanzado Apocalipsis sobre"
    Hechizos(25).PropioMsg = "Has lanzado Apocalipsis sobre t�."
    Hechizos(25).TargetMsg = "lanzo Apocalipsis t�."
    Hechizos(25).FXgrh = 13
    Hechizos(25).Loops = 1
    Hechizos(25).ManaRequerido = 640
    Hechizos(25).MaxHP = 85
    Hechizos(25).MinHP = 80
    Hechizos(25).MinSkill = 110
    Hechizos(25).Resis = 1
    Hechizos(25).SubeHP = 2
    Hechizos(25).Target = 3
    Hechizos(25).Tipo = 1
    Hechizos(25).WAV = 27
    Hechizos(26).Nombre = "Invocar elemental de fuego"
    Hechizos(26).Desc = "Invocar elemental de fuego"
    Hechizos(26).PalabrasMagicas = "Yur'rax"
    Hechizos(26).HechizeroMsg = "Has invocado un elemental de fuego"
    Hechizos(26).PropioMsg = ""
    Hechizos(26).TargetMsg = ""
    Hechizos(26).Cant = 1
    Hechizos(26).invoca = 1
    Hechizos(26).ManaRequerido = 620
    Hechizos(26).MinSkill = 100
    Hechizos(26).NumNpc = 93
    Hechizos(26).Target = 4
    Hechizos(26).Tipo = 4
    Hechizos(26).WAV = 17
    Hechizos(27).Nombre = "Invocar elemental de agua"
    Hechizos(27).Desc = "Invocar elemental de agua"
    Hechizos(27).PalabrasMagicas = "Mantra'rax"
    Hechizos(27).HechizeroMsg = "Has invocado un elemental de agua"
    Hechizos(27).PropioMsg = ""
    Hechizos(27).TargetMsg = ""
    Hechizos(27).Cant = 1
    Hechizos(27).invoca = 1
    Hechizos(27).ManaRequerido = 620
    Hechizos(27).MinSkill = 100
    Hechizos(27).NumNpc = 92
    Hechizos(27).Target = 4
    Hechizos(27).Tipo = 4
    Hechizos(27).WAV = 17
    Hechizos(28).Nombre = "Invocar elemental de tierra"
    Hechizos(28).Desc = "Invocar elemental de tierra"
    Hechizos(28).PalabrasMagicas = "Roc'rax"
    Hechizos(28).HechizeroMsg = "Has invocado un elemental de tierra"
    Hechizos(28).PropioMsg = ""
    Hechizos(28).TargetMsg = ""
    Hechizos(28).Cant = 1
    Hechizos(28).invoca = 1
    Hechizos(28).ManaRequerido = 620
    Hechizos(28).MinSkill = 100
    Hechizos(28).NumNpc = 94
    Hechizos(28).Target = 4
    Hechizos(28).Tipo = 4
    Hechizos(28).WAV = 17
    Hechizos(29).Nombre = "Invocar Cerbero"
    Hechizos(29).Desc = "Invocar Cerbero"
    Hechizos(29).PalabrasMagicas = "InF Cerb'RoX"
    Hechizos(29).HechizeroMsg = "Has implorado ayuda a los dioses!"
    Hechizos(29).PropioMsg = ""
    Hechizos(29).TargetMsg = ""
    Hechizos(29).Cant = 1
    Hechizos(29).invoca = 1
    Hechizos(29).ManaRequerido = 1820
    Hechizos(29).MinSkill = 150
    Hechizos(29).NumNpc = 259
    Hechizos(29).Target = 4
    Hechizos(29).Tipo = 4
    Hechizos(29).WAV = 17
    Hechizos(30).Nombre = "Ceguera"
    Hechizos(30).Desc = "Ceguera"
    Hechizos(30).PalabrasMagicas = ""
    Hechizos(30).HechizeroMsg = "Has lanzado ceguera sobre"
    Hechizos(30).PropioMsg = "Has lanzado ceguera sobre t�."
    Hechizos(30).TargetMsg = "lanzo ceguera t�."
    Hechizos(30).Ceguera = 1
    Hechizos(30).FXgrh = 95
    Hechizos(30).ManaRequerido = 620
    Hechizos(30).MinSkill = 70
    Hechizos(30).Target = 1
    Hechizos(30).Tipo = 2
    Hechizos(30).WAV = 17
    Hechizos(31).Nombre = "Estupidez"
    Hechizos(31).Desc = "Estupidez"
    Hechizos(31).PalabrasMagicas = ""
    Hechizos(31).HechizeroMsg = "Has lanzado Estupidez sobre"
    Hechizos(31).PropioMsg = "Has lanzado Estupidez sobre t�."
    Hechizos(31).TargetMsg = "lanzo Estupidez sobre t�."
    Hechizos(31).Estupidez = 1
    Hechizos(31).FXgrh = 95
    Hechizos(31).Loops = 3
    Hechizos(31).ManaRequerido = 620
    Hechizos(31).MinSkill = 70
    Hechizos(31).Target = 1
    Hechizos(31).Tipo = 2
    Hechizos(31).WAV = 17
    Hechizos(32).Nombre = "Curaci�n Milagrosa"
    Hechizos(32).Desc = "Curar heridas muy graves, restaura entre 150 y 200 puntos de salud."
    Hechizos(32).PalabrasMagicas = "SANCTIS DEUS"
    Hechizos(32).HechizeroMsg = "Has sanado a"
    Hechizos(32).PropioMsg = "Te has curado muchas heridas."
    Hechizos(32).TargetMsg = "te ha curado muchas heridas."
    Hechizos(32).FXgrh = 31
    Hechizos(32).Loops = 1
    Hechizos(32).ManaRequerido = 500
    Hechizos(32).MaxHP = 200
    Hechizos(32).MinHP = 150
    Hechizos(32).MinSkill = 130
    Hechizos(32).SubeHP = 1
    Hechizos(32).Target = 3
    Hechizos(32).Tipo = 1
    Hechizos(32).WAV = 18
    Hechizos(33).Nombre = "MINIApocalipsis"
    Hechizos(33).Desc = "Causa 65 a 70 puntos de da�o a la victima."
    Hechizos(33).PalabrasMagicas = "R�aM ShArek"
    Hechizos(33).HechizeroMsg = "Has lanzado MINIApocalipsis sobre"
    Hechizos(33).PropioMsg = "Has lanzado MINIApocalipsis sobre t�."
    Hechizos(33).TargetMsg = "lanzo MINIApocalipsis t�."
    Hechizos(33).FXgrh = 93
    Hechizos(33).Loops = 1
    Hechizos(33).ManaRequerido = 399
    Hechizos(33).MaxHP = 70
    Hechizos(33).MinHP = 65
    Hechizos(33).MinSkill = 60
    Hechizos(33).Resis = 1
    Hechizos(33).SubeHP = 2
    Hechizos(33).Target = 3
    Hechizos(33).Tipo = 1
    Hechizos(33).WAV = 27
    Hechizos(34).Nombre = "Rayo GM"
    Hechizos(34).Desc = "Causa 9000 puntos de da�o a la victima."
    Hechizos(34).PalabrasMagicas = "R�C U MAS"
    Hechizos(34).HechizeroMsg = "Has lanzado Rayo GM sobre"
    Hechizos(34).PropioMsg = ""
    Hechizos(34).TargetMsg = "lanzo muerte sobre t�."
    Hechizos(34).FXgrh = 15
    Hechizos(34).Loops = 3
    Hechizos(34).MaxHP = 9000
    Hechizos(34).MinHP = 9000
    Hechizos(34).Resis = 1
    Hechizos(34).SubeHP = 2
    Hechizos(34).Target = 3
    Hechizos(34).Tipo = 1
    Hechizos(34).WAV = 16
    Hechizos(35).Nombre = "Bomba de Humo"
    Hechizos(35).Desc = "ceguera , estupidez "
    Hechizos(35).PalabrasMagicas = "M�C Smoke"
    Hechizos(35).HechizeroMsg = "Has lanzado Bomba de Humo sobre"
    Hechizos(35).PropioMsg = "Has lanzado Bomba de Humo sobre t�."
    Hechizos(35).TargetMsg = "lanzo Bomba de Humo sobre t�."
    Hechizos(35).Ceguera = 1
    Hechizos(35).Estupidez = 1
    Hechizos(35).FXgrh = 74
    Hechizos(35).Loops = 2
    Hechizos(35).ManaRequerido = 700
    Hechizos(35).MinSkill = 70
    Hechizos(35).Target = 1
    Hechizos(35).Tipo = 2
    Hechizos(35).WAV = 17
    Hechizos(36).Nombre = "Invocar Super Elemental De Cristal"
    Hechizos(36).Desc = "Invocar Super Elemental De Cristal"
    Hechizos(36).PalabrasMagicas = "Ra�C Th�"
    Hechizos(36).HechizeroMsg = "Has invocado un Super Elemental De Cristal"
    Hechizos(36).PropioMsg = ""
    Hechizos(36).TargetMsg = ""
    Hechizos(36).Cant = 1
    Hechizos(36).invoca = 1
    Hechizos(36).ManaRequerido = 1800
    Hechizos(36).MinSkill = 200
    Hechizos(36).MinNivel = 40
    Hechizos(36).NumNpc = 107
    Hechizos(36).Target = 4
    Hechizos(36).Tipo = 4
    Hechizos(36).WAV = 17
    Hechizos(37).Nombre = "Poder Divino"
    Hechizos(37).Desc = "Resucitar un usuario muerto."
    Hechizos(37).PalabrasMagicas = "AHIL KN� X�R"
    Hechizos(37).HechizeroMsg = "Has resucitado a"
    Hechizos(37).PropioMsg = "Te has resucitado."
    Hechizos(37).TargetMsg = "te ha resucitado."
    Hechizos(37).FXgrh = 72
    Hechizos(37).Revivir = 1
    Hechizos(37).Target = 1
    Hechizos(37).Tipo = 2
    Hechizos(37).WAV = 20
    Hechizos(38).Nombre = "Ira De Dios"
    Hechizos(38).Desc = "Causa 150 a 180 puntos de da�o a todos en el �rea."
    Hechizos(38).PalabrasMagicas = "DEUS EX MACHINA"
    Hechizos(38).HechizeroMsg = "Has lanzado Ir� De Dios sobre"
    Hechizos(38).PropioMsg = "Has lanzado Ir� De Dios sobre t�."
    Hechizos(38).TargetMsg = "lanzo Ir� De Dios t�."
    Hechizos(38).FXgrh = 20
    Hechizos(38).Loops = 1
    Hechizos(38).MaxHP = 180
    Hechizos(38).MinHP = 150
    Hechizos(38).Resis = 1
    Hechizos(38).SubeHP = 4
    Hechizos(38).Target = 3
    Hechizos(38).Tipo = 1
    Hechizos(38).WAV = 27
    Hechizos(39).Nombre = "Curacion Divina"
    Hechizos(39).Desc = "Sana completamente, restaura la sed y el hambre."
    Hechizos(39).PalabrasMagicas = "SANCTIS CORPUS"
    Hechizos(39).HechizeroMsg = "Has sanado a"
    Hechizos(39).PropioMsg = "Te has curado algunas heridas."
    Hechizos(39).TargetMsg = "te ha curado algunas heridas."
    Hechizos(39).FXgrh = 106
    Hechizos(39).Loops = 1
    Hechizos(39).ManaRequerido = 1000
    Hechizos(39).MaxHam = 500
    Hechizos(39).MaxHP = 550
    Hechizos(39).MaxSed = 1500
    Hechizos(39).MinHam = 500
    Hechizos(39).MinHP = 550
    Hechizos(39).MinSed = 1500
    Hechizos(39).MinSkill = 130
    Hechizos(39).SubeHam = 1
    Hechizos(39).SubeHP = 1
    Hechizos(39).SubeSed = 1
    Hechizos(39).Target = 1
    Hechizos(39).Tipo = 1
    Hechizos(39).WAV = 18
    Hechizos(40).Nombre = "Celeridad II"
    Hechizos(40).Desc = "Aumenta la agilidad del usuario que recibe el spell"
    Hechizos(40).PalabrasMagicas = "YUP A'INC II"
    Hechizos(40).HechizeroMsg = "Has lanzado celeridad II sobre "
    Hechizos(40).PropioMsg = "Has lanzado celeridad II sobre t�."
    Hechizos(40).TargetMsg = "ha lanzado celeridad II sobre t�."
    Hechizos(40).FXgrh = 37
    Hechizos(40).ManaRequerido = 75
    Hechizos(40).MaxAgilidad = 15
    Hechizos(40).MinAgilidad = 10
    Hechizos(40).MinSkill = 60
    Hechizos(40).SubeAgilidad = 1
    Hechizos(40).Target = 1
    Hechizos(40).Tipo = 1
    Hechizos(40).WAV = 17
    Hechizos(41).Nombre = "Implosi�n"
    Hechizos(41).Desc = "Causa 120 a 130 puntos de da�o a las victimas cercanas al lanzador."
    Hechizos(41).PalabrasMagicas = "IMPLO'BUMX "
    Hechizos(41).HechizeroMsg = "Has lanzado Implosi�n sobre"
    Hechizos(41).PropioMsg = "Has lanzado Implosi�n sobre t�."
    Hechizos(41).TargetMsg = "lanzo Implosi�n t�."
    Hechizos(41).FXgrh = 23
    Hechizos(41).Loops = 1
    Hechizos(41).ManaRequerido = 2000
    Hechizos(41).MaxHP = 130
    Hechizos(41).MinHP = 120
    Hechizos(41).MinSkill = 200
    Hechizos(41).MinNivel = 40
    Hechizos(41).Resis = 1
    Hechizos(41).SubeHP = 3
    Hechizos(41).Target = 3
    Hechizos(41).Tipo = 1
    Hechizos(41).WAV = 27
    Hechizos(42).Nombre = "Curar Heridas Cr�ticas"
    Hechizos(42).Desc = "Curar heridas cr�ticas, restaura entre 45 y 70 puntos de salud."
    Hechizos(42).PalabrasMagicas = "EN CORP'CRI SANCTIS"
    Hechizos(42).HechizeroMsg = "Has sanado a"
    Hechizos(42).PropioMsg = "Te has curado bastantes heridas."
    Hechizos(42).TargetMsg = "te ha curado bastantes heridas."
    Hechizos(42).FXgrh = 31
    Hechizos(42).Loops = 1
    Hechizos(42).ManaRequerido = 150
    Hechizos(42).MaxHP = 70
    Hechizos(42).MinHP = 45
    Hechizos(42).MinSkill = 80
    Hechizos(42).SubeHP = 1
    Hechizos(42).Target = 3
    Hechizos(42).Tipo = 1
    Hechizos(42).WAV = 18
    Hechizos(43).Nombre = "Polymorph Animal"
    Hechizos(43).Desc = "Te transformas en otras criaturas."
    Hechizos(43).PalabrasMagicas = "MORPH^TRANSF ANIMAL"
    Hechizos(43).HechizeroMsg = "Has transformado a"
    Hechizos(43).PropioMsg = "Te has transformado."
    Hechizos(43).TargetMsg = "te ha tranformado."
    Hechizos(43).FXgrh = 25
    Hechizos(43).Loops = 2
    Hechizos(43).ManaRequerido = 2000
    Hechizos(43).MinSkill = 100
    Hechizos(43).Morph = 1
    Hechizos(43).MinNivel = 40
    Hechizos(43).Target = 1
    Hechizos(43).Tipo = 2
    Hechizos(43).WAV = 109
    Hechizos(44).Nombre = "LLuvia de Sangre"
    Hechizos(44).Desc = "Causa 85 a 95 puntos de da�o a la victima."
    Hechizos(44).PalabrasMagicas = "XUV FO SAN'G"
    Hechizos(44).HechizeroMsg = "Has lanzado Lluvia de Sangre sobre"
    Hechizos(44).PropioMsg = "Has lanzado Lluvia de Sangre sobre t�."
    Hechizos(44).TargetMsg = "lanzo LLuvia de Sangre sobre t�."
    Hechizos(44).FXgrh = 26
    Hechizos(44).Loops = 1
    Hechizos(44).ManaRequerido = 840
    Hechizos(44).MaxHP = 100
    Hechizos(44).MinHP = 85
    Hechizos(44).MinSkill = 150
    Hechizos(44).MinNivel = 30
    Hechizos(44).Resis = 1
    Hechizos(44).SubeHP = 2
    Hechizos(44).Target = 3
    Hechizos(44).Tipo = 1
    Hechizos(44).WAV = 27
    Hechizos(45).Nombre = "Llama de Dragon"
    Hechizos(45).Desc = "Causa 105 a 115 puntos de da�o a la victima."
    Hechizos(45).PalabrasMagicas = " ALL�H U DRAK"
    Hechizos(45).HechizeroMsg = "Has lanzado Llama de Dragon sobre"
    Hechizos(45).PropioMsg = "Has lanzado Llama de Dragon sobre t�."
    Hechizos(45).TargetMsg = "lanzo Llama de Dragon t�."
    Hechizos(45).FXgrh = 19
    Hechizos(45).Loops = 1
    Hechizos(45).ManaRequerido = 1000
    Hechizos(45).MaxHP = 110
    Hechizos(45).MinHP = 105
    Hechizos(45).MinSkill = 200
    Hechizos(45).MinNivel = 40
    Hechizos(45).Resis = 1
    Hechizos(45).SubeHP = 2
    Hechizos(45).Target = 3
    Hechizos(45).Tipo = 1
    Hechizos(45).WAV = 27
    Hechizos(46).Nombre = "Disparo Toxico"
    Hechizos(46).Desc = "Envenenamiento, provoca la muerte si no se contraresta el veneno."
    Hechizos(46).PalabrasMagicas = "SERP DISP"
    Hechizos(46).HechizeroMsg = "Has envenenado a"
    Hechizos(46).PropioMsg = "Te has envenenado."
    Hechizos(46).TargetMsg = "te ha envenenado."
    Hechizos(46).Envenena = 15
    Hechizos(46).FXgrh = 75
    Hechizos(46).Loops = 1
    Hechizos(46).ManaRequerido = 250
    Hechizos(46).MinSkill = 0
    Hechizos(46).Resis = 1
    Hechizos(46).Target = 1
    Hechizos(46).Tipo = 2
    Hechizos(46).WAV = 47
    Hechizos(47).Nombre = "Invocar Hada"
    Hechizos(47).Desc = "Invoca un Hada con poderosa magia de ataque."
    Hechizos(47).PalabrasMagicas = "Hada Duntra'rax"
    Hechizos(47).HechizeroMsg = "Has convocado un Hada!"
    Hechizos(47).PropioMsg = ""
    Hechizos(47).TargetMsg = ""
    Hechizos(47).Cant = 1
    Hechizos(47).invoca = 1
    Hechizos(47).ManaRequerido = 1820
    Hechizos(47).MinSkill = 180
    Hechizos(47).MinNivel = 30
    Hechizos(47).NumNpc = 129
    Hechizos(47).Target = 4
    Hechizos(47).Tipo = 4
    Hechizos(47).WAV = 47
    Hechizos(48).Nombre = "Invocar Genio"
    Hechizos(48).Desc = "Invoca un Genio con poderosa magia de ataque."
    Hechizos(48).PalabrasMagicas = "Genius Duntra'rax"
    Hechizos(48).HechizeroMsg = "Has convocado un Genio!"
    Hechizos(48).PropioMsg = ""
    Hechizos(48).TargetMsg = ""
    Hechizos(48).Cant = 1
    Hechizos(48).invoca = 1
    Hechizos(48).ManaRequerido = 2300
    Hechizos(48).MinSkill = 200
    Hechizos(48).MinNivel = 45
    Hechizos(48).NumNpc = 130
    Hechizos(48).Target = 4
    Hechizos(48).Tipo = 4
    Hechizos(48).WAV = 47
    Hechizos(49).Nombre = "Enredar"
    Hechizos(49).Desc = "Enreda por un momento a la victima."
    Hechizos(49).PalabrasMagicas = "HOAX PLANT"
    Hechizos(49).HechizeroMsg = "Has enredado a"
    Hechizos(49).PropioMsg = "Te has enredado."
    Hechizos(49).TargetMsg = "te ha enredado."
    Hechizos(49).FXgrh = 103
    Hechizos(49).Loops = 2
    Hechizos(49).MinSkill = 60
    Hechizos(49).Paraliza = 1
    Hechizos(49).Target = 3
    Hechizos(49).Tipo = 2
    Hechizos(49).WAV = 26
    Hechizos(50).Nombre = "Onda de Luz"
    Hechizos(50).Desc = "Causa 70 a 78 puntos de da�o a la victima."
    Hechizos(50).PalabrasMagicas = "ONX DE VI'T"
    Hechizos(50).HechizeroMsg = "Has lanzado Onda Luz sobre"
    Hechizos(50).PropioMsg = "Has Onda Luz sobre t�."
    Hechizos(50).TargetMsg = "lanzo Onda Luz sobre t�."
    Hechizos(50).FXgrh = 27
    Hechizos(50).Loops = 1
    Hechizos(50).ManaRequerido = 110
    Hechizos(50).MaxHP = 78
    Hechizos(50).MinHP = 70
    Hechizos(50).MinSkill = 125
    Hechizos(50).MinNivel = 30
    Hechizos(50).Resis = 1
    Hechizos(50).SubeHP = 2
    Hechizos(50).Target = 2
    Hechizos(50).Tipo = 1
    Hechizos(50).WAV = 16
    Hechizos(51).Nombre = "Llamarada"
    Hechizos(51).Desc = "Causa 75 a 80 puntos de da�o a la victima."
    Hechizos(51).PalabrasMagicas = "AMX FOX FI'R"
    Hechizos(51).HechizeroMsg = "Has lanzado Llamarada sobre"
    Hechizos(51).PropioMsg = "Has lanzado Llamarada sobre t�."
    Hechizos(51).TargetMsg = "lanzo Llamarada sobre t�."
    Hechizos(51).FXgrh = 22
    Hechizos(51).Loops = 1
    Hechizos(51).ManaRequerido = 400
    Hechizos(51).MaxHP = 80
    Hechizos(51).MinHP = 75
    Hechizos(51).MinSkill = 110
    Hechizos(51).Resis = 1
    Hechizos(51).SubeHP = 2
    Hechizos(51).Target = 3
    Hechizos(51).Tipo = 1
    Hechizos(51).WAV = 107
    Hechizos(52).Nombre = "Reavivar Sangre"
    Hechizos(52).Desc = "Aumenta la fuerza y Agilidad"
    Hechizos(52).PalabrasMagicas = "A'kron + A'INC"
    Hechizos(52).HechizeroMsg = "Has lanzado Reavivar Sangre sobre "
    Hechizos(52).PropioMsg = "Has lanzado Reavivar Sangre sobre t�."
    Hechizos(52).TargetMsg = "ha lanzado Reavivar Sangre sobre t�."
    Hechizos(52).FXgrh = 29
    Hechizos(52).MaxAgilidad = 15
    Hechizos(52).MaxFuerza = 15
    Hechizos(52).MinAgilidad = 10
    Hechizos(52).MinFuerza = 10
    Hechizos(52).SubeAgilidad = 1
    Hechizos(52).SubeFuerza = 1
    Hechizos(52).Target = 1
    Hechizos(52).Tipo = 1
    Hechizos(52).WAV = 17
    Hechizos(53).Nombre = "Poder De Lucifer"
    Hechizos(53).Desc = "Causa 150 a 180 puntos de da�o a todos en el �rea."
    Hechizos(53).PalabrasMagicas = " URUK UNGOL"
    Hechizos(53).HechizeroMsg = "Has lanzado Poder De Lucifer sobre"
    Hechizos(53).PropioMsg = "Has lanzado Poder De Lucifer sobre t�."
    Hechizos(53).TargetMsg = "lanzo Poder De Lucifer t�."
    Hechizos(53).FXgrh = 22
    Hechizos(53).Loops = 4
    Hechizos(53).MaxHP = 180
    Hechizos(53).MinHP = 150
    Hechizos(53).Resis = 1
    Hechizos(53).SubeHP = 4
    Hechizos(53).Target = 3
    Hechizos(53).Tipo = 1
    Hechizos(53).WAV = 107
    Hechizos(54).Nombre = ""
    Hechizos(54).Desc = ""
    Hechizos(54).PalabrasMagicas = ""
    Hechizos(54).HechizeroMsg = ""
    Hechizos(54).PropioMsg = ""
    Hechizos(54).TargetMsg = ""
    Hechizos(55).Nombre = "Detectar invisibilidad"
    Hechizos(55).Desc = "Remueve parcialmente los efectos de la invisibilidad"
    Hechizos(55).PalabrasMagicas = "An Ma�o naq v�k�"
    Hechizos(55).HechizeroMsg = "Has lanzado detectar invisibilidad"
    Hechizos(55).PropioMsg = "Te ha lanzado detectar invisibilidad"
    Hechizos(55).TargetMsg = "Detectas la invisibilidad alrededor tuyo."
    Hechizos(55).FXgrh = 23
    Hechizos(55).Loops = 7
    Hechizos(55).MinSkill = 100
    Hechizos(55).Target = 3
    Hechizos(55).Tipo = 2
    Hechizos(55).WAV = 16
    Hechizos(55).RemueveInvisibilidadParcial = 1
    Hechizos(55).ManaRequerido = 400
    Hechizos(55).Target = 4
    Hechizos(56).Nombre = "Repone Mana II"
    Hechizos(56).Desc = "Repone Mana"
    Hechizos(56).PalabrasMagicas = "CORP MANA"
    Hechizos(56).HechizeroMsg = "Has resturado mana a"
    Hechizos(56).PropioMsg = "Te has restaurado mana."
    Hechizos(56).TargetMsg = "te ha restaurado mana."
    Hechizos(56).MaMana = 60
    Hechizos(56).MiMana = 25
    Hechizos(56).SubeMana = 1
    Hechizos(56).Target = 3
    Hechizos(56).Tipo = 1
    Hechizos(56).WAV = 17
    Hechizos(57).Nombre = "Aura Protectora  "
    Hechizos(57).Desc = "Barrera M�gica que protege de golpes."
    Hechizos(57).PalabrasMagicas = "AURO PROT"
    Hechizos(57).HechizeroMsg = "Has protegido a"
    Hechizos(57).PropioMsg = "Te has protegido."
    Hechizos(57).TargetMsg = "te ha protegido."
    Hechizos(57).FXgrh = 91
    Hechizos(57).Loops = 1
    Hechizos(57).WAV = 16
    Hechizos(58).Nombre = "Aura Protectora II "
    Hechizos(58).Desc = "Barrera M�gica que protege de golpes."
    Hechizos(58).PalabrasMagicas = "AURO PROT MEX"
    Hechizos(58).HechizeroMsg = "Has protegido a"
    Hechizos(58).PropioMsg = "Te has protegido."
    Hechizos(58).TargetMsg = "te ha protegido."
    Hechizos(58).FXgrh = 109
    Hechizos(58).Loops = 1
    Hechizos(58).WAV = 16
    Hechizos(59).Nombre = "Arco Iris "
    Hechizos(59).Desc = ""
    Hechizos(59).PalabrasMagicas = ""
    Hechizos(59).HechizeroMsg = ""
    Hechizos(59).PropioMsg = ""
    Hechizos(59).TargetMsg = ""
    Hechizos(59).FXgrh = 92
    Hechizos(59).Loops = 2
    Hechizos(59).SubeHP = 2
    Hechizos(59).WAV = 16
    Hechizos(60).Nombre = "Restauraci�n"
    Hechizos(60).Desc = "Restaura entre 125 y 175 puntos de salud."
    Hechizos(60).PalabrasMagicas = "REST CORP SANC"
    Hechizos(60).HechizeroMsg = "Has sanado a"
    Hechizos(60).PropioMsg = "Te has curado algunas heridas."
    Hechizos(60).TargetMsg = "te ha curado algunas heridas."
    Hechizos(60).FXgrh = 116
    Hechizos(60).Loops = 1
    Hechizos(60).ManaRequerido = 400
    Hechizos(60).MaxHP = 175
    Hechizos(60).MinHP = 125
    Hechizos(60).MinSkill = 90
    Hechizos(60).SubeHP = 1
    Hechizos(60).Target = 3
    Hechizos(60).Tipo = 1
    Hechizos(60).WAV = 17
    Hechizos(61).Nombre = "Incinerar"
    Hechizos(61).Desc = "Causa 100 a 105 puntos de da�o a la victima."
    Hechizos(61).PalabrasMagicas = "FIRE 'AHC HUMUS"
    Hechizos(61).HechizeroMsg = "Has lanzado Incinerar sobre"
    Hechizos(61).PropioMsg = "Has lanzado Incinerar sobre t�."
    Hechizos(61).TargetMsg = "lanzo Incinerar sobre t�."
    Hechizos(61).FXgrh = 113
    Hechizos(61).Loops = 1
    Hechizos(61).ManaRequerido = 840
    Hechizos(61).MaxHP = 105
    Hechizos(61).MinHP = 100
    Hechizos(61).MinSkill = 150
    Hechizos(61).MinNivel = 30
    Hechizos(61).Resis = 1
    Hechizos(61).SubeHP = 2
    Hechizos(61).Target = 3
    Hechizos(61).Tipo = 1
    Hechizos(61).WAV = 27
    Hechizos(62).Nombre = "Ventisca"
    Hechizos(62).Desc = "Causa 80 a 100 puntos de da�o a la victima."
    Hechizos(62).PalabrasMagicas = "W�ND 'AHC RAI"
    Hechizos(62).HechizeroMsg = "Has lanzado Ventisca sobre"
    Hechizos(62).PropioMsg = "Has lanzado Ventisca sobre t�."
    Hechizos(62).TargetMsg = "lanzo Ventisca sobre t�."
    Hechizos(62).FXgrh = 114
    Hechizos(62).Loops = 1
    Hechizos(62).ManaRequerido = 840
    Hechizos(62).MaxHP = 100
    Hechizos(62).MinHP = 80
    Hechizos(62).MinSkill = 150
    Hechizos(62).MinNivel = 30
    Hechizos(62).Resis = 1
    Hechizos(62).SubeHP = 2
    Hechizos(62).Target = 3
    Hechizos(62).Tipo = 1
    Hechizos(62).WAV = 27
    Hechizos(63).Nombre = "Rayo de Fuego"
    Hechizos(63).Desc = "Causa 35 a 55 puntos de da�o a la victima."
    Hechizos(63).PalabrasMagicas = "RAX FOX FI'R"
    Hechizos(63).HechizeroMsg = "Has lanzado Rayo de Fuego sobre"
    Hechizos(63).PropioMsg = "Has lanzado Rayo de Fuego sobre t�."
    Hechizos(63).TargetMsg = "lanzo Rayo de Fuego sobre t�."
    Hechizos(63).FXgrh = 115
    Hechizos(63).Loops = 2
    Hechizos(63).ManaRequerido = 250
    Hechizos(63).MaxHP = 55
    Hechizos(63).MinHP = 35
    Hechizos(63).MinSkill = 75
    Hechizos(63).Resis = 1
    Hechizos(63).SubeHP = 2
    Hechizos(63).Target = 3
    Hechizos(63).Tipo = 1
    Hechizos(63).WAV = 107
    Hechizos(64).Nombre = "Rayo El�ctrico"
    Hechizos(64).Desc = "Causa 85 a 125 puntos de da�o a la victima."
    Hechizos(64).PalabrasMagicas = "RAX ELEC'X"
    Hechizos(64).HechizeroMsg = "Has lanzado Rayo El�ctrico sobre"
    Hechizos(64).PropioMsg = "Has lanzado Rayo El�ctrico sobre t�."
    Hechizos(64).TargetMsg = "lanzo Rayo El�ctrico sobre t�."
    Hechizos(64).FXgrh = 110
    Hechizos(64).Loops = 2
    Hechizos(64).ManaRequerido = 350
    Hechizos(64).MaxHP = 65
    Hechizos(64).MinHP = 45
    Hechizos(64).MinSkill = 75
    Hechizos(64).Resis = 1
    Hechizos(64).SubeHP = 2
    Hechizos(64).Target = 3
    Hechizos(64).Tipo = 1
    Hechizos(64).WAV = 108
    Hechizos(65).Nombre = "Petrificar"
    Hechizos(65).Desc = "Paraliza por un momento a la victima"
    Hechizos(65).PalabrasMagicas = "HOAX ROC"
    Hechizos(65).HechizeroMsg = "Has petrificado a"
    Hechizos(65).PropioMsg = "Te has petrificado."
    Hechizos(65).TargetMsg = "te ha inmovilizado."
    Hechizos(65).FXgrh = 108
    Hechizos(65).Loops = 1
    Hechizos(65).ManaRequerido = 450
    Hechizos(65).MinSkill = 60
    Hechizos(65).Paraliza = 1
    Hechizos(65).Target = 3
    Hechizos(65).Tipo = 2
    Hechizos(65).WAV = 16
    Hechizos(66).Nombre = "Inmovilizar "
    Hechizos(66).Desc = "Paraliza por un momento a la victima."
    Hechizos(66).PalabrasMagicas = "HOAX CEP'X"
    Hechizos(66).HechizeroMsg = "Has inmovilizado a"
    Hechizos(66).PropioMsg = "Te has inmovilizado."
    Hechizos(66).TargetMsg = "te ha inmovilizado."
    Hechizos(66).FXgrh = 104
    Hechizos(66).Loops = 3
    Hechizos(66).ManaRequerido = 450
    Hechizos(66).MinSkill = 60
    Hechizos(66).Paraliza = 1
    Hechizos(66).Target = 3
    Hechizos(66).Tipo = 2
    Hechizos(66).WAV = 26
    Hechizos(67).Nombre = "Circulo de Protecci�n"
    Hechizos(67).Desc = "Barrera M�gica que protege de golpes."
    Hechizos(67).PalabrasMagicas = "R�DO PROT MEX"
    Hechizos(67).HechizeroMsg = "Has protegido a"
    Hechizos(67).PropioMsg = "Te has protegido."
    Hechizos(67).TargetMsg = "te ha protegido."
    Hechizos(67).FXgrh = 102
    Hechizos(67).Loops = 3
    Hechizos(67).ManaRequerido = 800
    Hechizos(67).MinSkill = 100
    Hechizos(67).Protec = 10
    Hechizos(67).Target = 1
    Hechizos(67).Tipo = 2
    Hechizos(67).WAV = 16
    Hechizos(68).Nombre = "Aliento de Drag�n"
    Hechizos(68).Desc = "Causa 125 a 140 puntos de da�o a la victima."
    Hechizos(68).PalabrasMagicas = "IT� U DRAK"
    Hechizos(68).HechizeroMsg = "Has lanzado Aliento sobre"
    Hechizos(68).PropioMsg = ""
    Hechizos(68).TargetMsg = "lanzo Aliento de Drag�n sobre t�."
    Hechizos(68).FXgrh = 101
    Hechizos(68).Loops = 1
    Hechizos(68).ManaRequerido = 1200
    Hechizos(68).MaxHP = 140
    Hechizos(68).MinHP = 125
    Hechizos(68).MinSkill = 150
    Hechizos(68).MinNivel = 40
    Hechizos(68).Resis = 1
    Hechizos(68).SubeHP = 2
    Hechizos(68).Target = 3
    Hechizos(68).Tipo = 1
    Hechizos(68).WAV = 16
    Hechizos(69).Nombre = "Hechizo Mortal"
    Hechizos(69).Desc = "Causa 3500 a 5000 puntos de da�o a la victima."
    Hechizos(69).PalabrasMagicas = "Mor't DraG"
    Hechizos(69).HechizeroMsg = "Has lanzado Hechizo Mortal sobre"
    Hechizos(69).PropioMsg = "Has lanzado Hechizo Mortal sobre t�."
    Hechizos(69).TargetMsg = "lanzo Hechizo Mortal sobre t�."
    Hechizos(69).FXgrh = 33
    Hechizos(69).Loops = 1
    Hechizos(69).ManaRequerido = 2000
    Hechizos(69).MaxHP = 5000
    Hechizos(69).MinHP = 3500
    Hechizos(69).MinSkill = 200
    Hechizos(69).MinNivel = 70
    Hechizos(69).Resis = 1
    Hechizos(69).SubeHP = 2
    Hechizos(69).Target = 3
    Hechizos(69).Tipo = 1
    Hechizos(69).WAV = 27
    Hechizos(70).Nombre = "Disparo de Salva"
    Hechizos(70).Desc = "Causa 90 a 98 puntos de da�o a la victima."
    Hechizos(70).PalabrasMagicas = ""
    Hechizos(70).HechizeroMsg = "Has lanzado Disparo de Salva sobre"
    Hechizos(70).PropioMsg = "Has Disparo de Salva sobre t�."
    Hechizos(70).TargetMsg = "lanzo Disparo de Salva sobre t�."
    Hechizos(70).FXgrh = 14
    Hechizos(70).Loops = 1
    Hechizos(70).ManaRequerido = 75
    Hechizos(70).MaxHP = 98
    Hechizos(70).MinHP = 90
    Hechizos(70).MinSkill = 0
    Hechizos(70).MinNivel = 20
    Hechizos(70).Resis = 1
    Hechizos(70).SubeHP = 2
    Hechizos(70).Target = 2
    Hechizos(70).Tipo = 1
    Hechizos(70).WAV = 27
    Hechizos(71).Nombre = "Disparo de Conmoci�n"
    Hechizos(71).Desc = "Paraliza por un momento a la v�ctima."
    Hechizos(71).PalabrasMagicas = "HOAX VORP"
    Hechizos(71).HechizeroMsg = "Has Conmocionado a"
    Hechizos(71).PropioMsg = "Te has Conmocionado."
    Hechizos(71).TargetMsg = "te ha Conmocionado."
    Hechizos(71).FXgrh = 76
    Hechizos(71).Loops = 1
    Hechizos(71).ManaRequerido = 350
    Hechizos(71).MinSkill = 0
    Hechizos(71).Paraliza = 1
    Hechizos(71).Target = 3
    Hechizos(71).Tipo = 2
    Hechizos(71).WAV = 16
    Hechizos(72).Nombre = "Disparo Certero"
    Hechizos(72).Desc = "Causa 105 a 115 puntos de da�o a la victima."
    Hechizos(72).PalabrasMagicas = ""
    Hechizos(72).HechizeroMsg = "Has lanzado Disparo Certero sobre"
    Hechizos(72).PropioMsg = "Has lanzado Disparo Certero sobre t�."
    Hechizos(72).TargetMsg = "lanzo Disparo Certero t�."
    Hechizos(72).FXgrh = 105
    Hechizos(72).Loops = 1
    Hechizos(72).ManaRequerido = 680
    Hechizos(72).MaxHP = 115
    Hechizos(72).MinHP = 105
    Hechizos(72).MinSkill = 0
    Hechizos(72).MinNivel = 35
    Hechizos(72).Resis = 1
    Hechizos(72).SubeHP = 2
    Hechizos(72).Target = 3
    Hechizos(72).Tipo = 1
    Hechizos(72).WAV = 27
    Hechizos(73).Nombre = "Destello infernal"
    Hechizos(73).Desc = "Causa 75 a 80 puntos de da�o a la victima."
    Hechizos(73).PalabrasMagicas = "Oleek U'ber"
    Hechizos(73).HechizeroMsg = "Has lanzado Destello infernal sobre"
    Hechizos(73).PropioMsg = "Has lanzado Destello infernal sobre t�."
    Hechizos(73).TargetMsg = "lanzo Destello infernal t�."
    Hechizos(73).FXgrh = 128
    Hechizos(73).Loops = 1
    Hechizos(73).ManaRequerido = 640
    Hechizos(73).MaxHP = 85
    Hechizos(73).MinHP = 80
    Hechizos(73).MinSkill = 110
    Hechizos(73).Resis = 1
    Hechizos(73).SubeHP = 2
    Hechizos(73).Target = 3
    Hechizos(73).Tipo = 1
    Hechizos(73).WAV = 27
    Hechizos(74).Nombre = "Instinto Animal"
    Hechizos(74).PalabrasMagicas = "Ar A'Kazar"
    Hechizos(74).HechizeroMsg = "Has lanzado Instinto Animal."
    Hechizos(74).TargetMsg = "ha lanzado Instinto Animal sobre t�."
    Hechizos(74).PropioMsg = "Has lanzado Instinto Animal sobre t�."
    Hechizos(74).Tipo = 1
    Hechizos(74).WAV = 17
    Hechizos(74).FXgrh = 83
    Hechizos(74).Loops = 3
    Hechizos(74).SubeFuerza = 1
    Hechizos(74).MinFuerza = 10
    Hechizos(74).MaxFuerza = 15
    Hechizos(74).SubeAgilidad = 1
    Hechizos(40).MaxAgilidad = 15
    Hechizos(40).MinAgilidad = 10
    Hechizos(74).MinSkill = 0
    Hechizos(74).ManaRequerido = 0
    Hechizos(74).Target = 1
    Hechizos(74).Noesquivar = 1
    Hechizos(75).Nombre = "Purificar"
    Hechizos(75).Desc = "Causa 85 a 95 puntos de da�o a la victima y devuelve la mitad de da�o inflijido a ti."
    Hechizos(75).PalabrasMagicas = "TIV FO PUR'G"
    Hechizos(75).HechizeroMsg = "Has lanzado Purificar sobre"
    Hechizos(75).TargetMsg = "lanzo Purificar sobre t�."
    Hechizos(75).PropioMsg = "Has lanzado Purificar sobre t�."
    Hechizos(75).Tipo = 1
    Hechizos(75).WAV = 27
    Hechizos(75).FXgrh = 90
    Hechizos(75).Loops = 1
    Hechizos(75).Resis = 1
    Hechizos(75).SubeHP = 2
    Hechizos(75).MinHP = 85
    Hechizos(75).MaxHP = 100
    Hechizos(75).MinSkill = 150
    Hechizos(75).ManaRequerido = 1500
    Hechizos(75).MinNivel = 30
    Hechizos(75).Target = 3

    Exit Sub
Errorfeo:
    MsgBox ("error en cargarhechizos")

End Sub

Sub CargamosObjetos()

    Dim Leer As clsIniManager

    Set Leer = New clsIniManager

    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Object As Long, NumObjDatas As Long
    Leer.Initialize App.Path & "\Init\Obj.dat"

    'obtiene el numero de obj
    NumObjDatas = Val(Leer.GetValue("INIT", "NumObjs"))

    ReDim ObjData(1 To NumObjDatas) As ObjData

    'Llena la lista
    For Object = 1 To NumObjDatas

        ObjData(Object).Name = Leer.GetValue("OBJ" & Object, "Name")
        'ObjData(Object).Name = Leer.GetValue("OBJ" & Object, "Name")
        'pluto 2.17
        ObjData(Object).Magia = Val(Leer.GetValue("OBJ" & Object, "Magia"))

        'pluto:2.8.0
        ObjData(Object).Vendible = Val(Leer.GetValue("OBJ" & Object, "Vendible"))

        ObjData(Object).GrhIndex = Val(Leer.GetValue("OBJ" & Object, "GrhIndex"))

        ObjData(Object).ObjType = Val(Leer.GetValue("OBJ" & Object, "ObjType"))
        ObjData(Object).SubTipo = Val(Leer.GetValue("OBJ" & Object, "Subtipo"))
        'pluto:6.0A
        ObjData(Object).ArmaNpc = Val(Leer.GetValue("OBJ" & Object, "ArmaNpc"))

        ObjData(Object).Newbie = Val(Leer.GetValue("OBJ" & Object, "Newbie"))
        'pluto:2.3
        ObjData(Object).peso = 0    ' val(Leer.GetValue("OBJ" & Object, "Peso"))

        If ObjData(Object).SubTipo = OBJTYPE_ESCUDO Then
            ObjData(Object).ShieldAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))  ' * 2
            '[MerLiNz:6]
            ObjData(Object).Gemas = Val(Leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(Leer.GetValue("OBJ" & Object, "Diamantes"))

            '[\END]
        End If

        'pluto:6.2----------
        If ObjData(Object).ObjType = OBJTYPE_Anillo Then
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))  ' * 2
            ObjData(Object).Gemas = Val(Leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(Leer.GetValue("OBJ" & Object, "Diamantes"))

        End If

        '--------------------

        If ObjData(Object).SubTipo = OBJTYPE_CASCO Then

            ObjData(Object).CascoAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            '[MerLiNz:6]
            ObjData(Object).Gemas = Val(Leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(Leer.GetValue("OBJ" & Object, "Diamantes"))
            '[\END]
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))  '* 2

        End If

        If ObjData(Object).SubTipo = OBJTYPE_ALAS Then
            ObjData(Object).AlasAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            '[MerLiNz:6]
            ObjData(Object).Gemas = Val(Leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(Leer.GetValue("OBJ" & Object, "Diamantes"))
            '[\END]
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))  '* 2

        End If

        '[GAU]
        If ObjData(Object).SubTipo = OBJTYPE_BOTA Then
            ObjData(Object).Botas = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))  ' * 2

        End If

        '[GAU]
        ObjData(Object).Ropaje = Val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
        ObjData(Object).HechizoIndex = Val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))

        If ObjData(Object).ObjType = OBJTYPE_WEAPON Then
            ObjData(Object).WeaponAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).Apu�ala = Val(Leer.GetValue("OBJ" & Object, "Apu�ala"))
            ObjData(Object).Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))  ' * 2
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
            ObjData(Object).proyectil = Val(Leer.GetValue("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = Val(Leer.GetValue("OBJ" & Object, "Municiones"))
            '[MerLiNz:6]
            ObjData(Object).Gemas = Val(Leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(Leer.GetValue("OBJ" & Object, "Diamantes"))
            '[\END]
            ObjData(Object).SkArma = Val(Leer.GetValue("OBJ" & Object, "SKARMA"))
            ObjData(Object).SkArco = Val(Leer.GetValue("OBJ" & Object, "SKARCO"))

        End If

        If ObjData(Object).ObjType = OBJTYPE_ARMOUR Then
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))  ' * 2
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
            '[MerLiNz:6]
            ObjData(Object).Gemas = Val(Leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(Leer.GetValue("OBJ" & Object, "Diamantes"))
            'pluto:2.10
            ObjData(Object).ObjetoClan = Leer.GetValue("OBJ" & Object, "ObjetoClan")

            '[\END]
        End If

        If ObjData(Object).ObjType = OBJTYPE_HERRAMIENTAS Then
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))  '* 2
            '[MerLiNz:6]
            ObjData(Object).Gemas = Val(Leer.GetValue("OBJ" & Object, "Gemas"))
            ObjData(Object).Diamantes = Val(Leer.GetValue("OBJ" & Object, "Diamantes"))

            '[\END]
        End If

        If ObjData(Object).ObjType = OBJTYPE_INSTRUMENTOS Then
            ObjData(Object).Snd1 = Val(Leer.GetValue("OBJ" & Object, "SND1"))
            ObjData(Object).Snd2 = Val(Leer.GetValue("OBJ" & Object, "SND2"))
            ObjData(Object).Snd3 = Val(Leer.GetValue("OBJ" & Object, "SND3"))
            ObjData(Object).MinInt = Val(Leer.GetValue("OBJ" & Object, "MinInt"))

        End If

        ObjData(Object).LingoteIndex = Val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))

        If ObjData(Object).ObjType = 31 Or ObjData(Object).ObjType = 23 Then
            ObjData(Object).MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))

        End If

        ObjData(Object).MineralIndex = Val(Leer.GetValue("OBJ" & Object, "MineralIndex"))

        ObjData(Object).MaxHP = Val(Leer.GetValue("OBJ" & Object, "MaxHP"))
        ObjData(Object).MinHP = Val(Leer.GetValue("OBJ" & Object, "MinHP"))

        ObjData(Object).Mujer = Val(Leer.GetValue("OBJ" & Object, "Mujer"))
        ObjData(Object).Hombre = Val(Leer.GetValue("OBJ" & Object, "Hombre"))

        ObjData(Object).MinHam = Val(Leer.GetValue("OBJ" & Object, "MinHam"))
        ObjData(Object).MinSed = Val(Leer.GetValue("OBJ" & Object, "MinAgu"))

        'pluto:7.0
        ObjData(Object).MinDef = Val(Leer.GetValue("OBJ" & Object, "MINDEF"))
        ObjData(Object).MaxDef = Val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
        ObjData(Object).Defmagica = Val(Leer.GetValue("OBJ" & Object, "DEFMAGICA"))
        'nati:agrego DefCuerpo
        ObjData(Object).Defcuerpo = Val(Leer.GetValue("OBJ" & Object, "DEFCUERPO"))
        ObjData(Object).Drop = Val(Leer.GetValue("OBJ" & Object, "DROP"))

        'ObjData(Object).Defproyectil = val(Leer.GetValue("OBJ" & Object, "DEFPROYECTIL"))

        ObjData(Object).Respawn = Val(Leer.GetValue("OBJ" & Object, "ReSpawn"))

        ObjData(Object).RazaEnana = Val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
        ObjData(Object).razaelfa = Val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
        ObjData(Object).razavampiro = Val(Leer.GetValue("OBJ" & Object, "Razavampiro"))
        ObjData(Object).razaorca = Val(Leer.GetValue("OBJ" & Object, "Razaorca"))
        ObjData(Object).razahumana = Val(Leer.GetValue("OBJ" & Object, "Razahumana"))

        ObjData(Object).valor = Val(Leer.GetValue("OBJ" & Object, "Valor"))
        ObjData(Object).nocaer = Val(Leer.GetValue("OBJ" & Object, "nocaer"))
        ObjData(Object).objetoespecial = Val(Leer.GetValue("OBJ" & Object, "objetoespecial"))

        ObjData(Object).Crucial = Val(Leer.GetValue("OBJ" & Object, "Crucial"))

        ObjData(Object).Cerrada = Val(Leer.GetValue("OBJ" & Object, "abierta"))

        If ObjData(Object).Cerrada = 1 Then
            ObjData(Object).Llave = Val(Leer.GetValue("OBJ" & Object, "Llave"))
            ObjData(Object).Clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))

        End If

        If ObjData(Object).ObjType = OBJTYPE_PUERTAS Or ObjData(Object).ObjType = OBJTYPE_BOTELLAVACIA Or ObjData( _
           Object).ObjType = OBJTYPE_BOTELLALLENA Then
            ObjData(Object).IndexAbierta = Val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
            ObjData(Object).IndexCerrada = Val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
            ObjData(Object).IndexCerradaLlave = Val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))

        End If

        'Puertas y llaves
        ObjData(Object).Clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))

        ObjData(Object).texto = Leer.GetValue("OBJ" & Object, "Texto")
        ObjData(Object).GrhSecundario = Val(Leer.GetValue("OBJ" & Object, "VGrande"))

        ObjData(Object).Agarrable = Val(Leer.GetValue("OBJ" & Object, "Agarrable"))
        ObjData(Object).ForoID = Leer.GetValue("OBJ" & Object, "ID")

        Dim i As Integer, tStr As String

        For i = 1 To NUMCLASES

            tStr = Leer.GetValue("OBJ" & Object, "CP" & i)

            If tStr <> "" Then
                tStr = mid$(tStr, 1, Len(tStr) - 1)
                tStr = Right$(tStr, Len(tStr) - 1)
            End If

            ObjData(Object).ClaseProhibida(i) = tStr
        Next

        ObjData(Object).Resistencia = Val(Leer.GetValue("OBJ" & Object, "Resistencia"))

        'Pociones
        If ObjData(Object).ObjType = 11 Then
            ObjData(Object).TipoPocion = Val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
            ObjData(Object).MaxModificador = Val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
            ObjData(Object).MinModificador = Val(Leer.GetValue("OBJ" & Object, "MinModificador"))
            ObjData(Object).DuracionEfecto = Val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))

        End If

        ObjData(Object).SkCarpinteria = Val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))  '* 2

        If ObjData(Object).SkCarpinteria > 0 Then ObjData(Object).Madera = Val(Leer.GetValue("OBJ" & Object, "Madera"))

        If ObjData(Object).ObjType = OBJTYPE_BARCOS Then
            ObjData(Object).MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))

        End If

        If ObjData(Object).ObjType = OBJTYPE_FLECHAS Then
            ObjData(Object).MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))

        End If

        'Bebidas
        ObjData(Object).MinSta = Val(Leer.GetValue("OBJ" & Object, "MinST"))
        ObjData(Object).razavampiro = Val(Leer.GetValue("OBJ" & Object, "razavampiro"))
        'pluto:6.0A----
        ObjData(Object).Cregalos = Val(Leer.GetValue("OBJ" & Object, "Cregalos"))
        ObjData(Object).Pregalo = Val(Leer.GetValue("OBJ" & Object, "Pregalo"))
        
                'BORRAR
        
            If ObjData(Object).LingH > 0 Or ObjData(Object).LingP > 0 Or ObjData(Object).LingO > 0 Or ObjData(Object).Madera > 0 Or ObjData(Object).Gemas > 0 Or ObjData(Object).Diamantes > 0 Then
            
                NumTrabajo = NumTrabajo + 1
                ReDim Preserve DataTrabajo(1 To NumTrabajo) As Integer
                DataTrabajo(NumTrabajo) = Object

            End If
            'BORRAR

    Next Object

    Set Leer = Nothing

End Sub

