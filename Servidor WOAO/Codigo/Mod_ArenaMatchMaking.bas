Attribute VB_Name = "Mod_ArenaMatchMaking"
Option Explicit

Private QueueRanked(eRank.e_BRONCE To eRank.e_LAST - 1) As New Collection

Public Sub LogicAllQueue()

    Dim i         As Long, J As Long
    Dim DestUser1 As Long
    Dim DestUser2 As Long

    For i = eRank.e_BRONCE To eRank.e_LAST - 1

        With QueueRanked(i)
            If .Count > 1 Then

                For J = 2 To .Count Step 2 ' recorro de 2 en 2
                    ' Slot Primario.
                    ' j - 1
                    
                    ' Slot Secundario.
                    ' j
                    DestUser1 = .Item(J - 1)
                    DestUser2 = .Item(J)
                    
                    If SendToUsersInQueue(i, DestUser1, DestUser2) Then
                        ' Slot Primario.
                        Call RemoveUserQueue(DestUser1)
                    
                        ' Slot Secundario.
                        Call RemoveUserQueue(DestUser2)
                    End If
                Next J

            End If
        End With
    Next i

End Sub

Public Sub AddUserQueue(ByVal Userindex As Integer)

    With UserList(Userindex)
    
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "Estas muerto.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If Not GetBattleArenaActive() Then
            Call WriteConsoleMsg(Userindex, "Las Ranked se encuentran desactivadas.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If .flags.QueueArena > 0 Then
            Call WriteConsoleMsg(Userindex, "Ya te encuentras en busqueda de una ranked.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If .flags.ArenaBattleSlot > 0 Then
            Call WriteConsoleMsg(Userindex, "Ya te encuentras en una batalla rankeada.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        QueueRanked(GetUserRank(Userindex)).Add Userindex
        .flags.QueueArena = 1
        Call WriteConsoleMsg(Userindex, "Colocado en la busqueda de una ranked en " & GetUserRankString(Userindex) & ".", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
            
    End With

End Sub

Public Sub RemoveUserQueue(ByVal Userindex As Integer)

    Dim i As Long

    With QueueRanked(GetUserRank(Userindex))

        For i = 1 To .Count
            If .Item(i) = Userindex Then
                Call .Remove(i)
                Exit For
            End If
        Next i

    End With

End Sub

