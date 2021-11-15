Attribute VB_Name = "modQuest"
'[KEVIN]
Public Pista As String

Public Sub HacerQuest(ByVal UserIndex As Integer)
On Error GoTo errorhandler

Dim NpcPos As WorldPos
Dim NpcPos2 As WorldPos
Dim NpcPos3 As WorldPos
Dim NpcPos4 As WorldPos
Dim NpcPos5 As WorldPos
Dim CoorD As WorldPos
Dim RNumber As Integer

If EsNewbie(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Los newbies no pueden realizar estas quests!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If
    
If UserList(UserIndex).Quest.EnQuest = 1 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes terminar primero la quest que estás realizando para empezar otra!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Quest.Quest = NumeroQuests Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has realizados todas las quest!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

UserList(UserIndex).Quest.Quest = UserList(UserIndex).Quest.Quest + 1
UserList(UserIndex).Quest.EnQuest = 1

Select Case Quests(UserList(UserIndex).Quest.Quest).Objetivo
    Case 1
        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes matar " & Quests(UserList(UserIndex).Quest.Quest).NPCs & " npcs para recivir tu recompensa!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Case 2
        If UserList(UserIndex).Faccion.ArmadaReal = 0 And UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes matar " & Quests(UserList(UserIndex).Quest.Quest).UsUaRiOs & " usuarios para recivir tu recompensa!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
        ElseIf UserList(UserIndex).Faccion.ArmadaReal > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes matar " & Quests(UserList(UserIndex).Quest.Quest).Criminales & " criminales para recivir tu recompensa!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
        ElseIf UserList(UserIndex).Faccion.FuerzasCaos > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes matar " & Quests(UserList(UserIndex).Quest.Quest).Ciudadanos & " ciudadanos para recivir tu recompensa!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
        End If
    Case 3
        NpcPos.Map = val(ReadField(1, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas1"), 45))
        NpcPos.X = val(ReadField(2, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas1"), 45))
        NpcPos.Y = val(ReadField(3, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas1"), 45))
        
        NpcPos2.Map = val(ReadField(1, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas2"), 45))
        NpcPos2.X = val(ReadField(2, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas2"), 45))
        NpcPos2.Y = val(ReadField(3, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas2"), 45))
        
        NpcPos3.Map = val(ReadField(1, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas3"), 45))
        NpcPos3.X = val(ReadField(2, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas3"), 45))
        NpcPos3.Y = val(ReadField(3, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas3"), 45))
        
        NpcPos4.Map = val(ReadField(1, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas4"), 45))
        NpcPos4.X = val(ReadField(2, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas4"), 45))
        NpcPos4.Y = val(ReadField(3, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas4"), 45))
        
        NpcPos5.Map = val(ReadField(1, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas5"), 45))
        NpcPos5.X = val(ReadField(2, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas5"), 45))
        NpcPos5.Y = val(ReadField(3, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas5"), 45))
        
        RNumber = RandomNumber(1, 5)
        
        If RNumber = 1 Then
            CoorD = NpcPos
        ElseIf RNumber = 2 Then
            CoorD = NpcPos2
        ElseIf RNumber = 3 Then
            CoorD = NpcPos3
        ElseIf RNumber = 4 Then
            CoorD = NpcPos4
        ElseIf RNumber = 5 Then
            CoorD = NpcPos5
        End If
        
        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes encontrar a mi amigo npc para recivir tu recompensa! Pistas: Se pude encontrar en lugares característicos del juego, pero apresúrate porque si otro que está haciendo este tipo de quest lo encuentra primero puedes perderlo, en tal caso deberás clikearme y poner /MERINDO" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
        Call SpawnNpc(Quests(UserList(UserIndex).Quest.Quest).AmigoNpc, CoorD, True, False)
        
    Case 4
        NpcPos.Map = val(ReadField(1, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas1"), 45))
        NpcPos.X = val(ReadField(2, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas1"), 45))
        NpcPos.Y = val(ReadField(3, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas1"), 45))
        
        NpcPos2.Map = val(ReadField(1, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas2"), 45))
        NpcPos2.X = val(ReadField(2, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas2"), 45))
        NpcPos2.Y = val(ReadField(3, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas2"), 45))
        
        NpcPos3.Map = val(ReadField(1, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas3"), 45))
        NpcPos3.X = val(ReadField(2, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas3"), 45))
        NpcPos3.Y = val(ReadField(3, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas3"), 45))
        
        NpcPos4.Map = val(ReadField(1, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas4"), 45))
        NpcPos4.X = val(ReadField(2, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas4"), 45))
        NpcPos4.Y = val(ReadField(3, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas4"), 45))
        
        NpcPos5.Map = val(ReadField(1, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas5"), 45))
        NpcPos5.X = val(ReadField(2, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas5"), 45))
        NpcPos5.Y = val(ReadField(3, GetVar(DatPath & "Quests.dat", "Quest" & UserList(UserIndex).Quest.Quest, "Coordenadas5"), 45))
        
        RNumber = RandomNumber(1, 5)
        
        If RNumber = 1 Then
            CoorD = NpcPos
            Pista = Quests(UserList(UserIndex).Quest.Quest).Coordenadas1
        ElseIf RNumber = 2 Then
            CoorD = NpcPos2
            Pista = Quests(UserList(UserIndex).Quest.Quest).Coordenadas2
        ElseIf RNumber = 3 Then
            CoorD = NpcPos3
            Pista = Quests(UserList(UserIndex).Quest.Quest).Coordenadas3
        ElseIf RNumber = 4 Then
            CoorD = NpcPos4
            Pista = Quests(UserList(UserIndex).Quest.Quest).Coordenadas4
        ElseIf RNumber = 5 Then
            CoorD = NpcPos5
            Pista = Quests(UserList(UserIndex).Quest.Quest).Coordenadas5
        End If
        
        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes matar al npc que se encuentra en las coordenadas " & Pista & " para recivir tu recompensa! PD: Si no te apresuras y lo mata otro usuario perderás esta quest, en tal caso deberás clikearme y poner /MERINDO" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    
        Call SpawnNpc(Quests(UserList(UserIndex).Quest.Quest).CriaturaIndex, NpcPos, True, False)
End Select

errorhandler:
Call LogError(UserList(UserIndex).Name & "= Error en HacerQuest")
    
End Sub

Public Sub RecibirRecompensaQuest(ByVal UserIndex As Integer)
On Error GoTo errorhandler

Dim RObj As Obj

If EsNewbie(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Los newbies no pueden realizar estas quests!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Quest.Quest <= 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No has empezado ninguna Quest!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Quest.EnQuest = 1 Then
    If UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest Then
        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recivido tu recompensa por esta Quest!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
        Exit Sub
    End If
    
    Select Case Quests(UserList(UserIndex).Quest.Quest).Objetivo
        Case 1
            If UserList(UserIndex).Stats.NPCsMuertos >= Quests(UserList(UserIndex).Quest.Quest).NPCs Then
            
                If Quests(UserList(UserIndex).Quest.Quest).DaExp = 1 Then
                    Call AddtoVar(UserList(UserIndex).Stats.Exp, Quests(UserList(UserIndex).Quest.Quest).Exp, MAXEXP)
                    Call CheckUserLevel(UserIndex)
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido " & Quests(UserList(UserIndex).Quest.Quest).Exp & " puntos de experiencia!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                    UserList(UserIndex).Quest.RealizoQuest = 0
                    UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                    UserList(UserIndex).Quest.EnQuest = 0
                End If
                
                If Quests(UserList(UserIndex).Quest.Quest).DaOro = 1 Then
                    Call AddtoVar(UserList(UserIndex).Stats.GLD, Quests(UserList(UserIndex).Quest.Quest).Oro, MAXORO)
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido " & Quests(UserList(UserIndex).Quest.Quest).Oro & " monedas de oro!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                    UserList(UserIndex).Quest.RealizoQuest = 0
                    UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                    UserList(UserIndex).Quest.EnQuest = 0
                End If
                
                If Quests(UserList(UserIndex).Quest.Quest).DaObj = 1 Then
                    
                    RObj.Amount = 1
                    RObj.ObjIndex = Quests(UserList(UserIndex).Quest.Quest).Obj
                    
                    If Not MeterItemEnInventario(UserIndex, RObj) Then
                        Call TirarItemAlPiso(UserList(UserIndex).Pos, RObj)
                    End If
                    
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido un objeto!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                    UserList(UserIndex).Quest.RealizoQuest = 0
                    UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                    UserList(UserIndex).Quest.EnQuest = 0
                End If
                
            Else
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Todavía no has realizado el objetivo!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
            End If
        Case 2
            If UserList(UserIndex).Faccion.ArmadaReal = 0 And UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                If UserList(UserIndex).Faccion.CiudadanosMatados + UserList(UserIndex).Faccion.CriminalesMatados >= Quests(UserList(UserIndex).Quest.Quest).UsUaRiOs Then
                
                    If Quests(UserList(UserIndex).Quest.Quest).DaExp = 1 Then
                        Call AddtoVar(UserList(UserIndex).Stats.Exp, Quests(UserList(UserIndex).Quest.Quest).Exp, MAXEXP)
                        Call CheckUserLevel(UserIndex)
                        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido " & Quests(UserList(UserIndex).Quest.Quest).Exp & " puntos de experiencia!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                        UserList(UserIndex).Quest.RealizoQuest = 0
                        UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                        UserList(UserIndex).Quest.EnQuest = 0
                    End If
                    
                    If Quests(UserList(UserIndex).Quest.Quest).DaOro = 1 Then
                        Call AddtoVar(UserList(UserIndex).Stats.GLD, Quests(UserList(UserIndex).Quest.Quest).Oro, MAXORO)
                        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido " & Quests(UserList(UserIndex).Quest.Quest).Oro & " monedas de oro!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                        UserList(UserIndex).Quest.RealizoQuest = 0
                        UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                        UserList(UserIndex).Quest.EnQuest = 0
                    End If
                    
                    If Quests(UserList(UserIndex).Quest.Quest).DaObj = 1 Then
                        
                        RObj.Amount = 1
                        RObj.ObjIndex = Quests(UserList(UserIndex).Quest.Quest).Obj
                        
                        If Not MeterItemEnInventario(UserIndex, RObj) Then
                            Call TirarItemAlPiso(UserList(UserIndex).Pos, RObj)
                        End If
                        
                        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido un objeto!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                        UserList(UserIndex).Quest.RealizoQuest = 0
                        UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                        UserList(UserIndex).Quest.EnQuest = 0
                    End If
                    
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Todavía no has realizado el objetivo!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                End If
                
            ElseIf UserList(UserIndex).Faccion.ArmadaReal > 0 Then
                    If UserList(UserIndex).Faccion.CriminalesMatados >= Quests(UserList(UserIndex).Quest.Quest).Criminales Then
                    
                        If Quests(UserList(UserIndex).Quest.Quest).DaExp = 1 Then
                            Call AddtoVar(UserList(UserIndex).Stats.Exp, Quests(UserList(UserIndex).Quest.Quest).Exp, MAXEXP)
                            Call CheckUserLevel(UserIndex)
                            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido " & Quests(UserList(UserIndex).Quest.Quest).Exp & " puntos de experiencia!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                            UserList(UserIndex).Quest.RealizoQuest = 0
                            UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                            UserList(UserIndex).Quest.EnQuest = 0
                        End If
                        
                        If Quests(UserList(UserIndex).Quest.Quest).DaOro = 1 Then
                            Call AddtoVar(UserList(UserIndex).Stats.GLD, Quests(UserList(UserIndex).Quest.Quest).Oro, MAXORO)
                            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido " & Quests(UserList(UserIndex).Quest.Quest).Oro & " monedas de oro!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                            UserList(UserIndex).Quest.RealizoQuest = 0
                            UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                            UserList(UserIndex).Quest.EnQuest = 0
                        End If
                        
                        If Quests(UserList(UserIndex).Quest.Quest).DaObj = 1 Then
                            
                            RObj.Amount = 1
                            RObj.ObjIndex = Quests(UserList(UserIndex).Quest.Quest).Obj
                            
                            If Not MeterItemEnInventario(UserIndex, RObj) Then
                                Call TirarItemAlPiso(UserList(UserIndex).Pos, RObj)
                            End If
                            
                            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido un objeto!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                            UserList(UserIndex).Quest.RealizoQuest = 0
                            UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                            UserList(UserIndex).Quest.EnQuest = 0
                        End If
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Todavía no has realizado el objetivo!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                    End If
                    
            ElseIf UserList(UserIndex).Faccion.FuerzasCaos > 0 Then
                    If UserList(UserIndex).Faccion.CiudadanosMatados >= Quests(UserList(UserIndex).Quest.Quest).Ciudadanos Then
                    
                        If Quests(UserList(UserIndex).Quest.Quest).DaExp = 1 Then
                            Call AddtoVar(UserList(UserIndex).Stats.Exp, Quests(UserList(UserIndex).Quest.Quest).Exp, MAXEXP)
                            Call CheckUserLevel(UserIndex)
                            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido " & Quests(UserList(UserIndex).Quest.Quest).Exp & " puntos de experiencia!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                            UserList(UserIndex).Quest.RealizoQuest = 0
                            UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                            UserList(UserIndex).Quest.EnQuest = 0
                        End If
                        
                        If Quests(UserList(UserIndex).Quest.Quest).DaOro = 1 Then
                            Call AddtoVar(UserList(UserIndex).Stats.GLD, Quests(UserList(UserIndex).Quest.Quest).Oro, MAXORO)
                            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido " & Quests(UserList(UserIndex).Quest.Quest).Oro & " monedas de oro!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                            UserList(UserIndex).Quest.RealizoQuest = 0
                            UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                            UserList(UserIndex).Quest.EnQuest = 0
                        End If
                        
                        If Quests(UserList(UserIndex).Quest.Quest).DaObj = 1 Then
                            
                            RObj.Amount = 1
                            RObj.ObjIndex = Quests(UserList(UserIndex).Quest.Quest).Obj
                            
                            If Not MeterItemEnInventario(UserIndex, RObj) Then
                                Call TirarItemAlPiso(UserList(UserIndex).Pos, RObj)
                            End If
                            
                            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido un objeto!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                            UserList(UserIndex).Quest.RealizoQuest = 0
                            UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                            UserList(UserIndex).Quest.EnQuest = 0
                        End If
                        
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Todavía no has realizado el objetivo!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                    End If
            End If
            
        Case 3
            If UserList(UserIndex).Quest.RealizoQuest = 1 Then
                        If Quests(UserList(UserIndex).Quest.Quest).DaExp = 1 Then
                            Call AddtoVar(UserList(UserIndex).Stats.Exp, Quests(UserList(UserIndex).Quest.Quest).Exp, MAXEXP)
                            Call CheckUserLevel(UserIndex)
                            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido " & Quests(UserList(UserIndex).Quest.Quest).Exp & " puntos de experiencia!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                            UserList(UserIndex).Quest.RealizoQuest = 0
                            UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                            UserList(UserIndex).Quest.EnQuest = 0
                            Call QuitarNPC(UserList(UserIndex).Flags.NpcAmigo)
                        End If
                        
                        If Quests(UserList(UserIndex).Quest.Quest).DaOro = 1 Then
                            Call AddtoVar(UserList(UserIndex).Stats.GLD, Quests(UserList(UserIndex).Quest.Quest).Oro, MAXORO)
                            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido " & Quests(UserList(UserIndex).Quest.Quest).Oro & " monedas de oro!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                            UserList(UserIndex).Quest.RealizoQuest = 0
                            UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                            UserList(UserIndex).Quest.EnQuest = 0
                            Call QuitarNPC(UserList(UserIndex).Flags.NpcAmigo)
                        End If
                        
                        If Quests(UserList(UserIndex).Quest.Quest).DaObj = 1 Then
                            
                            RObj.Amount = 1
                            RObj.ObjIndex = Quests(UserList(UserIndex).Quest.Quest).Obj
                            
                            If Not MeterItemEnInventario(UserIndex, RObj) Then
                                Call TirarItemAlPiso(UserList(UserIndex).Pos, RObj)
                            End If
                            
                            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido un objeto!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                            UserList(UserIndex).Quest.RealizoQuest = 0
                            UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                            UserList(UserIndex).Quest.EnQuest = 0
                            Call QuitarNPC(UserList(UserIndex).Flags.NpcAmigo)
                        End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Todavía no has logrado el objetivo!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
            End If
        Case 4
            If UserList(UserIndex).Quest.RealizoQuest = 1 Then
                        If Quests(UserList(UserIndex).Quest.Quest).DaExp = 1 Then
                            Call AddtoVar(UserList(UserIndex).Stats.Exp, Quests(UserList(UserIndex).Quest.Quest).Exp, MAXEXP)
                            Call CheckUserLevel(UserIndex)
                            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido " & Quests(UserList(UserIndex).Quest.Quest).Exp & " puntos de experiencia!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                            UserList(UserIndex).Quest.RealizoQuest = 0
                            UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                            UserList(UserIndex).Quest.EnQuest = 0
                        End If
                        
                        If Quests(UserList(UserIndex).Quest.Quest).DaOro = 1 Then
                            Call AddtoVar(UserList(UserIndex).Stats.GLD, Quests(UserList(UserIndex).Quest.Quest).Oro, MAXORO)
                            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido " & Quests(UserList(UserIndex).Quest.Quest).Oro & " monedas de oro!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                            UserList(UserIndex).Quest.RealizoQuest = 0
                            UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                            UserList(UserIndex).Quest.EnQuest = 0
                        End If
                        
                        If Quests(UserList(UserIndex).Quest.Quest).DaObj = 1 Then
                            
                            RObj.Amount = 1
                            RObj.ObjIndex = Quests(UserList(UserIndex).Quest.Quest).Obj
                            
                            If Not MeterItemEnInventario(UserIndex, RObj) Then
                                Call TirarItemAlPiso(UserList(UserIndex).Pos, RObj)
                            End If
                            
                            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Como recompensa has recibido un objeto!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                            UserList(UserIndex).Quest.RealizoQuest = 0
                            UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
                            UserList(UserIndex).Quest.EnQuest = 0
                        End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Todavía no has logrado el objetivo!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
            End If
    End Select
Else
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No te encuentras en ninguna Quest actualmente!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
End If

Call SendUserStatsBox(UserIndex)

errorhandler:
Call LogError(UserList(UserIndex).Name & "= Error en RecivirRecompensaQuest")
                
End Sub

Public Sub CheckNpcAmigo(ByVal UserIndex As Integer)

If UserList(UserIndex).Quest.EnQuest = 1 Then
    If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype = NPCTYPE_AMIGOQUEST Then
        If Quests(UserList(UserIndex).Quest.Quest).Objetivo = 3 Then
            UserList(UserIndex).Quest.RealizoQuest = 1
            UserList(UserIndex).Flags.NpcAmigo = UserList(UserIndex).Flags.TargetNpc
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Felicitaciones, me has encontrado, ahora debes volver con mi compañero por tu recompenza!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
        Else
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Yo correspondo a otra quest!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
        End If
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Este Npc no es de la quest jajaja!" & FONTTYPE_INFO)
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "||No has comenzado ninguna quest!" & FONTTYPE_INFO)
End If

End Sub

Public Sub SendInfoQuest(ByVal UserIndex As Integer)

If EsNewbie(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Los newbies no pueden realizar estas quests!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

Select Case Quests(UserList(UserIndex).Quest.Quest).Objetivo
    Case 1
        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes matar " & Quests(UserList(UserIndex).Quest.Quest).NPCs & " npcs para recivir tu recompensa!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Case 2
        If UserList(UserIndex).Faccion.ArmadaReal = 0 And UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes matar " & Quests(UserList(UserIndex).Quest.Quest).UsUaRiOs & " usuarios para recivir tu recompensa!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
        ElseIf UserList(UserIndex).Faccion.ArmadaReal > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes matar " & Quests(UserList(UserIndex).Quest.Quest).Criminales & " criminales para recivir tu recompensa!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
        ElseIf UserList(UserIndex).Faccion.FuerzasCaos > 0 Then
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes matar " & Quests(UserList(UserIndex).Quest.Quest).Ciudadanos & " ciudadanos para recivir tu recompensa!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
        End If
    Case 3
        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes encontrar a mi amigo npc para recivir tu recompensa! Pistas: Se pude encontrar en lugares característicos del juego, pero apresúrate porque si otro que está haciendo este tipo de quest lo encuentra primero puedes perderlo, en tal caso deberás clikearme y poner /MERINDO" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
         
    Case 4
        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Debes matar al npc que se encuentra en las coordenadas " & Pista & " para recivir tu recompensa! PD: Si no te apresuras y lo mata otro usuario perderás esta quest, en tal caso deberás clikearme y poner /MERINDO" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
End Select

End Sub

Public Sub UserSeRinde(ByVal UserIndex As Integer)
On Error GoTo errorhandler


If EsNewbie(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Los newbies no pueden realizar estas quests!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

UserList(UserIndex).Quest.RealizoQuest = 0
UserList(UserIndex).Quest.Recompensa = UserList(UserIndex).Quest.Quest
UserList(UserIndex).Quest.EnQuest = 0

Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Te has rendido por lo tanto no has conseguido la recompensa, pero puedes continuar con la siguiente quest." & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))

errorhandler:
Call LogError(UserList(UserIndex).Name & "= Error en UserSeRinde")

End Sub

Public Sub CheckNpcEnemigo(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

If UserList(UserIndex).Quest.EnQuest = 1 Then
    If Npclist(NpcIndex).Flags.DeQuest = 1 Then
        If Quests(UserList(UserIndex).Quest.Quest).Objetivo = 4 Then
            UserList(UserIndex).Quest.RealizoQuest = 1
            Call SendData(ToIndex, UserIndex, 0, "||Has encontrado y eliminado a la criatura de la quest ahora ve por tu recompensa!" & FONTTYPE_FIGHT)
        End If
    End If
End If

End Sub
'[/KEVIN]
