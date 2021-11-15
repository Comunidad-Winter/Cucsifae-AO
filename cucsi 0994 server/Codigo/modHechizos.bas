Attribute VB_Name = "modHechizos"
'Argentum Online 0.9.0.4
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
Option Explicit

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(UserIndex).Flags.Invisible = 1 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim Daño As Integer
Dim defMagica As Integer

If Hechizos(Spell).SubeHP = 1 Then

    Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP + Daño
    If UserList(UserIndex).Stats.MinHP > UserList(UserIndex).Stats.MaxHP Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
    
    Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha restaurado " & Daño & " puntos de vida." & FONTTYPE_FIGHT)

ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    Daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    '[KEVIN]*********************************
    If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).TipoAnillo = 6 Then
            Daño = Daño - ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MaxModificador
        End If
    End If
    
    If UserList(UserIndex).Invent.Anillo2EqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).TipoAnillo = 6 Then
            Daño = Daño - ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).MaxModificador
        End If
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).DefMagia = 1 Then
            defMagica = RandomNumber(ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MagiaMinDef, ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MagiaMaxDef)
            Daño = Daño - defMagica
        End If
    End If
    
    If Daño < 1 Then Daño = 1
    '[/KEVIN]*****************************************

    If UserList(UserIndex).Flags.Privilegios = 0 Then UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MinHP - Daño
    
    Call SendData(ToIndex, UserIndex, 0, "||" & Npclist(NpcIndex).Name & " te ha quitado " & Daño & " puntos de vida." & FONTTYPE_FIGHT)
    
    'Muere
    If UserList(UserIndex).Stats.MinHP < 1 Then
        UserList(UserIndex).Stats.MinHP = 0
        Call UserDie(UserIndex)
    End If
    
End If

If Hechizos(Spell).Paraliza = 1 Then
     If UserList(UserIndex).Flags.Paralizado = 0 Then
          
          '[KEVIN]
          If EsDios(UserList(UserIndex).Name) Or EsSemiDios(UserList(UserIndex).Name) Then Exit Sub
          
          If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).TipoAnillo = 7 Then
            Exit Sub
            End If
          End If
          
          If UserList(UserIndex).Invent.Anillo2EqpObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).TipoAnillo = 7 Then
            Exit Sub
            End If
          End If
          '[/KEVIN]
     
          UserList(UserIndex).Flags.Paralizado = 1
          UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(Spell).WAV)
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
          Call SendData(ToIndex, UserIndex, 0, "PARADOK")
     End If
End If

Call SendUserStatsBox(UserIndex)

End Sub

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

On Error GoTo errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
errhandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, UserIndex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No tenes espacio para mas hechizos." & FONTTYPE_INFO)
    Else
        UserList(UserIndex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, UserIndex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
    End If
Else
    Call SendData(ToIndex, UserIndex, 0, "||Ya tenes ese hechizo." & FONTTYPE_INFO)
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal s As String, ByVal UserIndex As Integer)
On Error Resume Next

    Dim ind As String
    ind = UserList(UserIndex).Char.CharIndex
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbCyan & "°" & s & "°" & ind)
    Exit Sub
End Sub
Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean

If UserList(UserIndex).Flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.Map = UserList(UserIndex).Flags.TargetMap
    wp2.X = UserList(UserIndex).Flags.TargetX
    wp2.Y = UserList(UserIndex).Flags.TargetY
    
    If Distancia(UserList(UserIndex).Pos, wp2) > 18 Then
            'UserList(UserIndex).Flags.AdministrativeBan = 1
            'Call SendData(ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
            Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de lanzar un spell desde otro mapa.")
            'Call CloseSocket(UserIndex)
            Exit Function
    End If
    
    If UserList(UserIndex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
        If UserList(UserIndex).Stats.UserSkills(Magia) >= Hechizos(HechizoIndex).MinSkill Then
            PuedeLanzar = (UserList(UserIndex).Stats.MinSta > 0)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No tenes suficientes puntos de magia para lanzar este hechizo." & FONTTYPE_INFO)
            PuedeLanzar = False
        End If
    Else
            Call SendData(ToIndex, UserIndex, 0, "||No tenes suficiente mana." & FONTTYPE_INFO)
            PuedeLanzar = False
    End If
Else
   Call SendData(ToIndex, UserIndex, 0, "||No podes lanzar hechizos porque estas muerto." & FONTTYPE_INFO)
   PuedeLanzar = False
End If

End Function

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)

'Call LogTarea("HechizoInvocacion")

Dim H As Integer, j As Integer, ind As Integer, Index As Integer
Dim TargetPos As WorldPos


TargetPos.Map = UserList(UserIndex).Flags.TargetMap
TargetPos.X = UserList(UserIndex).Flags.TargetX
TargetPos.Y = UserList(UserIndex).Flags.TargetY

H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).Flags.Hechizo)
    
    
For j = 1 To Hechizos(H).Cant
    
    If UserList(UserIndex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(H).NumNpc, TargetPos, True, False)
        If ind <= MAXNPCS Then
            UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas + 1
            
            Index = FreeMascotaIndex(UserIndex)
            
            UserList(UserIndex).MascotasIndex(Index) = ind
            UserList(UserIndex).MascotasType(Index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = UserIndex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            Npclist(ind).GiveGLD = 0
            
            Call FollowAmo(ind)
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(UserIndex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case uInvocacion '
       Call HechizoInvocacion(UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    Call SendUserStatsBox(UserIndex)
End If


End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)

Dim b As Boolean
Dim tIndex As Integer
tIndex = UserList(UserIndex).Flags.TargetUser

Select Case Hechizos(uh).Tipo
    Case uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(UserIndex, b)
    Case uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropUsuario(UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    Call SendUserStatsBox(UserIndex)
    Call SendUserStatsBox(UserList(UserIndex).Flags.TargetUser)
    UserList(UserIndex).Flags.TargetUser = 0
End If

End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)



Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoNPC(UserList(UserIndex).Flags.TargetNpc, uh, b, UserIndex)
    Case uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropNPC(uh, UserList(UserIndex).Flags.TargetNpc, UserIndex, b)
End Select

If b Then
    Call SubirSkill(UserIndex, Magia)
    UserList(UserIndex).Flags.TargetNpc = 0
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
    Call SendUserStatsBox(UserIndex)
End If

End Sub
Sub LanzarHechizo(Index As Integer, UserIndex As Integer)



Dim uh As Integer
Dim exito As Boolean

uh = UserList(UserIndex).Stats.UserHechizos(Index)

If PuedeLanzar(UserIndex, uh) Then
    Select Case Hechizos(uh).Target
        
        Case uUsuarios
            If UserList(UserIndex).Flags.TargetUser > 0 Then
                Call HandleHechizoUsuario(UserIndex, uh)
                '[KEVIN]
                If UserList(UserIndex).Flags.Oculto = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
                    UserList(UserIndex).Flags.Oculto = 0
                    UserList(UserIndex).Flags.Invisible = 0
                    Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "KLBPA" & UserList(UserIndex).Char.CharIndex & ",0")
                    'Exit Sub
                End If
                '[/KEVIN]
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Este hechizo actua solo sobre usuarios." & FONTTYPE_INFO)
            End If
        Case uNPC
            If UserList(UserIndex).Flags.TargetNpc > 0 Then
                Call HandleHechizoNPC(UserIndex, uh)
                '[KEVIN]
                If UserList(UserIndex).Flags.Oculto = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
                    UserList(UserIndex).Flags.Oculto = 0
                    UserList(UserIndex).Flags.Invisible = 0
                    Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "KLBPA" & UserList(UserIndex).Char.CharIndex & ",0")
                    'Exit Sub
                End If
                '[/KEVIN]
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Este hechizo solo afecta a los npcs." & FONTTYPE_INFO)
            End If
        Case uUsuariosYnpc
            If UserList(UserIndex).Flags.TargetUser > 0 Then
                Call HandleHechizoUsuario(UserIndex, uh)
                '[KEVIN]
                If UserList(UserIndex).Flags.Oculto = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
                    UserList(UserIndex).Flags.Oculto = 0
                    UserList(UserIndex).Flags.Invisible = 0
                    Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "KLBPA" & UserList(UserIndex).Char.CharIndex & ",0")
                    'Exit Sub
                End If
                '[/KEVIN]
            ElseIf UserList(UserIndex).Flags.TargetNpc > 0 Then
                Call HandleHechizoNPC(UserIndex, uh)
                '[KEVIN]
                If UserList(UserIndex).Flags.Oculto = 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
                    UserList(UserIndex).Flags.Oculto = 0
                    UserList(UserIndex).Flags.Invisible = 0
                    Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "KLBPA" & UserList(UserIndex).Char.CharIndex & ",0")
                    'Exit Sub
                End If
                '[/KEVIN]
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Target invalido." & FONTTYPE_INFO)
            End If
        Case uTerreno
            Call HandleHechizoTerreno(UserIndex, uh)
            '[KEVIN]
            If UserList(UserIndex).Flags.Oculto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
                UserList(UserIndex).Flags.Oculto = 0
                UserList(UserIndex).Flags.Invisible = 0
                Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "KLBPA" & UserList(UserIndex).Char.CharIndex & ",0")
                'Exit Sub
            End If
            '[/KEVIN]
    End Select
    
End If
                

End Sub
Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)



Dim H As Integer, TU As Integer
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).Flags.Hechizo)
TU = UserList(UserIndex).Flags.TargetUser

'[KEVIN]PARA LOS CLANES**********
If Hechizos(H).Envenena = 1 Or Hechizos(H).Maldicion = 1 Or Hechizos(H).Paraliza = 1 Or Hechizos(H).Ceguera = 1 Or Hechizos(H).Estupidez = 1 Then

If UserList(UserIndex).Stats.Matrimonio <> "" Then
    If UserList(UserIndex).Stats.Matrimonio = UserList(TU).Name Then
        If UserList(TU).Genero = "Hombre" Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar a tu esposo." & FONTTYPE_INFO)
            Exit Sub
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar a tu esposa." & FONTTYPE_INFO)
            Exit Sub
        End If
    End If
End If

If UserList(UserIndex).GuildInfo.GuildName <> "" Then
    If UserIndex <> TU Then
        If UserList(UserIndex).Flags.EnTorneo = 0 Then
            If UserList(UserIndex).GuildInfo.GuildName = UserList(TU).GuildInfo.GuildName Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar a tus comprañeros de clan." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If
    End If
End If

End If
'[/KEVIN]********************************

If Hechizos(H).Invisibilidad = 1 Then
   UserList(TU).Flags.Invisible = 1
   Call SendData(ToMap, 0, UserList(TU).Pos.Map, "KLBPA" & UserList(TU).Char.CharIndex & ",1")
   Call InfoHechizo(UserIndex)
   b = True
End If

If Hechizos(H).Envenena = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).Flags.Envenenado = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).CuraVeneno = 1 Then
        UserList(TU).Flags.Envenenado = 0
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Maldicion = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).Flags.Maldicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).RemoverMaldicion = 1 Then
        UserList(TU).Flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Bendicion = 1 Then
        UserList(TU).Flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Paraliza = 1 Then
     If UserList(TU).Flags.Paralizado = 0 Then
            If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
            
            If UserIndex <> TU Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TU)
            End If
            
            '[KEVIN]*********************
            'LO SACO PORQUE SERÍA MUCHA VENTAJA
            'If UserList(TU).Invent.AnilloEqpObjIndex > 0 Then
                'If ObjData(UserList(TU).Invent.AnilloEqpObjIndex).TipoAnillo = 7 Then
                'Exit Sub
                'End If
            'End If
            
            'If UserList(TU).Invent.Anillo2EqpObjIndex > 0 Then
                'If ObjData(UserList(TU).Invent.Anillo2EqpObjIndex).TipoAnillo = 7 Then
                'Exit Sub
                'End If
            'End If
            '[/KEVIN]*****************************
            
            UserList(TU).Flags.Paralizado = 1
            UserList(TU).Counters.Paralisis = IntervaloParalizado
            Call SendData(ToIndex, TU, 0, "PARADOK")
            Call InfoHechizo(UserIndex)
            b = True
    End If
End If

If Hechizos(H).RemoverParalisis = 1 Then
    If UserList(TU).Flags.Paralizado = 1 Then
                UserList(TU).Flags.Paralizado = 0
                Call SendData(ToIndex, TU, 0, "PARADOK")
                Call InfoHechizo(UserIndex)
                b = True
    End If
End If

If Hechizos(H).Revivir = 1 Then
    If UserList(TU).Flags.Muerto = 1 Then
        If Not Criminal(TU) Then
                If TU <> UserIndex Then
                    Call AddtoVar(UserList(UserIndex).Reputacion.NobleRep, 500, MAXREP)
                    Call SendData(ToIndex, UserIndex, 0, "||¡Los Dioses te sonrien, has ganado 500 puntos de nobleza!." & FONTTYPE_INFO)
                End If
        End If
        
        Call RevivirUsuario(TU)
    End If
    Call InfoHechizo(UserIndex)
    b = True
End If

If Hechizos(H).Ceguera = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).Flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
        Call SendData(ToIndex, TU, 0, "BLKB")
        Call InfoHechizo(UserIndex)
        b = True
End If

If Hechizos(H).Estupidez = 1 Then
        If Not PuedeAtacar(UserIndex, TU) Then Exit Sub
        If UserIndex <> TU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, TU)
        End If
        UserList(TU).Flags.Estupidez = 1
        UserList(TU).Counters.Estupidez = IntervaloParalizado
        Call SendData(ToIndex, TU, 0, "LBKL")
        Call InfoHechizo(UserIndex)
        b = True
End If

End Sub
Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)



If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).Flags.Invisible = 1
   b = True
End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        Exit Sub
   End If
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).Flags.Envenenado = 1
   b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).Flags.Envenenado = 0
   b = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        Exit Sub
   End If
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).Flags.Maldicion = 1
   b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).Flags.Maldicion = 0
   b = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
   Call InfoHechizo(UserIndex)
   Npclist(NpcIndex).Flags.Bendicion = 1
   b = True
End If

If Hechizos(hIndex).Paraliza = 1 Then
   If Npclist(NpcIndex).Flags.AfectaParalisis = 0 Then
            '[KEVIN]
            If Npclist(NpcIndex).NPCtype = 2 Or Npclist(NpcIndex).NPCtype = 7 Then Exit Sub
            '[/KEVIN]
   
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).Flags.Paralizado = 1
            Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado * 2
            b = True
   Else
      Call SendData(ToIndex, UserIndex, 0, "||El npc es inmune a este hechizo." & FONTTYPE_FIGHT)
   End If
End If

If Hechizos(hIndex).RemoverParalisis = 1 Then
   If Npclist(NpcIndex).Flags.Paralizado = 1 Then
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).Flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
   Else
      Call SendData(ToIndex, UserIndex, 0, "||El npc no esta paralizado." & FONTTYPE_FIGHT)
   End If
End If

 


End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)

Dim Daño As Integer


'Salud
If Hechizos(hIndex).SubeHP = 1 Then
    Daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex)
    Call AddtoVar(Npclist(NpcIndex).Stats.MinHP, Daño, Npclist(NpcIndex).Stats.MaxHP)
    Call SendData(ToIndex, UserIndex, 0, "||Has curado " & Daño & " puntos de salud a la criatura." & FONTTYPE_FIGHT)
    b = True
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    
    If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    
    '[KEVIN]*************************************************
    If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).TipoAnillo = 8 Then
                Daño = Daño + ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MaxModificador
                Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
        Else
                If UserList(UserIndex).Invent.Anillo2EqpObjIndex > 0 Then
                     If ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).TipoAnillo = 8 Then
                            Daño = Daño + ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).MaxModificador
                            Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
                     Else
                            Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
                     End If
                End If
        End If
    ElseIf UserList(UserIndex).Invent.Anillo2EqpObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).TipoAnillo = 8 Then
                Daño = Daño + ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).MaxModificador
                Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
            Else
                Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
            End If
    Else
        Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
    End If
    '[/KEVIN]*************************************************************
    
    Call InfoHechizo(UserIndex)
    b = True
    Call NpcAtacado(NpcIndex, UserIndex)
    If Npclist(NpcIndex).Flags.Snd2 > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Npclist(NpcIndex).Flags.Snd2)
    
    '[KEVIN]
    Call CalcularDarExp(UserIndex, NpcIndex, Daño)
    SendData ToIndex, UserIndex, 0, "||Le has causado " & Daño & " puntos de daño a la criatura!" & FONTTYPE_FIGHT
    '[/KEVIN]
    
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - Daño
    
    If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, UserIndex)
    End If
End If

End Sub
Sub InfoHechizo(ByVal UserIndex As Integer)


    Dim H As Integer
    H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).Flags.Hechizo)
    
    
    Call DecirPalabrasMagicas(Hechizos(H).PalabrasMagicas, UserIndex)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & Hechizos(H).WAV)
    
    If UserList(UserIndex).Flags.TargetUser > 0 Then
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserList(UserIndex).Flags.TargetUser).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    ElseIf UserList(UserIndex).Flags.TargetNpc > 0 Then
        Call SendData(ToPCArea, UserIndex, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos.Map, "CFX" & Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex & "," & Hechizos(H).FXgrh & "," & Hechizos(H).loops)
    End If
    
    If UserList(UserIndex).Flags.TargetUser > 0 Then
        If UserIndex <> UserList(UserIndex).Flags.TargetUser Then
            Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).HechizeroMsg & " " & UserList(UserList(UserIndex).Flags.TargetUser).Name & FONTTYPE_FIGHT)
            Call SendData(ToIndex, UserList(UserIndex).Flags.TargetUser, 0, "||" & UserList(UserIndex).Name & " " & Hechizos(H).TargetMsg & FONTTYPE_FIGHT)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).PropioMsg & FONTTYPE_FIGHT)
        End If
    ElseIf UserList(UserIndex).Flags.TargetNpc > 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||" & Hechizos(H).HechizeroMsg & " " & "la criatura." & FONTTYPE_FIGHT)
    End If
    
End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)

Dim H As Integer
Dim Daño As Integer
Dim tempChr As Integer
Dim defMagica As Integer
    
H = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).Flags.Hechizo)
tempChr = UserList(UserIndex).Flags.TargetUser

'[KEVIN] PARA LOS CLANES**********
If Hechizos(H).SubeHam = 2 Or Hechizos(H).SubeSed = 2 Or Hechizos(H).SubeAgilidad = 2 Or Hechizos(H).SubeFuerza = 2 Or Hechizos(H).SubeHP = 2 Or Hechizos(H).SubeMana = 2 Or Hechizos(H).SubeSta = 2 Then

If UserList(UserIndex).Stats.Matrimonio <> "" Then
    If UserList(UserIndex).Stats.Matrimonio = UserList(tempChr).Name Then
        If UserList(tempChr).Genero = "Hombre" Then
            Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar a tu esposo." & FONTTYPE_INFO)
            Exit Sub
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar a tu esposa." & FONTTYPE_INFO)
            Exit Sub
        End If
    End If
End If

If UserList(UserIndex).GuildInfo.GuildName <> "" Then
    If UserIndex <> tempChr Then
        If UserList(UserIndex).Flags.EnTorneo = 0 Then
            If UserList(UserIndex).GuildInfo.GuildName = UserList(tempChr).GuildInfo.GuildName Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes atacar a tus comprañeros de clan." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If
    End If
End If

End If
'[/KEVIN]********************************

      
'Hambre
If Hechizos(H).SubeHam = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    Daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    Call AddtoVar(UserList(tempChr).Stats.MinHam, _
         Daño, UserList(tempChr).Stats.MaxHam)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    b = True
    
ElseIf Hechizos(H).SubeHam = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    Daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - Daño
    
    If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de hambre a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de hambre." & FONTTYPE_FIGHT)
    End If
    
    Call EnviarHambreYsed(tempChr)
    
    b = True
    
    If UserList(tempChr).Stats.MinHam < 1 Then
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).Flags.Hambre = 1
    End If
    
End If

'Sed
If Hechizos(H).SubeSed = 1 Then
    
    Call InfoHechizo(UserIndex)
    
    Call AddtoVar(UserList(tempChr).Stats.MinAGU, Daño, _
         UserList(tempChr).Stats.MaxAGU)
         
    If UserIndex <> tempChr Then
      Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
      Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
      Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeSed = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - Daño
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de sed a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de sed." & FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).Flags.Sed = 1
    End If
    
    b = True
End If

' <-------- Agilidad ---------->
If Hechizos(H).SubeAgilidad = 1 Then
    
    Call InfoHechizo(UserIndex)
    Daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).Flags.DuracionEfecto = 96
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Agilidad), Daño, MAXATRIBUTOS)
    UserList(tempChr).Flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).Flags.TomoPocion = True
    Daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).Flags.DuracionEfecto = 56
    UserList(tempChr).Stats.UserAtributos(Agilidad) = UserList(tempChr).Stats.UserAtributos(Agilidad) - Daño
    If UserList(tempChr).Stats.UserAtributos(Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Agilidad) = MINATRIBUTOS
    b = True
    
End If

' <-------- Fuerza ---------->
If Hechizos(H).SubeFuerza = 1 Then
    
    Call InfoHechizo(UserIndex)
    Daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    
    UserList(tempChr).Flags.DuracionEfecto = 96
    
    Call AddtoVar(UserList(tempChr).Stats.UserAtributos(Fuerza), Daño, MAXATRIBUTOS)
    UserList(tempChr).Flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeFuerza = 2 Then

    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    UserList(tempChr).Flags.TomoPocion = True
    
    Daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).Flags.DuracionEfecto = 56
    UserList(tempChr).Stats.UserAtributos(Fuerza) = UserList(tempChr).Stats.UserAtributos(Fuerza) - Daño
    If UserList(tempChr).Stats.UserAtributos(Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(Fuerza) = MINATRIBUTOS
    b = True
    
End If

'Salud
If Hechizos(H).SubeHP = 1 Then
    Daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
    
    Call InfoHechizo(UserIndex)
    
    Call AddtoVar(UserList(tempChr).Stats.MinHP, Daño, _
         UserList(tempChr).Stats.MaxHP)
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de vida." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de vida." & FONTTYPE_FIGHT)
    End If
    
    b = True
ElseIf Hechizos(H).SubeHP = 2 Then
    
    If UserIndex = tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes atacarte a vos mismo." & FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    Daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    
    '[KEVIN]*************************************************
    If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).TipoAnillo = 8 Then
                Daño = Daño + ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MaxModificador
                Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
        Else
                If UserList(UserIndex).Invent.Anillo2EqpObjIndex > 0 Then
                     If ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).TipoAnillo = 8 Then
                            Daño = Daño + ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).MaxModificador
                            Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
                     Else
                            Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
                     End If
                 End If
        End If
    ElseIf UserList(UserIndex).Invent.Anillo2EqpObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).TipoAnillo = 8 Then
                Daño = Daño + ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).MaxModificador
                Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
            Else
                Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
            End If
    Else
        Daño = Daño + Porcentaje(Daño, 3 * UserList(UserIndex).Stats.ELV)
    End If
    '[/KEVIN]************************************************************
    
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    '[KEVIN]*********************************
    If UserList(tempChr).Invent.AnilloEqpObjIndex > 0 Then
        If ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).TipoAnillo = 6 Then
        Daño = Daño - ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).MaxModificador
        End If
    End If
    
    If UserList(tempChr).Invent.Anillo2EqpObjIndex > 0 Then
        If ObjData(UserList(tempChr).Invent.Anillo2EqpObjIndex).TipoAnillo = 6 Then
        Daño = Daño - ObjData(UserList(tempChr).Invent.Anillo2EqpObjIndex).MaxModificador
        End If
    End If
    
    If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
        If ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefMagia = 1 Then
            defMagica = RandomNumber(ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).MagiaMinDef, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).MagiaMaxDef)
            Daño = Daño - defMagica
        End If
    End If
    
    If Daño < 1 Then Daño = 1
    '[/KEVIN]****************************************
    
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - Daño
    
    Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de vida a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
    Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de vida." & FONTTYPE_FIGHT)
    
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        Call ContarMuerte(tempChr, UserIndex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, UserIndex)
        Call UserDie(tempChr)
    End If
    
    b = True
End If

'Mana
If Hechizos(H).SubeMana = 1 Then
    
    Call InfoHechizo(UserIndex)
    Call AddtoVar(UserList(tempChr).Stats.MinMAN, Daño, UserList(tempChr).Stats.MaxMAN)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de mana a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - Daño
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    b = True
    
End If

'Stamina
If Hechizos(H).SubeSta = 1 Then
    Call InfoHechizo(UserIndex)
    Call AddtoVar(UserList(tempChr).Stats.MinSta, Daño, _
         UserList(tempChr).Stats.MaxSta)
    If UserIndex <> tempChr Then
         Call SendData(ToIndex, UserIndex, 0, "||Le has restaurado " & Daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
         Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha restaurado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has restaurado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeSta = 2 Then
    If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
    If UserIndex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
    End If
    
    Call InfoHechizo(UserIndex)
    
    If UserIndex <> tempChr Then
        Call SendData(ToIndex, UserIndex, 0, "||Le has quitado " & Daño & " puntos de vitalidad a " & UserList(tempChr).Name & FONTTYPE_FIGHT)
        Call SendData(ToIndex, tempChr, 0, "||" & UserList(UserIndex).Name & " te ha quitado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(ToIndex, UserIndex, 0, "||Te has quitado " & Daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - Daño
    
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
End If


End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(UserIndex, Slot, UserList(UserIndex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(UserIndex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(UserIndex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then

    Call SendData(ToIndex, UserIndex, 0, "SHS" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).Nombre)

Else

    Call SendData(ToIndex, UserIndex, 0, "SHS" & Slot & "," & "0" & "," & "(None)")

End If


End Sub
