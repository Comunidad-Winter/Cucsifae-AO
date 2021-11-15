Attribute VB_Name = "ModFacciones"
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

Public ArmaduraImperial1 As Integer
Public ArmaduraImperial2 As Integer
Public ArmaduraImperial3 As Integer
Public TunicaMagoImperial As Integer
Public TunicaMagoImperialEnanos As Integer


Public ArmaduraCaos1 As Integer
Public TunicaMagoCaos As Integer
Public TunicaMagoCaosEnanos As Integer
Public ArmaduraCaos2 As Integer
Public ArmaduraCaos3 As Integer

'[KEVIN]
Public ArmaduraImpMujer1 As Integer
Public ArmaduraImpMujer2 As Integer
Public ArmaduraImpMujer3 As Integer
Public TunicaMagoImpMujer As Integer
Public TunicaMagoImpEnanosMujer As Integer

Public ArmaduraCaosMujer1 As Integer
Public ArmaduraCaosMujer2 As Integer
Public ArmaduraCaosMujer3 As Integer
Public TunicaMagoCaosMujer As Integer
Public TunicaMagoCaosEnanoMujer As Integer

Public JIIArmaduraImperial1 As Integer
Public JIIArmaduraImperial2 As Integer
Public JIIArmaduraImperial3 As Integer
Public JIITunicaMagoImperial As Integer
Public JIITunicaMagoImperialEnanos As Integer


Public JIIArmaduraCaos1 As Integer
Public JIITunicaMagoCaos As Integer
Public JIITunicaMagoCaosEnanos As Integer
Public JIIArmaduraCaos2 As Integer
Public JIIArmaduraCaos3 As Integer

Public JIIArmaduraImpMujer1 As Integer
Public JIIArmaduraImpMujer2 As Integer
Public JIIArmaduraImpMujer3 As Integer
Public JIITunicaMagoImpMujer As Integer
Public JIITunicaMagoImpEnanosMujer As Integer

Public JIIArmaduraCaosMujer1 As Integer
Public JIIArmaduraCaosMujer2 As Integer
Public JIIArmaduraCaosMujer3 As Integer
Public JIITunicaMagoCaosMujer As Integer
Public JIITunicaMagoCaosEnanoMujer As Integer

Public JIIIArmaduraImperial1 As Integer
Public JIIIArmaduraImperial2 As Integer
Public JIIIArmaduraImperial3 As Integer
Public JIIITunicaMagoImperial As Integer
Public JIIITunicaMagoImperialEnanos As Integer


Public JIIIArmaduraCaos1 As Integer
Public JIIITunicaMagoCaos As Integer
Public JIIITunicaMagoCaosEnanos As Integer
Public JIIIArmaduraCaos2 As Integer
Public JIIIArmaduraCaos3 As Integer

Public JIIIArmaduraImpMujer1 As Integer
Public JIIIArmaduraImpMujer2 As Integer
Public JIIIArmaduraImpMujer3 As Integer
Public JIIITunicaMagoImpMujer As Integer
Public JIIITunicaMagoImpEnanosMujer As Integer

Public JIIIArmaduraCaosMujer1 As Integer
Public JIIIArmaduraCaosMujer2 As Integer
Public JIIIArmaduraCaosMujer3 As Integer
Public JIIITunicaMagoCaosMujer As Integer
Public JIIITunicaMagoCaosEnanoMujer As Integer

Public JIVArmaduraImperial1 As Integer
Public JIVArmaduraImperial2 As Integer
Public JIVArmaduraImperial3 As Integer
Public JIVTunicaMagoImperial As Integer
Public JIVTunicaMagoImperialEnanos As Integer


Public JIVArmaduraCaos1 As Integer
Public JIVTunicaMagoCaos As Integer
Public JIVTunicaMagoCaosEnanos As Integer
Public JIVArmaduraCaos2 As Integer
Public JIVArmaduraCaos3 As Integer

Public JIVArmaduraImpMujer1 As Integer
Public JIVArmaduraImpMujer2 As Integer
Public JIVArmaduraImpMujer3 As Integer
Public JIVTunicaMagoImpMujer As Integer
Public JIVTunicaMagoImpEnanosMujer As Integer

Public JIVArmaduraCaosMujer1 As Integer
Public JIVArmaduraCaosMujer2 As Integer
Public JIVArmaduraCaosMujer3 As Integer
Public JIVTunicaMagoCaosMujer As Integer
Public JIVTunicaMagoCaosEnanoMujer As Integer
'[/KEVIN]

Public Const ExpAlUnirse = 100000
Public Const ExpX100 = 10000

'[KEVIN]
Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.ArmadaReal > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya perteneces a las tropas reales!!! Ve a combatir criminales!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Maldito insolente!!! vete de aqui seguidor de las sombras!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No se permiten criminales en el ejercito imperial!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.CriminalesMatados < 10 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos 10 criminales, solo has matado " & UserList(UserIndex).Faccion.CriminalesMatados & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 18 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 18!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.CiudadanosMatados > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

Select Case UserList(UserIndex).Raza
    Case "Humano"
        UserList(UserIndex).Faccion.ArmadaReal = 1
    Case "Elfo"
        UserList(UserIndex).Faccion.ArmadaReal = 2
    Case "Elfo Oscuro"
        UserList(UserIndex).Faccion.ArmadaReal = 3
    Case "Gnomo"
        UserList(UserIndex).Faccion.ArmadaReal = 4
    Case "Enano"
        UserList(UserIndex).Faccion.ArmadaReal = 5
    Case "Orco"
        UserList(UserIndex).Faccion.ArmadaReal = 6
End Select

UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.CriminalesMatados \ 100

Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Bienvenido a al Ejercito Imperial!!!, aqui tienes tu armadura. Por cada centena de criminales que acabes te dare un recompensa, buena suerte soldado!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))

If UserList(UserIndex).Faccion.RecibioArmaduraReal = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
    
'[/KEVIN] MODIFICADO
    If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = TunicaMagoImperialEnanos
               Else
                      MiObj.ObjIndex = TunicaMagoImperial
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = ArmaduraImperial3
                  Else
                      MiObj.ObjIndex = ArmaduraImperial1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = ArmaduraImperial3
                  Else
                      MiObj.ObjIndex = ArmaduraImperial2
                  End If
        End If
    Else
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = TunicaMagoImpEnanosMujer
               Else
                      MiObj.ObjIndex = TunicaMagoImpMujer
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = ArmaduraImpMujer3
                  Else
                      MiObj.ObjIndex = ArmaduraImpMujer1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                     MiObj.ObjIndex = ArmaduraImpMujer3
                  Else
                      MiObj.ObjIndex = ArmaduraImpMujer2
                  End If
        End If
    End If
    
    '[/KEVIN]
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 1
End If

If UserList(UserIndex).Faccion.RecibioExpInicialReal = 0 Then
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpAlUnirse, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecibioExpInicialReal = 1
    Call CheckUserLevel(UserIndex)
End If

Select Case UserList(UserIndex).Raza
    Case "Humano"
        Call LogEjercitoReal(UserList(UserIndex).Name)
    Case "Elfo"
        Call LogEjercitoRealElfico(UserList(UserIndex).Name)
    Case "Elfo Oscuro"
        Call LogEjercitoRealElficoOscuro(UserList(UserIndex).Name)
    Case "Gnomo"
        Call LogEjercitoRealGnomico(UserList(UserIndex).Name)
    Case "Enano"
        Call LogEjercitoRealEnano(UserList(UserIndex).Name)
    Case "Orco"
        Call LogEjercitoRealOrco(UserList(UserIndex).Name)
End Select


'[/KEVIN]
End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.CriminalesMatados \ 100 = _
   UserList(UserIndex).Faccion.RecompensasReal Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, mata 100 criminales mas para recibir la proxima!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
Else
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpX100, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.RecompensasReal + 1
    Call CheckUserLevel(UserIndex)
    '[KEVIN]
    If UserList(UserIndex).Faccion.RecompensasReal = 3 Then
            Call DarArmaduraJII(UserIndex)
    ElseIf UserList(UserIndex).Faccion.RecompensasReal = 7 Then
            Call DarArmaduraJIII(UserIndex)
    ElseIf UserList(UserIndex).Faccion.RecompensasReal = 10 Then
            Call DarArmaduraJIV(UserIndex)
    End If
    '[/KEVIN]
End If

End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)
On Error GoTo errhand

UserList(UserIndex).Faccion.ArmadaReal = 0

Dim LoopC As Integer

For LoopC = 1 To MAX_INVENTORY_SLOTS
    If ObjData(UserList(UserIndex).Invent.Object(LoopC).ObjIndex).Real = 1 Then
        If UserList(UserIndex).Invent.Object(LoopC).Equipped = 1 Then
            Call Desequipar(UserIndex, LoopC)
        End If
    End If
Next LoopC

Call SendData(ToIndex, UserIndex, 0, "||Has sido expulsado de las tropas reales!!!." & FONTTYPE_FIGHT)

Exit Sub

errhand:
Call LogError("Error en ExpulsarFaccionReal, " & UserList(UserIndex).Name)

End Sub
'[KEVIN]
Public Function TituloReal(ByVal UserIndex As Integer) As String

Select Case UserList(UserIndex).Faccion.RecompensasReal
    Case 0
        TituloReal = "Aprendiz real"
    Case 1
        TituloReal = "Iniciado"
    Case 2
        TituloReal = "Escudero"
    Case 3
        TituloReal = "Soldado Real"
    Case 4
        TituloReal = "Teniente Real"
    Case 5
        TituloReal = "General real"
    Case 6
        TituloReal = "Guardian del bien"
    Case 7
        TituloReal = "Protector del Bien"
    Case 8
        TituloReal = "Emisario del Bien"
    Case 9
        TituloReal = "Caballero Real"
    Case Else
        TituloReal = "Elite Real"
End Select

End Function
'[/KEVIN]

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
'[KEVIN]
If Not Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Largate de aqui, bufon!!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya perteneces a las tropas del caos!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.ArmadaReal > 0 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Las sombras reinaran en Argentum, largate de aqui estupido ciudadano.!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If Not Criminal(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ja ja ja tu no eres bienvenido aqui!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Faccion.CiudadanosMatados < 50 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes matar al menos 50 ciudadanos, solo has matado " & UserList(UserIndex).Faccion.CiudadanosMatados & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Exit Sub
End If

Select Case UserList(UserIndex).Raza
    Case "Humano"
        UserList(UserIndex).Faccion.FuerzasCaos = 1
    Case "Elfo"
        UserList(UserIndex).Faccion.FuerzasCaos = 2
    Case "Elfo Oscuro"
        UserList(UserIndex).Faccion.FuerzasCaos = 3
    Case "Gnomo"
        UserList(UserIndex).Faccion.FuerzasCaos = 4
    Case "Enano"
        UserList(UserIndex).Faccion.FuerzasCaos = 5
    Case "Orco"
        UserList(UserIndex).Faccion.FuerzasCaos = 6
End Select

UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.CiudadanosMatados \ 100

Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Bienvenido a al lado oscuro!!!, aqui tienes tu armadura. Por cada centena de ciudadanos que acabes te dare un recompensa, buena suerte soldado!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))

If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0 Then
    Dim MiObj As Obj
    MiObj.Amount = 1
    
    '[KEVIN] MODIFICADO
    If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
            If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                    MiObj.ObjIndex = TunicaMagoCaosEnanos
            Else
                    MiObj.ObjIndex = TunicaMagoCaos
            End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = ArmaduraCaos3
                  Else
                      MiObj.ObjIndex = ArmaduraCaos1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = ArmaduraCaos3
                  Else
                      MiObj.ObjIndex = ArmaduraCaos2
                  End If
        End If
    Else
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
            If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = TunicaMagoCaosEnanoMujer
            Else
                      MiObj.ObjIndex = TunicaMagoCaosMujer
            End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = ArmaduraCaosMujer3
                  Else
                      MiObj.ObjIndex = ArmaduraCaosMujer1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = ArmaduraCaosMujer3
                  Else
                      MiObj.ObjIndex = ArmaduraCaosMujer2
                  End If
        End If
    End If
    
    '[/KEVIN]
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 1
End If

If UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0 Then
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpAlUnirse, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecibioExpInicialCaos = 1
    Call CheckUserLevel(UserIndex)
End If

Select Case UserList(UserIndex).Raza
    Case "Humano"
        Call LogEjercitoCaos(UserList(UserIndex).Name)
    Case "Elfo"
        Call LogEjercitoCaosElfico(UserList(UserIndex).Name)
    Case "Elfo Oscuro"
        Call LogEjercitoCaosElficoOscuro(UserList(UserIndex).Name)
    Case "Gnomo"
        Call LogEjercitoCaosGnomico(UserList(UserIndex).Name)
    Case "Enano"
        Call LogEjercitoCaosEnano(UserList(UserIndex).Name)
    Case "Orco"
        Call LogEjercitoCaosOrco(UserList(UserIndex).Name)
End Select
'[/KEVIN]
End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)

If UserList(UserIndex).Faccion.CiudadanosMatados \ 100 = _
   UserList(UserIndex).Faccion.RecompensasCaos Then
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Ya has recibido tu recompensa, mata 100 ciudadanos mas para recibir la proxima!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
Else
    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Aqui tienes tu recompensa noble guerrero!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
    Call AddtoVar(UserList(UserIndex).Stats.Exp, ExpX100, MAXEXP)
    Call SendData(ToIndex, UserIndex, 0, "||Has ganado " & ExpX100 & " puntos de experiencia." & FONTTYPE_FIGHT)
    UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.RecompensasCaos + 1
    Call CheckUserLevel(UserIndex)
    '[KEVIN]
    If UserList(UserIndex).Faccion.RecompensasCaos = 3 Then
            Call DarArmaduraJIIC(UserIndex)
    ElseIf UserList(UserIndex).Faccion.RecompensasCaos = 7 Then
            Call DarArmaduraJIIIC(UserIndex)
    ElseIf UserList(UserIndex).Faccion.RecompensasCaos = 10 Then
            Call DarArmaduraJIVC(UserIndex)
    End If
    '[/KEVIN]
End If


End Sub

Public Sub ExpulsarCaos(ByVal UserIndex As Integer)

UserList(UserIndex).Faccion.FuerzasCaos = 0

'[KEVIN]
Dim LoopC As Integer

For LoopC = 1 To MAX_INVENTORY_SLOTS
    If ObjData(UserList(UserIndex).Invent.Object(LoopC).ObjIndex).Caos = 1 Then
        If UserList(UserIndex).Invent.Object(LoopC).Equipped = 1 Then
            Call Desequipar(UserIndex, LoopC)
        End If
    End If
Next LoopC
'[/KEVIN]

Call SendData(ToIndex, UserIndex, 0, "||Has sido expulsado del ejercito del caos!!!." & FONTTYPE_FIGHT)
End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
Select Case UserList(UserIndex).Faccion.RecompensasCaos
    Case 0
        TituloCaos = "Esclavo de las sombras"
    Case 1
        TituloCaos = "Guerrero del caos"
    Case 2
        TituloCaos = "Teniente del caos"
    Case 3
        TituloCaos = "Comandante del caos"
    Case 4
        TituloCaos = "General del caos"
    Case 5
        TituloCaos = "Elite del caos"
    Case 6
        TituloCaos = "Asolador de las sombras"
    Case 7
        TituloCaos = "Caballero Oscuro"
    Case 8
        TituloCaos = "Asesino del caos"
    Case Else
        TituloCaos = "Adorador del demonio"
End Select


End Function
'[KEVIN]
Public Sub DarArmaduraJII(ByVal UserIndex As Integer)

Dim MiObj As Obj

If UserList(UserIndex).Faccion.RecibioArmaduraReal = 1 Then
    MiObj.Amount = 1
    
    If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIITunicaMagoImperialEnanos
               Else
                      MiObj.ObjIndex = JIITunicaMagoImperial
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIArmaduraImperial3
                  Else
                      MiObj.ObjIndex = JIIArmaduraImperial1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIArmaduraImperial3
                  Else
                      MiObj.ObjIndex = JIIArmaduraImperial2
                  End If
        End If
    Else
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIITunicaMagoImpEnanosMujer
               Else
                      MiObj.ObjIndex = JIITunicaMagoImpMujer
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIArmaduraImpMujer3
                  Else
                      MiObj.ObjIndex = JIIArmaduraImpMujer1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                     MiObj.ObjIndex = JIIArmaduraImpMujer3
                  Else
                      MiObj.ObjIndex = JIIArmaduraImpMujer2
                  End If
        End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 2

End If
'[/KEVIN]
End Sub
Public Sub DarArmaduraJIIC(ByVal UserIndex As Integer)

Dim MiObj As Obj

If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 1 Then
    MiObj.Amount = 1
    
    If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIITunicaMagoCaosEnanos
               Else
                      MiObj.ObjIndex = JIITunicaMagoCaos
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIArmaduraCaos3
                  Else
                      MiObj.ObjIndex = JIIArmaduraCaos1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIArmaduraCaos3
                  Else
                      MiObj.ObjIndex = JIIArmaduraCaos2
                  End If
        End If
    Else
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIITunicaMagoCaosEnanoMujer
               Else
                      MiObj.ObjIndex = JIITunicaMagoCaosMujer
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIArmaduraCaosMujer3
                  Else
                      MiObj.ObjIndex = JIIArmaduraCaosMujer1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                     MiObj.ObjIndex = JIIArmaduraCaosMujer3
                  Else
                      MiObj.ObjIndex = JIIArmaduraCaosMujer2
                  End If
        End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 2
End If
    
End Sub

'[KEVIN]
Public Sub DarArmaduraJIII(ByVal UserIndex As Integer)

Dim MiObj As Obj

If UserList(UserIndex).Faccion.RecibioArmaduraReal = 2 Then
    MiObj.Amount = 1
    
    If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIITunicaMagoImperialEnanos
               Else
                      MiObj.ObjIndex = JIIITunicaMagoImperial
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIIArmaduraImperial3
                  Else
                      MiObj.ObjIndex = JIIIArmaduraImperial1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIIArmaduraImperial3
                  Else
                      MiObj.ObjIndex = JIIIArmaduraImperial2
                  End If
        End If
    Else
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIITunicaMagoImpEnanosMujer
               Else
                      MiObj.ObjIndex = JIIITunicaMagoImpMujer
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIIArmaduraImpMujer3
                  Else
                      MiObj.ObjIndex = JIIIArmaduraImpMujer1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                     MiObj.ObjIndex = JIIIArmaduraImpMujer3
                  Else
                      MiObj.ObjIndex = JIIIArmaduraImpMujer2
                  End If
        End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 3
End If
'[/KEVIN]
End Sub
'[KEVIN]
Public Sub DarArmaduraJIIIC(ByVal UserIndex As Integer)

Dim MiObj As Obj

If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 2 Then
    MiObj.Amount = 1
    
    If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIITunicaMagoCaosEnanos
               Else
                      MiObj.ObjIndex = JIIITunicaMagoCaos
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIIArmaduraCaos3
                  Else
                      MiObj.ObjIndex = JIIIArmaduraCaos1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIIArmaduraCaos3
                  Else
                      MiObj.ObjIndex = JIIIArmaduraCaos2
                  End If
        End If
    Else
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIITunicaMagoCaosEnanoMujer
               Else
                      MiObj.ObjIndex = JIIITunicaMagoCaosMujer
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIIIArmaduraCaosMujer3
                  Else
                      MiObj.ObjIndex = JIIIArmaduraCaosMujer1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                     MiObj.ObjIndex = JIIIArmaduraCaosMujer3
                  Else
                      MiObj.ObjIndex = JIIIArmaduraCaosMujer2
                  End If
        End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 3
End If

End Sub
'[KEVIN]
Public Sub DarArmaduraJIV(ByVal UserIndex As Integer)

Dim MiObj As Obj

If UserList(UserIndex).Faccion.RecibioArmaduraReal = 3 Then
    MiObj.Amount = 1
    
    If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIVTunicaMagoImperialEnanos
               Else
                      MiObj.ObjIndex = JIVTunicaMagoImperial
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIVArmaduraImperial3
                  Else
                      MiObj.ObjIndex = JIVArmaduraImperial1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIVArmaduraImperial3
                  Else
                      MiObj.ObjIndex = JIVArmaduraImperial2
                  End If
        End If
    Else
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIVTunicaMagoImpEnanosMujer
               Else
                      MiObj.ObjIndex = JIVTunicaMagoImpMujer
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIVArmaduraImpMujer3
                  Else
                      MiObj.ObjIndex = JIVArmaduraImpMujer1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                     MiObj.ObjIndex = JIVArmaduraImpMujer3
                  Else
                      MiObj.ObjIndex = JIVArmaduraImpMujer2
                  End If
        End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 4
End If
'[/KEVIN]
End Sub

'[KEVIN]
Public Sub DarArmaduraJIVC(ByVal UserIndex As Integer)

Dim MiObj As Obj

If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 3 Then
    MiObj.Amount = 1
    
    If UCase$(UserList(UserIndex).Genero) = "HOMBRE" Then
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIVTunicaMagoCaosEnanos
               Else
                      MiObj.ObjIndex = JIVTunicaMagoCaos
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIVArmaduraCaos3
                  Else
                      MiObj.ObjIndex = JIVArmaduraCaos1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIVArmaduraCaos3
                  Else
                      MiObj.ObjIndex = JIVArmaduraCaos2
                  End If
        End If
    Else
        If UCase(UserList(UserIndex).Clase) = "MAGO" Then
               If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                  UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIVTunicaMagoCaosEnanoMujer
               Else
                      MiObj.ObjIndex = JIVTunicaMagoCaosMujer
               End If
        ElseIf UCase(UserList(UserIndex).Clase) = "GUERRERO" Or _
               UCase(UserList(UserIndex).Clase) = "CAZADOR" Or _
               UCase(UserList(UserIndex).Clase) = "PALADIN" Or _
               UCase(UserList(UserIndex).Clase) = "BANDIDO" Or _
               UCase(UserList(UserIndex).Clase) = "ASESINO" Then
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                      MiObj.ObjIndex = JIVArmaduraCaosMujer3
                  Else
                      MiObj.ObjIndex = JIVArmaduraCaosMujer1
                  End If
        Else
                  If UCase(UserList(UserIndex).Raza) = "ENANO" Or _
                     UCase(UserList(UserIndex).Raza) = "GNOMO" Then
                     MiObj.ObjIndex = JIVArmaduraCaosMujer3
                  Else
                      MiObj.ObjIndex = JIVArmaduraCaosMujer2
                  End If
        End If
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If

    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 4
End If

End Sub
