Attribute VB_Name = "mdlComandosGMS"
Option Explicit

Sub ComandosGMS(ByVal Userindex As Integer, ByVal rdata As String)

On Error GoTo errorhandler:




Dim sndData As String
Dim CadenaOriginal As String

Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim auxind As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim Name As String
Dim ind
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim cliMD5 As String

'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
 If UserList(Userindex).flags.Privilegios = 0 Then Exit Sub
'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<


'<<<<<<<<<<<<<<<<<<<< Consejeros <<<<<<<<<<<<<<<<<<<<

'/rem comentario
If UCase$(Left$(rdata, 4)) = "/REM" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    Call LogGM(UserList(Userindex).Name, "Comentario: " & rdata, (UserList(Userindex).flags.Privilegios = 1))
    Call SendData(ToIndex, Userindex, 0, "||Comentario salvado..." & FONTTYPE_INFO)
    Exit Sub
End If

'HORA
If UCase$(Left$(rdata, 5)) = "/HORA" Then
    Call LogGM(UserList(Userindex).Name, "Hora.", (UserList(Userindex).flags.Privilegios = 1))
    rdata = Right$(rdata, Len(rdata) - 5)
    Call SendData(ToAll, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_INFO)
    Exit Sub
End If

'¿Donde esta?
If UCase$(Left$(rdata, 7)) = "/DONDE " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    Call SendData(ToIndex, Userindex, 0, "||Ubicacion  " & UserList(tIndex).Name & ": " & UserList(tIndex).Pos.Map & ", " & UserList(tIndex).Pos.X & ", " & UserList(tIndex).Pos.Y & "." & FONTTYPE_INFO)
    Call LogGM(UserList(Userindex).Name, "/Donde", (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If

'Nro de enemigos
If UCase$(Left$(rdata, 6)) = "/NENE " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    If MapaValido(val(rdata)) Then
        Call SendData(ToIndex, Userindex, 0, "NENE" & NPCHostiles(val(rdata)))
        Call LogGM(UserList(Userindex).Name, "Numero enemigos en mapa " & rdata, (UserList(Userindex).flags.Privilegios = 1))
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/SHOW SOS" Then
    Dim M As String
    For N = 1 To Ayuda.longitud
        M = Ayuda.VerElemento(N)
        Call SendData(ToIndex, Userindex, 0, "RSOS" & M)
    Next N
    Call SendData(ToIndex, Userindex, 0, "MSOS")
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "SOSDONE" Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Call Ayuda.Quitar(rdata, vbTab)
    Exit Sub
End If

'Teleportar
If UCase$(Left$(rdata, 7)) = "/TELEP " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    mapa = val(ReadField(2, rdata, 32))
    If Not MapaValido(mapa) Then Exit Sub
    Name = ReadField(1, rdata, 32)
    If Name = "" Then Exit Sub
    If UCase$(Name) <> "YO" Then
        If UserList(Userindex).flags.Privilegios = 1 Then
            Exit Sub
        End If
        tIndex = NameIndex(Name)
    Else
        tIndex = Userindex
    End If
    X = val(ReadField(3, rdata, 32))
    Y = val(ReadField(4, rdata, 32))
    If Not InMapBounds(mapa, X, Y) Then Exit Sub
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    '[KEVIN]
    UserList(tIndex).flags.EnTorneo = 0
    '[/KEVIN]
    Call WarpUserChar(tIndex, mapa, X, Y, True)
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " transportado." & FONTTYPE_INFO)
    Call LogGM(UserList(Userindex).Name, "Transporto a " & UserList(tIndex).Name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If

'IR A
If UCase$(Left$(rdata, 5)) = "/IRA " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    

    Call WarpUserChar(Userindex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y + 1, True)
    
    If UserList(Userindex).flags.AdminInvisible = 0 Then Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " se ha trasportado hacia donde te encontras." & FONTTYPE_INFO)
    Call LogGM(UserList(Userindex).Name, "/IRA " & UserList(tIndex).Name & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y, (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(rdata) = "/TELEPLOC" Then
    Call WarpUserChar(Userindex, UserList(Userindex).flags.TargetMap, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY, True)
    Call LogGM(UserList(Userindex).Name, "/TELEPLOC a x:" & UserList(Userindex).flags.TargetX & " Y:" & UserList(Userindex).flags.TargetY & " Map:" & UserList(Userindex).Pos.Map, (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If

'Haceme invisible vieja!
If UCase$(rdata) = "/INVISIBLE" Then
    Call DoAdminInvisible(Userindex)
    Call LogGM(UserList(Userindex).Name, "/INVISIBLE", (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If

'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
If UserList(Userindex).flags.Privilegios < 2 Then
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/REVIVIR " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    Name = rdata
    If UCase$(Name) <> "YO" Then
        tIndex = NameIndex(Name)
    Else
        tIndex = Userindex
    End If
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    UserList(tIndex).flags.Muerto = 0
    UserList(tIndex).Stats.MinHP = UserList(tIndex).Stats.MaxHP
    Call DarCuerpoDesnudo(tIndex)
    Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, val(tIndex), UserList(tIndex).Char.Body, UserList(tIndex).OrigChar.Head, UserList(tIndex).Char.Heading, UserList(tIndex).Char.WeaponAnim, UserList(tIndex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim, UserList(Userindex).flags.AdminInvisible, 0)
    Call SendUserStatsBox(val(tIndex))
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " te há resucitado." & FONTTYPE_INFO)
    Call LogGM(UserList(Userindex).Name, "Resucito a " & UserList(tIndex).Name, (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If

'INFO DE USER
If UCase$(Left$(rdata, 6)) = "/INFO " Then
    Call LogGM(UserList(Userindex).Name, rdata, (UserList(Userindex).flags.Privilegios = 1))
    
    rdata = Right$(rdata, Len(rdata) - 6)
    
    tIndex = NameIndex(rdata)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If

    SendUserStatsTxt Userindex, tIndex
    Exit Sub
End If

'INV DEL USER
If UCase$(Left$(rdata, 5)) = "/INV " Then
    Call LogGM(UserList(Userindex).Name, rdata, (UserList(Userindex).flags.Privilegios = 1))
    
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = NameIndex(rdata)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If

    SendUserInvTxt Userindex, tIndex
    Exit Sub
End If

'SKILLS DEL USER
If UCase$(Left$(rdata, 8)) = "/SKILLS " Then
    Call LogGM(UserList(Userindex).Name, rdata, (UserList(Userindex).flags.Privilegios = 1))
    
    rdata = Right$(rdata, Len(rdata) - 8)
    
    tIndex = NameIndex(rdata)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If

    SendUserSkillsTxt Userindex, tIndex
    Exit Sub
End If

If UCase$(rdata) = "/ONLINEGM" Then
        tStr = ""
        For LoopC = 1 To LastUser
            If (UserList(LoopC).Name <> "") And UserList(LoopC).flags.Privilegios <> 0 Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(ToIndex, Userindex, 0, "||" & tStr & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, Userindex, 0, "||No hay GMs Online" & FONTTYPE_INFO)
        End If
        Exit Sub
End If

'Bloquear
If UCase$(Left$(rdata, 5)) = "/BLOQ" Then
    Call LogGM(UserList(Userindex).Name, "/BLOQ", (UserList(Userindex).flags.Privilegios = 1))
    rdata = Right$(rdata, Len(rdata) - 5)
    If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Blocked = 0 Then
        MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Blocked = 1
        Call Bloquear(ToMap, Userindex, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, 1)
    Else
        MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Blocked = 0
        Call Bloquear(ToMap, Userindex, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, 0)
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/CARCEL " Then
    
    Dim i As Integer
    
    rdata = Right$(rdata, Len(rdata) - 8)
    
    Name = ReadField(1, rdata, 32)
    i = val(ReadField(1, rdata, 32))
    Name = Right$(rdata, Len(rdata) - (Len(Name) + 1))
    
    tIndex = NameIndex(Name)
    
'    If ucase$(Name) = "MORGOLOCK" Then Exit Sub
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > UserList(Userindex).flags.Privilegios Then
        Call SendData(ToIndex, Userindex, 0, "||No podes encarcelar a alguien con jerarquia mayor a la tuya." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If i < 1 Then Exit Sub
    
    If i > 30 Then
        Call SendData(ToIndex, Userindex, 0, "||No podes encarcelar por mas de 30 minutos." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    UserList(tIndex).flags.EnTorneo = 0
    Call Encarcelar(tIndex, i, UserList(Userindex).Name)
    
    Exit Sub
End If

'[KEVIN]
If UCase(Left$(rdata, 8)) = "/VERINSC" Then
    Dim mm As String
    For N = 1 To InscTorneo.longitud
        mm = InscTorneo.VerElemento(N)
        Call SendData(ToIndex, Userindex, 0, "IRSOS" & mm)
    Next N
    Call SendData(ToIndex, Userindex, 0, "IMSOS")
    Exit Sub
End If

If UCase(Left$(rdata, 8)) = "ISOSDONE" Then
    rdata = Right$(rdata, Len(rdata) - 8)
    Call InscTorneo.Quitar(rdata)
    Exit Sub
End If

If UCase(Left$(rdata, 7)) = "/ABRIR " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    
    If Not Numeric(rdata) Then
        Call SendData(ToIndex, Userindex, 0, "||Debes escribir el nivel mínimo de los pjs a entrar" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    MinLevel = rdata
        
    If MinLevel < 13 Then
        Call SendData(ToIndex, Userindex, 0, "||Los newbies no pueden participar de los torneos" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    InscAbierta = True
    Call SendData(ToIndex, Userindex, 0, "||Has abierto la inscripción." & FONTTYPE_INFO)
    Exit Sub
End If

If UCase(Left$(rdata, 7)) = "/CERRAR" Then
    rdata = Right$(rdata, Len(rdata) - 7)
    InscAbierta = False
    MinLevel = 1
    Call SendData(ToIndex, Userindex, 0, "||Has cerrado la inscripción." & FONTTYPE_INFO)
    Exit Sub
End If
'[/KEVIN]

'PERDON
If UCase$(Left$(rdata, 7)) = "/PERDON" Then
    rdata = Right$(rdata, Len(rdata) - 8)
    tIndex = NameIndex(rdata)
    If tIndex > 0 Then
        
        If EsNewbie(tIndex) Then
                Call VolverCiudadano(tIndex)
        Else
                Call LogGM(UserList(Userindex).Name, "Intento perdonar un personaje de nivel avanzado.", (UserList(Userindex).flags.Privilegios = 1))
                Call SendData(ToIndex, Userindex, 0, "||Solo se permite perdonar newbies." & FONTTYPE_INFO)
        End If
        
    End If
    Exit Sub
End If

'[KEVIN]
If UCase$(Left$(rdata, 9)) = "/PERDONC " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    tIndex = NameIndex(rdata)
    If tIndex > 0 Then
    
    Call VolverCiudadano(tIndex)

    End If
    Exit Sub
End If
'[/KEVIN]

'Echar usuario
If UCase$(Left$(rdata, 7)) = "/ECHAR " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(rdata)
    If UCase$(rdata) = "NEB" Then Exit Sub
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > UserList(Userindex).flags.Privilegios Then
        Call SendData(ToIndex, Userindex, 0, "||No podes echar a alguien con jerarquia mayor a la tuya." & FONTTYPE_INFO)
        Exit Sub
    End If
        
    Call SendData(ToAll, 0, 0, "||" & UserList(Userindex).Name & " echo a " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
    Call CloseSocket(tIndex)
    Call LogGM(UserList(Userindex).Name, "Echo a " & UserList(tIndex).Name, (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If


If UCase$(Left$(rdata, 5)) = "/BAN " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    tIndex = NameIndex(ReadField(2, rdata, Asc("@")))
    Name = ReadField(1, rdata, Asc("@"))
    
    If UCase$(rdata) = "NEB" Then Exit Sub
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).flags.Privilegios > UserList(Userindex).flags.Privilegios Then
        Call SendData(ToIndex, Userindex, 0, "||No podes banear a al alguien de mayor jerarquia." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call LogBan(tIndex, Userindex, Name)
    
    Call SendData(ToAdmins, 0, 0, "||" & UserList(Userindex).Name & " echo a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
    Call SendData(ToAdmins, 0, 0, "||" & UserList(Userindex).Name & " Banned a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
    
    'Ponemos el flag de ban a 1
    UserList(tIndex).flags.Ban = 1
    
    If UserList(tIndex).flags.Privilegios > 0 Then
            UserList(Userindex).flags.Ban = 1
            Call CloseSocket(Userindex)
            Call SendData(ToAdmins, 0, 0, "||" & UserList(Userindex).Name & " banned by the server por bannear un Administrador." & FONTTYPE_FIGHT)
    End If
    
    Call LogGM(UserList(Userindex).Name, "Echo a " & UserList(tIndex).Name, (UserList(Userindex).flags.Privilegios = 1))
    Call LogGM(UserList(Userindex).Name, "BAN a " & UserList(tIndex).Name, (UserList(Userindex).flags.Privilegios = 1))
    Call CloseSocket(tIndex)
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/UNBAN " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Call UnBan(rdata)
    Call LogGM(UserList(Userindex).Name, "/UNBAN a " & rdata, (UserList(Userindex).flags.Privilegios = 1))
    Call SendData(ToIndex, Userindex, 0, "||" & rdata & " unbanned." & FONTTYPE_INFO)
    Exit Sub
End If

'[KEVIN]
'Ver logs de sh en la consola
If UCase$(Left$(rdata, 6)) = "/VERSH" Then
    rdata = Right$(rdata, Len(rdata) - 6)

    If UserList(Userindex).flags.VerSH = False Then
        UserList(Userindex).flags.VerSH = True
        Call SendData(ToIndex, Userindex, 0, "||VerSh ON." & FONTTYPE_CELESTE)
    Else
        UserList(Userindex).flags.VerSH = False
        Call SendData(ToIndex, Userindex, 0, "||VerSh OFF." & FONTTYPE_CELESTE)
    End If

    Exit Sub
End If
'[/KEVIN]

'SEGUIR
If UCase$(rdata) = "/SEGUIR" Then
    If UserList(Userindex).flags.TargetNpc > 0 Then
        Call DoFollow(UserList(Userindex).flags.TargetNpc, UserList(Userindex).Name)
    End If
    Exit Sub
End If

'Summon
If UCase$(Left$(rdata, 5)) = "/SUM " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " há sido trasportado." & FONTTYPE_INFO)
    Call WarpUserChar(tIndex, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y + 1, True)
    
    Call LogGM(UserList(Userindex).Name, "/SUM " & UserList(tIndex).Name & " Map:" & UserList(Userindex).Pos.Map & " X:" & UserList(Userindex).Pos.X & " Y:" & UserList(Userindex).Pos.Y, (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If

'[KEVIN]
'Summon pa Torneo
If UCase$(Left$(rdata, 8)) = "/SUMTNO " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " há sido trasportado." & FONTTYPE_INFO)
    Call WarpUserChar(tIndex, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y + 1, True)
    
    UserList(tIndex).flags.EnTorneo = 1

    Call LogGM(UserList(Userindex).Name, "/SUMTNO " & UserList(Userindex).Name & " Map:" & UserList(Userindex).Pos.Map & " X:" & UserList(Userindex).Pos.X & " Y:" & UserList(Userindex).Pos.Y, (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If

'Pone el flag de torneo en 0 de todos los usuarios
If UCase$(Left$(rdata, 7)) = "/RFLAGS" Then
    rdata = Right$(rdata, Len(rdata) - 7)
    
    Dim LoopU As Integer
    For LoopU = 1 To LastUser
        If UserList(LoopU).flags.EnTorneo = 1 Then UserList(LoopU).flags.EnTorneo = 0
    Next LoopU
    
    MinLevel = 1
    
    Exit Sub
End If
'[/KEVIN]

'Crear criatura
If UCase$(Left$(rdata, 3)) = "/CC" Then
   Call EnviarSpawnList(Userindex)
   Exit Sub
End If

'Spawn!!!!!
If UCase$(Left$(rdata, 3)) = "SPA" Then
    rdata = Right$(rdata, Len(rdata) - 3)
    
    If val(rdata) > 0 And val(rdata) < UBound(SpawnList) + 1 Then _
          Call SpawnNpc(SpawnList(val(rdata)).NpcIndex, UserList(Userindex).Pos, True, False)
          
          Call LogGM(UserList(Userindex).Name, "Sumoneo " & SpawnList(val(rdata)).NpcName, (UserList(Userindex).flags.Privilegios = 1))
          
    Exit Sub
End If

'Resetea el inventario
If UCase$(rdata) = "/RESETINV" Then
    rdata = Right$(rdata, Len(rdata) - 9)
    If UserList(Userindex).flags.TargetNpc = 0 Then Exit Sub
    Call ResetNpcInv(UserList(Userindex).flags.TargetNpc)
    Call LogGM(UserList(Userindex).Name, "/RESETINV " & Npclist(UserList(Userindex).flags.TargetNpc).Name, (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If

'/Clean
If UCase$(rdata) = "/LIMPIAR" Then
    Call LimpiarMundo
    Exit Sub
End If

'Mensaje del servidor
If UCase$(Left$(rdata, 6)) = "/RMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(Userindex).Name, "Mensaje Broadcast:" & rdata, (UserList(Userindex).flags.Privilegios = 1))
    If rdata <> "" Then
        Call SendData(ToAll, 0, 0, "||" & UserList(Userindex).Name & ": " & rdata & FONTTYPE_TALK & ENDC)
    End If
    Exit Sub
End If

'[KEVIN]
If UCase$(Left$(rdata, 9)) = "/SERVMSG " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    Call LogGM(UserList(Userindex).Name, "Mensaje al Servidor:" & rdata, (UserList(Userindex).flags.Privilegios = 1))
    If rdata <> "" Then
        frmMain.rtfGmMsg.SelText = UserList(Userindex).Name & ": " & rdata & vbCrLf
        Call SendData(ToIndex, Userindex, 0, "||Has enviado un mensaje al server" & FONTTYPE_INFO)
    End If
    Exit Sub
End If
'[/KEVIN]

'Ip del nick
If UCase$(Left$(rdata, 8)) = "/IPNICK " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    tIndex = NameIndex(UCase$(rdata))
    If tIndex > 0 Then
       Call SendData(ToIndex, Userindex, 0, "||El ip de " & rdata & " es " & UserList(tIndex).ip & FONTTYPE_INFO)
    End If
    Exit Sub
End If

'Ip del nick
If UCase$(Left$(rdata, 8)) = "/NICKIP " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    tIndex = IP_Index(rdata)
    If tIndex > 0 Then
       Call SendData(ToIndex, Userindex, 0, "||El nick del ip " & rdata & " es " & UserList(tIndex).Name & FONTTYPE_INFO)
    End If
    Exit Sub
End If

'[KEVIN]
If UCase$(Left$(rdata, 8)) = "$USKILL " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    
    tIndex = NameIndex(UCase(rdata))
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(ToIndex, Userindex, 0, "PNSKILL" & UserList(tIndex).Stats.UserSkills(1) & "," & UserList(tIndex).Stats.UserSkills(2) & "," & UserList(tIndex).Stats.UserSkills(3) & "," & UserList(tIndex).Stats.UserSkills(4) & "," & UserList(tIndex).Stats.UserSkills(5) & "," & UserList(tIndex).Stats.UserSkills(6) & "," & UserList(tIndex).Stats.UserSkills(7) & "," & UserList(tIndex).Stats.UserSkills(8) & "," & UserList(tIndex).Stats.UserSkills(9) & "," & UserList(tIndex).Stats.UserSkills(10) & "," & UserList(tIndex).Stats.UserSkills(11) & "," & UserList(tIndex).Stats.UserSkills(12) & "," & UserList(tIndex).Stats.UserSkills(13) & "," & UserList(tIndex).Stats.UserSkills(14) & "," & UserList(tIndex).Stats.UserSkills(15) & "," & UserList(tIndex).Stats.UserSkills(16) & "," & UserList(tIndex).Stats.UserSkills(17) & "," & UserList(tIndex).Stats.UserSkills(18) & "," & UserList(tIndex).Stats.UserSkills(19) & "," & UserList(tIndex).Stats.UserSkills(20) & "," & UserList(tIndex).Stats.UserSkills(21) _
    & "," & UserList(tIndex).Stats.UserSkills(22) & "," & UserList(tIndex).Stats.UserSkills(23))
    
    Exit Sub
End If

If UCase(rdata) = "/TOOL" Then

Call SendData(ToIndex, Userindex, 0, "HGM")
    
Exit Sub
End If
'[/KEVIN]


'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
If UserList(Userindex).flags.Privilegios < 3 Then
    Exit Sub
End If

'Ban x IP
If UCase(Left(rdata, 6)) = "/BANIP" Then
    Dim BanIP As String, XNick As Boolean
    
    rdata = Right(rdata, Len(rdata) - 7)
    'busca primero la ip del nick
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        XNick = False
        Call LogGM(UserList(Userindex).Name, "/BanIP " & rdata, (UserList(Userindex).flags.Privilegios = 1))
        BanIP = rdata
    Else
        XNick = True
        Call LogGM(UserList(Userindex).Name, "/BanIP " & UserList(tIndex).Name & " - " & UserList(tIndex).ip, (UserList(Userindex).flags.Privilegios = 1))
        BanIP = UserList(tIndex).ip
    End If
    
    'se fija si esta baneada
    For LoopC = 1 To BanIps.Count
        If BanIps.Item(LoopC) = BanIP Then
            Call SendData(ToIndex, Userindex, 0, "||La IP " & BanIP & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
            Exit Sub
        End If
    Next LoopC
    
    BanIps.Add BanIP
    Call SendData(ToAdmins, Userindex, 0, "||" & UserList(Userindex).Name & " Baneo la IP " & BanIP & FONTTYPE_FIGHT)
    
    If XNick = True Then
        Call LogBan(tIndex, Userindex, "Ban por IP desde Nick")
        
        Call SendData(ToAdmins, 0, 0, "||" & UserList(Userindex).Name & " echo a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(Userindex).Name & " Banned a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
        
        'Ponemos el flag de ban a 1
        UserList(tIndex).flags.Ban = 1
        
        Call LogGM(UserList(Userindex).Name, "Echo a " & UserList(tIndex).Name, (UserList(Userindex).flags.Privilegios = 1))
        Call LogGM(UserList(Userindex).Name, "BAN a " & UserList(tIndex).Name, (UserList(Userindex).flags.Privilegios = 1))
        Call CloseSocket(tIndex)
    End If
    
    Exit Sub
End If

'Desbanea una IP
If UCase(Left(rdata, 8)) = "/UNBANIP" Then
    
    
    rdata = Right(rdata, Len(rdata) - 9)
    Call LogGM(UserList(Userindex).Name, "/UNBANIP " & rdata, (UserList(Userindex).flags.Privilegios = 1))
    
    For LoopC = 1 To BanIps.Count
        If BanIps.Item(LoopC) = rdata Then
            BanIps.Remove LoopC
            Call SendData(ToIndex, Userindex, 0, "||La IP " & BanIP & " se ha quitado de la lista de bans." & FONTTYPE_INFO)
            Exit Sub
        End If
    Next LoopC
    
    Call SendData(ToIndex, Userindex, 0, "||La IP " & rdata & " NO se encuentra en la lista de bans." & FONTTYPE_INFO)
    
    Exit Sub
End If

'Crear Teleport
If UCase(Left(rdata, 3)) = "/CT" Then
    '/ct mapa_dest x_dest y_dest
    rdata = Right(rdata, Len(rdata) - 4)
    Call LogGM(UserList(Userindex).Name, "/CT: " & rdata, (UserList(Userindex).flags.Privilegios = 1))
    mapa = ReadField(1, rdata, 32)
    X = ReadField(2, rdata, 32)
    Y = ReadField(3, rdata, 32)
    
    If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y - 1).OBJInfo.ObjIndex > 0 Then
        Exit Sub
    End If
    If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y - 1).TileExit.Map > 0 Then
        Exit Sub
    End If
    If MapaValido(mapa) = False Or InMapBounds(mapa, X, Y) = False Then
        Exit Sub
    End If
    
    Dim ET As Obj
    ET.Amount = 1
    ET.ObjIndex = 378
    
    Call MakeObj(ToMap, 0, UserList(Userindex).Pos.Map, ET, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y - 1)
    MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y - 1).TileExit.Map = mapa
    MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y - 1).TileExit.X = X
    MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y - 1).TileExit.Y = Y
    
    Exit Sub
End If

'Destruir Teleport
'toma el ultimo click
If UCase(Left(rdata, 3)) = "/DT" Then
    '/dt
    Call LogGM(UserList(Userindex).Name, "/DT", (UserList(Userindex).flags.Privilegios = 1))
    
    mapa = UserList(Userindex).flags.TargetMap
    X = UserList(Userindex).flags.TargetX
    Y = UserList(Userindex).flags.TargetY
    
    If ObjData(MapData(mapa, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_TELEPORT And _
        MapData(mapa, X, Y).TileExit.Map > 0 Then
        Call EraseObj(ToMap, 0, mapa, MapData(mapa, X, Y).OBJInfo.Amount, mapa, X, Y)
        MapData(mapa, X, Y).TileExit.Map = 0
        MapData(mapa, X, Y).TileExit.X = 0
        MapData(mapa, X, Y).TileExit.Y = 0
    End If
    
    Exit Sub
End If

'Destruir
If UCase$(Left$(rdata, 5)) = "/DEST" Then
    Call LogGM(UserList(Userindex).Name, "/DEST", (UserList(Userindex).flags.Privilegios = 1))
    rdata = Right$(rdata, Len(rdata) - 5)
    Call EraseObj(ToMap, Userindex, UserList(Userindex).Pos.Map, 10000, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y)
    Exit Sub
End If

'Bloquear
If UCase$(Left$(rdata, 5)) = "/BLOQ" Then
    Call LogGM(UserList(Userindex).Name, "/BLOQ", (UserList(Userindex).flags.Privilegios = 1))
    rdata = Right$(rdata, Len(rdata) - 5)
    If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Blocked = 0 Then
        MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Blocked = 1
        Call Bloquear(ToMap, Userindex, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, 1)
    Else
        MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Blocked = 0
        Call Bloquear(ToMap, Userindex, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, 0)
    End If
    Exit Sub
End If

'Quitar NPC
If UCase$(rdata) = "/MATA" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    If UserList(Userindex).flags.TargetNpc = 0 Then Exit Sub
    Call QuitarNPC(UserList(Userindex).flags.TargetNpc)
    Call LogGM(UserList(Userindex).Name, "/MATA " & Npclist(UserList(Userindex).flags.TargetNpc).Name, (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If

'Quita todos los NPCs del area
If UCase$(rdata) = "/MASSKILL" Then
    For Y = UserList(Userindex).Pos.Y - MinYBorder + 1 To UserList(Userindex).Pos.Y + MinYBorder - 1
            For X = UserList(Userindex).Pos.X - MinXBorder + 1 To UserList(Userindex).Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                    If MapData(UserList(Userindex).Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(UserList(Userindex).Pos.Map, X, Y).NpcIndex)
            Next X
    Next Y
    Call LogGM(UserList(Userindex).Name, "/MASSKILL", (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If

'Quita todos los NPCs del area
If UCase$(rdata) = "/LIMPIAR" Then
        Call LimpiarMundo
        Exit Sub
End If

'Mensaje del sistema
If UCase$(Left$(rdata, 6)) = "/SMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(Userindex).Name, "Mensaje de sistema:" & rdata, (UserList(Userindex).flags.Privilegios = 1))
    Call SendData(ToAll, 0, 0, "!!" & rdata & ENDC)
    
    Exit Sub
End If

'Crear criatura, toma directamente el indice
If UCase$(Left$(rdata, 5)) = "/ACC " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   Call SpawnNpc(val(rdata), UserList(Userindex).Pos, True, False)
   Exit Sub
End If

'Crear criatura con respawn, toma directamente el indice
If UCase$(Left$(rdata, 6)) = "/RACC " Then
   rdata = Right$(rdata, Len(rdata) - 6)
   Call SpawnNpc(val(rdata), UserList(Userindex).Pos, True, True)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AI1 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraImperial1 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AI2 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraImperial1 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AI3 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraImperial3 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AI4 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   TunicaMagoImperial = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AC1 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraCaos1 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AC2 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraCaos2 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AC3 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   ArmaduraCaos3 = val(rdata)
   Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/AC4 " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   TunicaMagoCaos = val(rdata)
   Exit Sub
End If



'Comando para depurar la navegacion
If UCase$(rdata) = "/NAVE" Then
    If UserList(Userindex).flags.Navegando = 1 Then
        UserList(Userindex).flags.Navegando = 0
    Else
        UserList(Userindex).flags.Navegando = 1
    End If
    Exit Sub
End If

'Apagamos
If UCase$(rdata) = "/APAGAR" Then
    If UCase$(UserList(Userindex).Name) <> "NEB" Then
        Call LogGM(UserList(Userindex).Name, "¡¡¡Intento apagar el server!!!", (UserList(Userindex).flags.Privilegios = 1))
        Exit Sub
    End If
    'Log
    mifile = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #mifile
    Print #mifile, Date & " " & Time & " server apagado por " & UserList(Userindex).Name & ". "
    Close #mifile
    Unload frmMain
    Exit Sub
End If

'CONDENA
If UCase$(Left$(rdata, 7)) = "/CONDEN" Then
    rdata = Right$(rdata, Len(rdata) - 8)
    tIndex = NameIndex(rdata)
    If tIndex > 0 Then Call VolverCriminal(tIndex)
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/RAJAR " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(UCase$(rdata))
    If tIndex > 0 Then
        If UserList(tIndex).Faccion.FuerzasCaos > 0 Then
            Call ExpulsarCaos(tIndex)
        ElseIf UserList(tIndex).Faccion.ArmadaReal > 0 Then
            Call ExpulsarFaccionReal(tIndex)
        End If
    End If
    Exit Sub
End If

'[KEVIN]
If UCase$(Left$(rdata, 10)) = "/CREAROBJ " Then
    rdata = Right$(rdata, Len(rdata) - 10)
    Dim ObjC As Obj
    
    ObjC.ObjIndex = ReadField(1, rdata, 32)
    ObjC.Amount = ReadField(2, rdata, 32)
    
    MeterItemEnInventario Userindex, ObjC
    
    Call LogGM(UserList(Userindex).Name, "Creó " & ObjC.Amount & " unidades del objeto " & ObjC.ObjIndex, (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If
'[/KEVIN]

'[KEVIN]
If UCase$(rdata) = "/BOOTALL" Then
    
    Dim BootAll As Integer
    
    For BootAll = 1 To LastUser
        If BootAll <> Userindex Then
            CloseSocket (BootAll)
        End If
    Next BootAll
    
    Call LogGM(UserList(Userindex).Name, "¡¡¡ECHO A TODOS!!!", (UserList(Userindex).flags.Privilegios = 1))
    
    Exit Sub
End If
'[/KEVIN]

'[KEVIN]MODIFICADO
'MODIFICA CARACTER
If UCase(Left$(rdata, 5)) = "/MOD " Then
    Call LogGM(UserList(Userindex).Name, rdata, (UserList(Userindex).flags.Privilegios = 1))
    rdata = Right$(rdata, Len(rdata) - 5)
    tIndex = NameIndex(ReadField(1, rdata, 32))
    Arg1 = ReadField(2, rdata, 32)
    Arg2 = ReadField(3, rdata, 32)
    Arg3 = ReadField(4, rdata, 32)
    Arg4 = ReadField(5, rdata, 32)
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Select Case UCase(Arg1)
    
        Case "ORO"
            
                UserList(tIndex).Stats.GLD = val(Arg2)
                Call SendUserStatsBox(tIndex)
           
        Case "EXP"
            
                If UserList(tIndex).Stats.Exp + val(Arg2) > _
                   UserList(tIndex).Stats.ELU Then
                   Dim resto
                   resto = val(Arg2) - UserList(tIndex).Stats.ELU
                   UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + UserList(tIndex).Stats.ELU
                   Call CheckUserLevel(tIndex)
                   UserList(tIndex).Stats.Exp = UserList(tIndex).Stats.Exp + resto
                Else
                   UserList(tIndex).Stats.Exp = val(Arg2)
                End If
                Call SendUserStatsBox(tIndex)
            
        Case "BODY"
            Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, val(Arg2), UserList(tIndex).Char.Head, UserList(tIndex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim, UserList(Userindex).flags.AdminInvisible, UserList(Userindex).flags.Invisible)
            Exit Sub
        Case "HEAD"
            Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, UserList(tIndex).Char.Body, val(Arg2), UserList(tIndex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim, UserList(Userindex).flags.AdminInvisible, UserList(Userindex).flags.Invisible)
            Exit Sub
        Case "CRI"
            UserList(tIndex).Faccion.CriminalesMatados = val(Arg2)
            Call SendUserStatsBox(tIndex)
            Exit Sub
        Case "CIU"
            UserList(tIndex).Faccion.CiudadanosMatados = val(Arg2)
            Call SendUserStatsBox(tIndex)
            Exit Sub
        Case "LEVEL"
            UserList(tIndex).Stats.ELV = val(Arg2)
            Call SendUserStatsBox(tIndex)
            Exit Sub
        Case "AGIL"
            UserList(tIndex).Stats.UserAtributos(Agilidad) = val(Arg2)
            UserList(tIndex).Stats.UserAtributosBackUP(Agilidad) = val(Arg2)
            Call SendUserStatsBox(tIndex)
            Exit Sub
        Case "FUERZA"
            UserList(tIndex).Stats.UserAtributos(Fuerza) = val(Arg2)
            UserList(tIndex).Stats.UserAtributosBackUP(Fuerza) = val(Arg2)
            Call SendUserStatsBox(tIndex)
            Exit Sub
        Case "CONST"
            UserList(tIndex).Stats.UserAtributos(Constitucion) = val(Arg2)
            UserList(tIndex).Stats.UserAtributosBackUP(Constitucion) = val(Arg2)
            Call SendUserStatsBox(tIndex)
            Exit Sub
        Case "INTEL"
            UserList(tIndex).Stats.UserAtributos(Inteligencia) = val(Arg2)
            UserList(tIndex).Stats.UserAtributosBackUP(Inteligencia) = val(Arg2)
            Call SendUserStatsBox(tIndex)
            Exit Sub
        Case "CARIS"
            UserList(tIndex).Stats.UserAtributos(Carisma) = val(Arg2)
            UserList(tIndex).Stats.UserAtributosBackUP(Carisma) = val(Arg2)
            Call SendUserStatsBox(tIndex)
            Exit Sub
        Case "MAXHP"
            UserList(tIndex).Stats.MaxHP = val(Arg2)
            UserList(tIndex).Stats.MinHP = val(Arg2)
            Call SendUserStatsBox(tIndex)
            Exit Sub
        Case "MAXSTA"
            UserList(tIndex).Stats.MaxSta = val(Arg2)
            UserList(tIndex).Stats.MinSta = val(Arg2)
            Call SendUserStatsBox(tIndex)
            Exit Sub
        Case "MAXMAN"
            UserList(tIndex).Stats.MaxMAN = val(Arg2)
            UserList(tIndex).Stats.MinMAN = val(Arg2)
            Call SendUserStatsBox(tIndex)
            Exit Sub
        Case "MAXHIT"
            UserList(tIndex).Stats.MaxHIT = val(Arg2)
            UserList(tIndex).Stats.MinHIT = val(Arg2) - 1
            UserList(tIndex).Stats.MaxHitBK = val(Arg2)
            UserList(tIndex).Stats.MinHitBK = val(Arg2) - 1
            Call SendUserStatsBox(tIndex)
            Exit Sub
        Case "DEF"
            UserList(tIndex).Stats.Def = val(Arg2)
            Call SendUserStatsBox(tIndex)
            Exit Sub
        Case Else
            Call SendData(ToIndex, Userindex, 0, "||Comando no permitido." & FONTTYPE_INFO)
            Exit Sub
    End Select

    Exit Sub
End If
'[/KEVIN]


If UCase$(Left$(rdata, 9)) = "/DOBACKUP" Then
    Call DoBackUp
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/GRABAR" Then
    Call GuardarUsuarios
    Exit Sub
End If

If UCase$(Left$(rdata, 11)) = "/BORRAR SOS" Then
    Call Ayuda.Reset
    Exit Sub
End If

'[KEVIN]
If UCase(Left$(rdata, 11)) = "/BORRAR INS" Then
    Call InscTorneo.Reset
    Exit Sub
End If
'[/KEVIN]

If UCase$(Left$(rdata, 9)) = "/SHOW INT" Then
    Call frmMain.mnuMostrar_Click
    Exit Sub
End If

If UCase$(rdata) = "/LLUVIA" Then
    Lloviendo = Not Lloviendo
    Call SendData(ToAll, 0, 0, "LLU")
    Exit Sub
End If

If UCase$(rdata) = "/PASSDAY" Then
    Call DayElapsed
    Exit Sub
End If

'[KEVIN]
If UCase$(Left$(rdata, 7)) = "/CURAR " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(UCase(rdata))
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    UserList(tIndex).Stats.MinHP = UserList(tIndex).Stats.MaxHP
    Call SendUserStatsBox(tIndex)
    
    
    Exit Sub
End If
'[/KEVIN]

'[KEVIN]
If UCase$(Left$(rdata, 8)) = "$USKILL " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    
    tIndex = NameIndex(UCase(rdata))
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, Userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(ToIndex, Userindex, 0, "PNSKILL" & UserList(tIndex).Stats.UserSkills(1) & "," & UserList(tIndex).Stats.UserSkills(2) & "," & UserList(tIndex).Stats.UserSkills(3) & "," & UserList(tIndex).Stats.UserSkills(4) & "," & UserList(tIndex).Stats.UserSkills(5) & "," & UserList(tIndex).Stats.UserSkills(6) & "," & UserList(tIndex).Stats.UserSkills(7) & "," & UserList(tIndex).Stats.UserSkills(8) & "," & UserList(tIndex).Stats.UserSkills(9) & "," & UserList(tIndex).Stats.UserSkills(10) & "," & UserList(tIndex).Stats.UserSkills(11) & "," & UserList(tIndex).Stats.UserSkills(12) & "," & UserList(tIndex).Stats.UserSkills(13) & "," & UserList(tIndex).Stats.UserSkills(14) & "," & UserList(tIndex).Stats.UserSkills(15) & "," & UserList(tIndex).Stats.UserSkills(16) & "," & UserList(tIndex).Stats.UserSkills(17) & "," & UserList(tIndex).Stats.UserSkills(18) & "," & UserList(tIndex).Stats.UserSkills(19) & "," & UserList(tIndex).Stats.UserSkills(20) & "," & UserList(tIndex).Stats.UserSkills(21) _
    & "," & UserList(tIndex).Stats.UserSkills(22) & "," & UserList(tIndex).Stats.UserSkills(23))
    
    Exit Sub
End If

'(PA CAMBIAR EL INTERVALO DE LOS MACROS)
If UCase$(Left$(rdata, 8)) = "/MACROI " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    
    Dim Interval As Integer
    
    If Not Numeric(rdata) Then Exit Sub
    
    Interval = rdata
    
    If Interval > 5000 Then
        Call SendData(ToIndex, Userindex, 0, "||Debes colocar un número menor a 5000!!")
        Exit Sub
    End If
    
    Call SendData(ToAll, 0, 0, "CMR" & Interval)
    
    Exit Sub
End If
'[/KEVIN]

If UCase(rdata) = "/RESETEAR" Then
    Call Restart
    Exit Sub
End If

'[KEVIN]
If UCase(rdata) = "/GTOOL" Then

Call SendData(ToIndex, Userindex, 0, "GMT")
    
Exit Sub
End If

'Agregar un Dios
If UCase$(Left$(rdata, 13)) = "/AGREGARDIOS " Then
    rdata = Right$(rdata, Len(rdata) - 13)
    Dim Temp, GmNum, NumGMs As Integer
    Arg1 = rdata
    
    If UCase(UserList(Userindex).Name) <> "NEB" Then Exit Sub
    
    If FileExist(CharPath & UCase(Arg1) & ".chr", vbNormal) = False Then
        Call SendData(ToIndex, Userindex, 0, "||El usuario no existe.")
        Exit Sub
    End If
    
    NumGMs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
    For GmNum = 1 To NumGMs
        If UCase(Arg1) = UCase(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & GmNum)) Then
            Call SendData(ToIndex, Userindex, 0, "||El usuario ya es dios.")
            Exit Sub
        End If
    Next GmNum
    
    Temp = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses")) + 1
    Call WriteVar(IniPath & "Server.ini", "INIT", "Dioses", Str(Temp))
    Call WriteVar(IniPath & "Server.ini", "Dioses", "Dios" & Temp, Arg1)
    Call SendData(ToIndex, Userindex, 0, "||Has agregado a " & Arg1 & " a la lista de los Dioses." & FONTTYPE_INFO)
    
    tIndex = NameIndex(Arg1)
    If tIndex > 0 Then
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " te agregó a la lista de los Dioses." & FONTTYPE_INFO)
    Call LogGM(UserList(Userindex).Name, "Agregó a " & rdata & " a la lista de los Dioses", (UserList(Userindex).flags.Privilegios = 1))
    End If
    Exit Sub
End If

'Agregar un Semi Dios
If UCase$(Left$(rdata, 13)) = "/AGREGARSEMI " Then
    rdata = Right$(rdata, Len(rdata) - 13)
    Dim NumTemp, GmsNum, NumSGMs As Integer
    Arg1 = rdata
    
    If UCase(UserList(Userindex).Name) <> "NEB" Then Exit Sub
    
    If FileExist(CharPath & UCase(Arg1) & ".chr", vbNormal) = False Then
        Call SendData(ToIndex, Userindex, 0, "||El usuario no existe.")
        Exit Sub
    End If
    
    NumSGMs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
    For GmsNum = 1 To NumSGMs
        If UCase(Arg1) = UCase(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & GmsNum)) Then
            Call SendData(ToIndex, Userindex, 0, "||El usuario ya es SemiDios.")
            Exit Sub
        End If
    Next GmsNum
    
    NumTemp = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses")) + 1
    Call WriteVar(IniPath & "Server.ini", "INIT", "SemiDioses", Str(NumTemp))
    Call WriteVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & NumTemp, Arg1)
    Call SendData(ToIndex, Userindex, 0, "||Has agregado a " & Arg1 & " a la lista de los SemiDioses." & FONTTYPE_INFO)
    
    tIndex = NameIndex(Arg1)
    If tIndex > 0 Then
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " te agregó a la lista de los SemiDioses." & FONTTYPE_INFO)
    Call LogGM(UserList(Userindex).Name, "Agregó a " & rdata & " a la lista de los SemiDioses", (UserList(Userindex).flags.Privilegios = 1))
    End If
    Exit Sub
End If

'Echar un Dios
If UCase$(Left$(rdata, 11)) = "/ECHARDIOS " Then
    rdata = Right$(rdata, Len(rdata) - 11)
    Dim NumGMGs, GmGNum As Integer
    Arg1 = rdata
    
    If UCase(Arg1) = "NEB" Then Exit Sub
    If UCase(Arg1) = "CUCSIFAE" Then Exit Sub
    If UCase(Arg1) = "MAGNUM" Then Exit Sub
    
    NumGMGs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
    For GmGNum = 1 To NumGMGs
        If UCase(Arg1) = UCase(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & GmGNum)) Then
            Call WriteVar(IniPath & "Server.ini", "Dioses", "Dios" & GmGNum, "¡ECHADO!")
            Call SendData(ToIndex, Userindex, 0, "||Has removido a " & Arg1 & " de la lista de Dioses." & FONTTYPE_INFO)
            tIndex = NameIndex(Arg1)
            If tIndex > 0 Then
                Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " te a removido de la lista de Dioses." & FONTTYPE_INFO)
                'Call modName(tIndex)
            End If
            Call LogGM(UserList(Userindex).Name, "Echó a " & rdata & " de la lista de los Dioses", (UserList(Userindex).flags.Privilegios = 1))
            Exit Sub
        End If
    Next GmGNum
    Call SendData(ToIndex, Userindex, 0, "||" & Arg1 & " no fue encontrado en la lista de los dioses." & FONTTYPE_INFO)
    Exit Sub
End If

'Echar un SemiDios
If UCase$(Left$(rdata, 11)) = "/ECHARSEMI " Then
    rdata = Right$(rdata, Len(rdata) - 11)
    Dim NumGMs2, GmNum2 As Integer
    Arg1 = rdata

    NumGMs2 = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
    For GmNum2 = 1 To NumGMs2
        If UCase(Arg1) = UCase(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & GmNum2)) Then
            Call WriteVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & GmNum2, "¡ECHADO!")
            Call SendData(ToIndex, Userindex, 0, "||Has removido a " & Arg1 & " de la lista de Dioses." & FONTTYPE_INFO)
            tIndex = NameIndex(Arg1)
            If tIndex > 0 Then
                Call SendData(ToIndex, tIndex, 0, "||" & UserList(Userindex).Name & " te a removido de la lista de SemiDioses." & FONTTYPE_INFO)
                'Call modName(tIndex)
            End If
            Call LogGM(UserList(Userindex).Name, "Echó a " & rdata & " de la lista de los SemiDioses", (UserList(Userindex).flags.Privilegios = 1))
            Exit Sub
        End If
    Next GmNum2
    Call SendData(ToIndex, Userindex, 0, "||" & Arg1 & " no fue encontrado en la lista de los SemiDioses." & FONTTYPE_INFO)
    Exit Sub
End If

'Teleport en masa (todos los pjs del area)
If UCase$(Left$(rdata, 10)) = "/TELEPALL " Then
    rdata = Right$(rdata, Len(rdata) - 10)
    mapa = val(ReadField(1, rdata, 32))
    If Not MapaValido(mapa) Then Exit Sub
    Dim X2 As Integer
    Dim Y2 As Integer
    X2 = val(ReadField(2, rdata, 32))
    Y2 = val(ReadField(3, rdata, 32))
    If Not InMapBounds(mapa, X2, Y2) Then Exit Sub
    
    For Y = UserList(Userindex).Pos.Y - MinYBorder + 1 To UserList(Userindex).Pos.Y + MinYBorder - 1
            For X = UserList(Userindex).Pos.X - MinXBorder + 1 To UserList(Userindex).Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                    If MapData(UserList(Userindex).Pos.Map, X, Y).Userindex > 0 Then Call WarpUserChar(MapData(UserList(Userindex).Pos.Map, X, Y).Userindex, mapa, X2, Y2, True)
            Next X
    Next Y
    
    Call LogGM(UserList(Userindex).Name, "Hizo un teleport en masa al " & "Mapa" & mapa & " X:" & X2 & " Y:" & Y2, (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub

End If

'Destrabar Usuario
If UCase$(Left$(rdata, 11)) = "/DESTRABAR " Then
    rdata = Right$(rdata, Len(rdata) - 11)
    mapa = val(ReadField(2, rdata, Asc("-")))
    
    If Not MapaValido(mapa) Then
        Call SendData(ToIndex, Userindex, 0, "||Mapa Incorrecto.")
        Exit Sub
    End If
    
    Name = ReadField(1, rdata, 45)
    X = val(ReadField(3, rdata, 45))
    Y = val(ReadField(4, rdata, 45))
    
    If FileExist(CharPath & UCase(Name) & ".chr", vbNormal) = False Then
        Call SendData(ToIndex, Userindex, 0, "||El usuario no existe.")
        Exit Sub
    End If
    
    Call WriteVar(CharPath & Name & ".chr", "INIT", "Position", mapa & "-" & X & "-" & Y)
    
    Exit Sub
End If

'CAMBIAR LA PASS
If UCase$(Left$(rdata, 6)) = "/PASS " Then
    'rdata = Right$(rdata, Len(rdata) - 6)
    
    Dim Pasd As String
    Dim Nme As String
    
    Nme = ReadField(1, rdata, 32)
    Pasd = ReadField(2, rdata, 32)
    
    If FileExist(CharPath & UCase(Nme) & ".chr", vbNormal) = False Then
        Call SendData(ToIndex, Userindex, 0, "||El usuario no existe.")
        Exit Sub
    End If
    
    Call WriteVar(CharPath & UCase(Nme) & ".chr", "INIT", "Password", MD5String(Pasd))
    Call LogGM(UserList(Userindex).Name, "Le cambio el password a " & Nme, (UserList(Userindex).flags.Privilegios = 1))
    Exit Sub
End If

'[/KEVIN]


Exit Sub

errorhandler:
 'MsgBox ("Error en mdlComandosGMS")
 Call LogError("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(Userindex).Name & "UI:" & Userindex & " N: " & Err.number & " D: " & Err.Description)
 'Call CloseSocket(UserIndex)
 Call CloseSocketSL(Userindex)
 Call Cerrar_Usuario(Userindex)

End Sub
