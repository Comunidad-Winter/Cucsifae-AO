Attribute VB_Name = "TCP"
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

'Buffer en bytes de cada socket
Public Const SOCKET_BUFFER_SIZE = 2048

'Cuantos comandos de cada cliente guarda el server
Public Const COMMAND_BUFFER_SIZE = 1000

Public Const NingunArma = 2

'RUTAS DE ENVIO DE DATOS
Public Const ToIndex = 0 'Envia a un solo User
Public Const ToAll = 1 'A todos los Users
Public Const ToMap = 2 'Todos los Usuarios en el mapa
Public Const ToPCArea = 3 'Todos los Users en el area de un user determinado
Public Const ToNone = 4 'Ninguno
Public Const ToAllButIndex = 5 'Todos menos el index
Public Const ToMapButIndex = 6 'Todos en el mapa menos el indice
Public Const ToGM = 7
Public Const ToNPCArea = 8 'Todos los Users en el area de un user determinado
Public Const ToGuildMembers = 9
Public Const ToAdmins = 10

' General constants used with most of the controls
Public Const INVALID_HANDLE = -1
Public Const CONTROL_ERRIGNORE = 0
Public Const CONTROL_ERRDISPLAY = 1


' SocietWrench Control Actions
Public Const SOCKET_OPEN = 1
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_LISTEN = 3
Public Const SOCKET_ACCEPT = 4
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_CLOSE = 7
Public Const SOCKET_DISCONNECT = 7
Public Const SOCKET_ABORT = 8

' SocketWrench Control States
Public Const SOCKET_NONE = 0
Public Const SOCKET_IDLE = 1
Public Const SOCKET_LISTENING = 2
Public Const SOCKET_CONNECTING = 3
Public Const SOCKET_ACCEPTING = 4
Public Const SOCKET_RECEIVING = 5
Public Const SOCKET_SENDING = 6
Public Const SOCKET_CLOSING = 7

' Societ Address Families
Public Const AF_UNSPEC = 0
Public Const AF_UNIX = 1
Public Const AF_INET = 2

' Societ Types
Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5

' Protocol Types
Public Const IPPROTO_IP = 0
Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_GGP = 2
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_PUP = 12
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_IDP = 22
Public Const IPPROTO_ND = 77
Public Const IPPROTO_RAW = 255
Public Const IPPROTO_MAX = 256


' Network Addpesses
Public Const INADDR_ANY = "0.0.0.0"
Public Const INADDR_LOOPBACK = "127.0.0.1"
Public Const INADDR_NONE = "255.055.255.255"

' Shutdown Values
Public Const SOCKET_READ = 0
Public Const SOCKET_WRITE = 1
Public Const SOCKET_READWRITE = 2

' SocketWrench Error Pesponse
Public Const SOCKET_ERRIGNORE = 0
Public Const SOCKET_ERRDISPLAY = 1

' SocketWrench Error Aodes
Public Const WSABASEERR = 24000
Public Const WSAEINTR = 24004
Public Const WSAEBADF = 24009
Public Const WSAEACCES = 24013
Public Const WSAEFAULT = 24014
Public Const WSAEINVAL = 24022
Public Const WSAEMFILE = 24024
Public Const WSAEWOULDBLOCK = 24035
Public Const WSAEINPROGRESS = 24036
Public Const WSAEALREADY = 24037
Public Const WSAENOTSOCK = 24038
Public Const WSAEDESTADDRREQ = 24039
Public Const WSAEMSGSIZE = 24040
Public Const WSAEPROTOTYPE = 24041
Public Const WSAENOPROTOOPT = 24042
Public Const WSAEPROTONOSUPPORT = 24043
Public Const WSAESOCKTNOSUPPORT = 24044
Public Const WSAEOPNOTSUPP = 24045
Public Const WSAEPFNOSUPPORT = 24046
Public Const WSAEAFNOSUPPORT = 24047
Public Const WSAEADDRINUSE = 24048
Public Const WSAEADDRNOTAVAIL = 24049
Public Const WSAENETDOWN = 24050
Public Const WSAENETUNREACH = 24051
Public Const WSAENETRESET = 24052
Public Const WSAECONNABORTED = 24053
Public Const WSAECONNRESET = 24054
Public Const WSAENOBUFS = 24055
Public Const WSAEISCONN = 24056
Public Const WSAENOTCONN = 24057
Public Const WSAESHUTDOWN = 24058
Public Const WSAETOOMANYREFS = 24059
Public Const WSAETIMEDOUT = 24060
Public Const WSAECONNREFUSED = 24061
Public Const WSAELOOP = 24062
Public Const WSAENAMETOOLONG = 24063
Public Const WSAEHOSTDOWN = 24064
Public Const WSAEHOSTUNREACH = 24065
Public Const WSAENOTEMPTY = 24066
Public Const WSAEPROCLIM = 24067
Public Const WSAEUSERS = 24068
Public Const WSAEDQUOT = 24069
Public Const WSAESTALE = 24070
Public Const WSAEREMOTE = 24071
Public Const WSASYSNOTREADY = 24091
Public Const WSAVERNOTSUPPORTED = 24092
Public Const WSANOTINITIALISED = 24093
Public Const WSAHOST_NOT_FOUND = 25001
Public Const WSATRY_AGAIN = 25002
Public Const WSANO_RECOVERY = 25003
Public Const WSANO_DATA = 25004
Public Const WSANO_ADDRESS = 2500

Sub DarCuerpoYCabeza(UserBody As Integer, UserHead As Integer, Raza As String, Gen As String)

Select Case Gen
   Case "Hombre"
        Select Case Raza
        
                Case "Humano"
                    UserHead = CInt(RandomNumber(1, 14))
                    If UserHead > 14 Then UserHead = 14
                    UserBody = 1
                Case "Elfo"
                    UserHead = CInt(RandomNumber(1, 4)) + 100
                    If UserHead > 104 Then UserHead = 104
                    UserBody = 2
                Case "Elfo Oscuro"
                    UserHead = CInt(RandomNumber(1, 3)) + 200
                    If UserHead > 203 Then UserHead = 203
                    UserBody = 3
                Case "Enano"
                    UserHead = RandomNumber(1, 2) + 300
                    If UserHead > 302 Then UserHead = 302
                    UserBody = 52
                Case "Gnomo"
                    UserHead = RandomNumber(1, 1) + 400
                    If UserHead > 401 Then UserHead = 401
                    UserBody = 52
                '[KEVIN]
                Case "Orco"
                    UserHead = RandomNumber(1, 5) + 501
                    If UserHead > 506 Then UserHead = 506
                    UserBody = 210
                '[/KEVIN]
                Case Else
                    UserHead = 1
                    UserBody = 1
            
        End Select
   Case "Mujer"
        Select Case Raza
                Case "Humano"
                    UserHead = CInt(RandomNumber(1, 5)) + 69
                    If UserHead > 74 Then UserHead = 74
                    UserBody = 1
                Case "Elfo"
                    UserHead = CInt(RandomNumber(1, 5)) + 169
                    If UserHead > 174 Then UserHead = 174
                    UserBody = 2
                Case "Elfo Oscuro"
                    UserHead = CInt(RandomNumber(1, 6)) + 269
                    If UserHead > 275 Then UserHead = 275
                    UserBody = 3
                Case "Gnomo"
                    UserHead = RandomNumber(1, 2) + 469
                    If UserHead > 471 Then UserHead = 471
                    UserBody = 52
                Case "Enano"
                    UserHead = 370
                    UserBody = 52
                '[KEVIN]
                Case "Orco"
                    UserHead = RandomNumber(1, 2) + 506
                    If UserHead > 508 Then UserHead = 508
                    UserBody = 210
                '[/KEVIN]
                Case Else
                    UserHead = 70
                    UserBody = 1
        End Select
End Select

   
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function


Function NombrePermitido(ByVal Nombre As String) As Boolean
Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(Nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
    End If
Next i

NombrePermitido = True

End Function

Function ValidateAtrib(ByVal UserIndex As Integer, ByVal UserRaza As String) As Boolean
Dim LoopC As Integer
Dim AAgilidad As Integer
Dim AFuerza As Integer
Dim ACarisma As Integer
Dim AConstitucion As Integer
Dim AInteligencia As Integer

AAgilidad = UserList(UserIndex).Stats.UserAtributos(Agilidad)
AFuerza = UserList(UserIndex).Stats.UserAtributos(Fuerza)
ACarisma = UserList(UserIndex).Stats.UserAtributos(Carisma)
AConstitucion = UserList(UserIndex).Stats.UserAtributos(Constitucion)
AInteligencia = UserList(UserIndex).Stats.UserAtributos(Inteligencia)

For LoopC = 1 To NUMATRIBUTOS
    If UserList(UserIndex).Stats.UserAtributos(LoopC) > 20 Or UserList(UserIndex).Stats.UserAtributos(LoopC) < 1 Then Exit Function
Next LoopC
    
Select Case UCase(UserRaza)
    Case "HUMANO"
        If AFuerza = 22 And _
            AAgilidad = 20 And _
            AInteligencia = 18 And _
            ACarisma = 20 And _
            AConstitucion = 21 Then
            
            ValidateAtrib = False
            Exit Function
        End If
    Case "ELFO"
        If AFuerza = 19 And _
            AAgilidad = 21 And _
            AInteligencia = 23 And _
            ACarisma = 20 And _
            AConstitucion = 20 Then
            
            ValidateAtrib = False
            Exit Function
        End If
    Case "ELFO OSCURO"
        If AFuerza = 21 And _
            AAgilidad = 20 And _
            AInteligencia = 21 And _
            ACarisma = 19 And _
            AConstitucion = 21 Then
            
            ValidateAtrib = False
            Exit Function
        End If
    Case "ENANO"
        If AFuerza = 22 And _
            AAgilidad = 20 And _
            AInteligencia = 15 And _
            ACarisma = 19 And _
            AConstitucion = 24 Then
            
            ValidateAtrib = False
            Exit Function
        End If
    Case "GNOMO"
        If AFuerza = 17 And _
            AAgilidad = 22 And _
            AInteligencia = 24 And _
            ACarisma = 22 And _
            AConstitucion = 19 Then
            
            ValidateAtrib = False
            Exit Function
        End If
    Case "ORCO"
        If AFuerza = 24 And _
            AAgilidad = 18 And _
            AInteligencia = 15 And _
            ACarisma = 18 And _
            AConstitucion = 23 Then
            
            ValidateAtrib = False
            Exit Function
        End If
End Select

ValidateAtrib = True

End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
        Exit Function
        If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100
    End If
Next LoopC

ValidateSkills = True
    

End Function

Sub ConnectNewUser(UserIndex As Integer, Name As String, Password As String, Body As Integer, Head As Integer, UserRaza As String, UserSexo As String, UserClase As String, _
UA1 As String, UA2 As String, UA3 As String, UA4 As String, UA5 As String, _
US1 As String, US2 As String, US3 As String, US4 As String, US5 As String, _
US6 As String, US7 As String, US8 As String, US9 As String, US10 As String, _
US11 As String, US12 As String, US13 As String, US14 As String, US15 As String, _
US16 As String, US17 As String, US18 As String, US19 As String, US20 As String, _
US21 As String, US22 As String, US23 As String, US24 As String, UserEmail As String, Hogar As String)

If Not NombrePermitido(Name) Then
    Call SendData(ToIndex, UserIndex, 0, "ERRLos nombres de los personajes deben pertencer a la fantasia, el nombre indicado es invalido.")
    Exit Sub
End If

If Not AsciiValidos(Name) Then
    Call SendData(ToIndex, UserIndex, 0, "ERRNombre invalido.")
    Exit Sub
End If

If Name = "" Then
    Call SendData(ToIndex, UserIndex, 0, "ERRDebes escribir tu nick.")
    Exit Sub
End If

Dim LoopC As Integer
Dim totalskpts As Long
  
'¿Existe el personaje?
If FileExist(CharPath & UCase(Name) & ".chr", vbNormal) = True Then
    Call SendData(ToIndex, UserIndex, 0, "ERRYa existe el personaje.")
    Exit Sub
End If

UserList(UserIndex).Flags.Muerto = 0
UserList(UserIndex).Flags.Escondido = 0



UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.BurguesRep = 0
UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.NobleRep = 1000
UserList(UserIndex).Reputacion.PlebeRep = 30

UserList(UserIndex).Reputacion.Promedio = 30 / 6


UserList(UserIndex).Name = Name
UserList(UserIndex).Clase = UserClase
UserList(UserIndex).Raza = UserRaza
UserList(UserIndex).Genero = UserSexo
UserList(UserIndex).Email = UserEmail
UserList(UserIndex).Hogar = Hogar

UserList(UserIndex).Stats.UserAtributos(Fuerza) = Abs(CInt(UA1))
UserList(UserIndex).Stats.UserAtributos(Inteligencia) = Abs(CInt(UA2))
UserList(UserIndex).Stats.UserAtributos(Agilidad) = Abs(CInt(UA3))
UserList(UserIndex).Stats.UserAtributos(Carisma) = Abs(CInt(UA4))
UserList(UserIndex).Stats.UserAtributos(Constitucion) = Abs(CInt(UA5))

UserList(UserIndex).ClaseModificadorEvasion = ModificadorEvasion(UserList(UserIndex).Clase)
UserList(UserIndex).ClaseModificadorPoderAtaqueArmas = ModificadorPoderAtaqueArmas(UserList(UserIndex).Clase)
UserList(UserIndex).ClaseModificadorPoderAtaqueProyectiles = ModificadorPoderAtaqueProyectiles(UserList(UserIndex).Clase)
UserList(UserIndex).ClaseModicadorDañoClaseArmas = ModicadorDañoClaseArmas(UserList(UserIndex).Clase)
UserList(UserIndex).ClaseModificadorPoderAtaqueProyectiles = ModicadorDañoClaseProyectiles(UserList(UserIndex).Clase)
UserList(UserIndex).ClaseModEvasionDeEscudoClase = ModEvasionDeEscudoClase(UserList(UserIndex).Clase)
    


'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS ATRIBUTOS %%%%%%%%%%%%%
If Not ValidateAtrib(UserIndex, UserRaza) Then
        Call SendData(ToIndex, UserIndex, 0, "ERRAtributos invalidos.")
        Exit Sub
End If
'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS ATRIBUTOS %%%%%%%%%%%%%

Select Case UCase(UserRaza)
    Case "HUMANO"
        UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) + 2
        UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) - 2
        UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 1
    Case "ELFO"
        UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) + 1
        UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) - 1
        UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) + 3
    Case "ELFO OSCURO"
        UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) + 1
        UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) + 1
        UserList(UserIndex).Stats.UserAtributos(Carisma) = UserList(UserIndex).Stats.UserAtributos(Carisma) - 1
        UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 1
    Case "ENANO"
        UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) + 2
        UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 4
        UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) - 5
        UserList(UserIndex).Stats.UserAtributos(Carisma) = UserList(UserIndex).Stats.UserAtributos(Carisma) - 1
    Case "GNOMO"
        UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) - 3
        UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) + 4
        UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) + 2
        UserList(UserIndex).Stats.UserAtributos(Carisma) = UserList(UserIndex).Stats.UserAtributos(Carisma) + 2
        UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) - 1
    Case "ORCO"
        UserList(UserIndex).Stats.UserAtributos(Fuerza) = UserList(UserIndex).Stats.UserAtributos(Fuerza) + 4
        UserList(UserIndex).Stats.UserAtributos(Inteligencia) = UserList(UserIndex).Stats.UserAtributos(Inteligencia) - 5
        UserList(UserIndex).Stats.UserAtributos(Agilidad) = UserList(UserIndex).Stats.UserAtributos(Agilidad) - 2
        UserList(UserIndex).Stats.UserAtributos(Constitucion) = UserList(UserIndex).Stats.UserAtributos(Constitucion) + 3
        UserList(UserIndex).Stats.UserAtributos(Carisma) = UserList(UserIndex).Stats.UserAtributos(Carisma) - 2
End Select



UserList(UserIndex).Stats.UserSkills(1) = val(US1)
UserList(UserIndex).Stats.UserSkills(2) = val(US2)
UserList(UserIndex).Stats.UserSkills(3) = val(US3)
UserList(UserIndex).Stats.UserSkills(4) = val(US4)
UserList(UserIndex).Stats.UserSkills(5) = val(US5)
UserList(UserIndex).Stats.UserSkills(6) = val(US6)
UserList(UserIndex).Stats.UserSkills(7) = val(US7)
UserList(UserIndex).Stats.UserSkills(8) = val(US8)
UserList(UserIndex).Stats.UserSkills(9) = val(US9)
UserList(UserIndex).Stats.UserSkills(10) = val(US10)
UserList(UserIndex).Stats.UserSkills(11) = val(US11)
UserList(UserIndex).Stats.UserSkills(12) = val(US12)
UserList(UserIndex).Stats.UserSkills(13) = val(US13)
UserList(UserIndex).Stats.UserSkills(14) = val(US14)
UserList(UserIndex).Stats.UserSkills(15) = val(US15)
UserList(UserIndex).Stats.UserSkills(16) = val(US16)
UserList(UserIndex).Stats.UserSkills(17) = val(US17)
UserList(UserIndex).Stats.UserSkills(18) = val(US18)
UserList(UserIndex).Stats.UserSkills(19) = val(US19)
UserList(UserIndex).Stats.UserSkills(20) = val(US20)
UserList(UserIndex).Stats.UserSkills(21) = val(US21)
'[KEVIN]
UserList(UserIndex).Stats.UserSkills(22) = val(US22)
UserList(UserIndex).Stats.UserSkills(23) = val(US23)
UserList(UserIndex).Stats.UserSkills(24) = val(US24)
'[/KEVIN]

totalskpts = 0

'Abs PREVINENE EL HACKEO DE LOS SKILLS %%%%%%%%%%%%%
For LoopC = 1 To NUMSKILLS
    totalskpts = totalskpts + Abs(UserList(UserIndex).Stats.UserSkills(LoopC))
Next LoopC


If totalskpts > 10 Then
    Call LogHackAttemp(UserList(UserIndex).Name & " intento hackear los skills.")
    Call BorrarUsuario(UserList(UserIndex).Name)
    Call CloseSocket(UserIndex)
    Exit Sub
End If
'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%

UserList(UserIndex).Password = Password
UserList(UserIndex).Char.Heading = SOUTH

Call Randomize(Timer)
Call DarCuerpoYCabeza(UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Raza, UserList(UserIndex).Genero)
UserList(UserIndex).OrigChar = UserList(UserIndex).Char
   
 
UserList(UserIndex).Char.WeaponAnim = NingunArma
UserList(UserIndex).Char.ShieldAnim = NingunEscudo
UserList(UserIndex).Char.CascoAnim = NingunCasco

'[KEVIN]
UserList(UserIndex).Stats.Matrimonio = ""
'[KEVIN]

UserList(UserIndex).Stats.MET = 1
Dim MiInt
MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Constitucion) \ 3)

UserList(UserIndex).Stats.MaxHP = 15 + MiInt
UserList(UserIndex).Stats.MinHP = 15 + MiInt

UserList(UserIndex).Stats.FIT = 1


MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(UserIndex).Stats.MaxSta = 20 * MiInt
UserList(UserIndex).Stats.MinSta = 20 * MiInt


UserList(UserIndex).Stats.MaxAGU = 100
UserList(UserIndex).Stats.MinAGU = 100

UserList(UserIndex).Stats.MaxHam = 100
UserList(UserIndex).Stats.MinHam = 100


'<-----------------MANA----------------------->
If UserClase = "Mago" Then
    MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Inteligencia)) / 3
    UserList(UserIndex).Stats.MaxMAN = 100 + MiInt
    UserList(UserIndex).Stats.MinMAN = 100 + MiInt
ElseIf UserClase = "Clerigo" Or UserClase = "Druida" _
    Or UserClase = "Bardo" Or UserClase = "Asesino" Then
        MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(Inteligencia)) / 4
        UserList(UserIndex).Stats.MaxMAN = 50
        UserList(UserIndex).Stats.MinMAN = 50
Else
    UserList(UserIndex).Stats.MaxMAN = 0
    UserList(UserIndex).Stats.MinMAN = 0
End If

If UserClase = "Mago" Or UserClase = "Clerigo" Or _
   UserClase = "Druida" Or UserClase = "Bardo" Or _
   UserClase = "Asesino" Then
        UserList(UserIndex).Stats.UserHechizos(1) = 1
        UserList(UserIndex).Stats.UserHechizos(2) = 2
        UserList(UserIndex).Stats.UserHechizos(3) = 6
End If

UserList(UserIndex).Stats.MaxHIT = 2
UserList(UserIndex).Stats.MinHIT = 1

UserList(UserIndex).Stats.GLD = 0

'[KEVIN]
UserList(UserIndex).Quest.EnQuest = 0
UserList(UserIndex).Quest.Quest = 0
UserList(UserIndex).Quest.RealizoQuest = 0
UserList(UserIndex).Quest.RealizoQuest = 0
'[/KEVIN]

UserList(UserIndex).Stats.Exp = 0
UserList(UserIndex).Stats.ELU = 300
UserList(UserIndex).Stats.ELV = 1


'???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
UserList(UserIndex).Invent.NroItems = 9

UserList(UserIndex).Invent.Object(1).ObjIndex = 467
UserList(UserIndex).Invent.Object(1).Amount = 100

UserList(UserIndex).Invent.Object(2).ObjIndex = 468
UserList(UserIndex).Invent.Object(2).Amount = 100

UserList(UserIndex).Invent.Object(3).ObjIndex = 460
UserList(UserIndex).Invent.Object(3).Amount = 1
UserList(UserIndex).Invent.Object(3).Equipped = 1

Select Case UserRaza
    Case "Humano"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 463
    Case "Elfo"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 464
    Case "Elfo Oscuro"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 465
    Case "Enano"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 466
    Case "Gnomo"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 466
    '[KEVIN]
    Case "Orco"
        UserList(UserIndex).Invent.Object(4).ObjIndex = 737
    '[/KEVIN]
End Select

'[KEVIN] MODIFICADO

UserList(UserIndex).Invent.Object(4).Amount = 1
UserList(UserIndex).Invent.Object(4).Equipped = 1

UserList(UserIndex).Invent.Object(5).ObjIndex = 461
UserList(UserIndex).Invent.Object(5).Amount = 20

UserList(UserIndex).Invent.Object(6).ObjIndex = 462
UserList(UserIndex).Invent.Object(6).Amount = 20

UserList(UserIndex).Invent.Object(7).ObjIndex = 559
UserList(UserIndex).Invent.Object(7).Amount = 20

If UserRaza = "Humano" Or UserRaza = "Elfo" Or UserRaza = "Elfo Oscuro" Or UserRaza = "Orco" Then
    If UserClase = "Mago" Or UserClase = "Clerigo" Or _
    UserClase = "Druida" Or UserClase = "Bardo" Then
        UserList(UserIndex).Invent.Object(8).ObjIndex = 571
        UserList(UserIndex).Invent.Object(8).Amount = 1
    End If
        If UserClase = "Guerrero" Or UserClase = "Asesino" Or _
        UserClase = "Ladron" Or UserClase = "Bandido" Or UserClase = "Paladin" Or UserClase = "Cazador" Then
            UserList(UserIndex).Invent.Object(8).ObjIndex = 569
            UserList(UserIndex).Invent.Object(8).Amount = 1
            UserList(UserIndex).Invent.Object(9).ObjIndex = 572
            UserList(UserIndex).Invent.Object(9).Amount = 1
        End If
            If UserClase = "Pescador" Then
            UserList(UserIndex).Invent.Object(8).ObjIndex = 574
            UserList(UserIndex).Invent.Object(8).Amount = 1
            UserList(UserIndex).Invent.Object(9).ObjIndex = 576
            UserList(UserIndex).Invent.Object(9).Amount = 1
            ElseIf UserClase = "Leñador" Then
            UserList(UserIndex).Invent.Object(8).ObjIndex = 573
            UserList(UserIndex).Invent.Object(8).Amount = 1
            UserList(UserIndex).Invent.Object(9).ObjIndex = 576
            UserList(UserIndex).Invent.Object(9).Amount = 1
            ElseIf UserClase = "Minero" Then
            UserList(UserIndex).Invent.Object(8).ObjIndex = 575
            UserList(UserIndex).Invent.Object(8).Amount = 1
            UserList(UserIndex).Invent.Object(9).ObjIndex = 576
            UserList(UserIndex).Invent.Object(9).Amount = 1
            End If
Else
    If UserClase = "Guerrero" Or UserClase = "Asesino" Or _
        UserClase = "Ladron" Or UserClase = "Bandido" Or UserClase = "Paladin" Or UserClase = "Cazador" Then
            UserList(UserIndex).Invent.Object(8).ObjIndex = 570
            UserList(UserIndex).Invent.Object(8).Amount = 1
            UserList(UserIndex).Invent.Object(9).ObjIndex = 572
            UserList(UserIndex).Invent.Object(9).Amount = 1
        End If
            If UserClase = "Pescador" Then
            UserList(UserIndex).Invent.Object(8).ObjIndex = 574
            UserList(UserIndex).Invent.Object(8).Amount = 1
            UserList(UserIndex).Invent.Object(9).ObjIndex = 576
            UserList(UserIndex).Invent.Object(9).Amount = 1
            ElseIf UserClase = "Leñador" Then
            UserList(UserIndex).Invent.Object(8).ObjIndex = 573
            UserList(UserIndex).Invent.Object(8).Amount = 1
            UserList(UserIndex).Invent.Object(9).ObjIndex = 576
            UserList(UserIndex).Invent.Object(9).Amount = 1
            ElseIf UserClase = "Minero" Then
            UserList(UserIndex).Invent.Object(8).ObjIndex = 575
            UserList(UserIndex).Invent.Object(8).Amount = 1
            UserList(UserIndex).Invent.Object(9).ObjIndex = 576
            UserList(UserIndex).Invent.Object(9).Amount = 1
            End If
End If

'[KEVIN]

UserList(UserIndex).BancoInvent.NroItems = 0

UserList(UserIndex).BancoInvent.Object(1).Amount = 0
UserList(UserIndex).BancoInvent.Object(1).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(2).Amount = 0
UserList(UserIndex).BancoInvent.Object(2).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(3).Amount = 0
UserList(UserIndex).BancoInvent.Object(3).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(4).Amount = 0
UserList(UserIndex).BancoInvent.Object(4).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(5).Amount = 0
UserList(UserIndex).BancoInvent.Object(5).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(6).Amount = 0
UserList(UserIndex).BancoInvent.Object(6).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(7).Amount = 0
UserList(UserIndex).BancoInvent.Object(7).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(8).Amount = 0
UserList(UserIndex).BancoInvent.Object(8).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(9).Amount = 0
UserList(UserIndex).BancoInvent.Object(9).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(10).Amount = 0
UserList(UserIndex).BancoInvent.Object(10).ObjIndex = 0

UserList(UserIndex).BancoInvent.Object(11).Amount = 0
UserList(UserIndex).BancoInvent.Object(11).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(12).Amount = 0
UserList(UserIndex).BancoInvent.Object(12).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(13).Amount = 0
UserList(UserIndex).BancoInvent.Object(13).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(14).Amount = 0
UserList(UserIndex).BancoInvent.Object(14).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(15).Amount = 0
UserList(UserIndex).BancoInvent.Object(15).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(16).Amount = 0
UserList(UserIndex).BancoInvent.Object(16).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(17).Amount = 0
UserList(UserIndex).BancoInvent.Object(17).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(18).Amount = 0
UserList(UserIndex).BancoInvent.Object(18).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(19).Amount = 0
UserList(UserIndex).BancoInvent.Object(19).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(20).Amount = 0
UserList(UserIndex).BancoInvent.Object(20).ObjIndex = 0

UserList(UserIndex).BancoInvent.Object(21).Amount = 0
UserList(UserIndex).BancoInvent.Object(21).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(22).Amount = 0
UserList(UserIndex).BancoInvent.Object(22).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(23).Amount = 0
UserList(UserIndex).BancoInvent.Object(23).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(24).Amount = 0
UserList(UserIndex).BancoInvent.Object(24).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(25).Amount = 0
UserList(UserIndex).BancoInvent.Object(25).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(26).Amount = 0
UserList(UserIndex).BancoInvent.Object(26).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(27).Amount = 0
UserList(UserIndex).BancoInvent.Object(27).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(28).Amount = 0
UserList(UserIndex).BancoInvent.Object(28).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(29).Amount = 0
UserList(UserIndex).BancoInvent.Object(29).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(30).Amount = 0
UserList(UserIndex).BancoInvent.Object(30).ObjIndex = 0

UserList(UserIndex).BancoInvent.Object(31).Amount = 0
UserList(UserIndex).BancoInvent.Object(31).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(32).Amount = 0
UserList(UserIndex).BancoInvent.Object(32).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(33).Amount = 0
UserList(UserIndex).BancoInvent.Object(33).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(34).Amount = 0
UserList(UserIndex).BancoInvent.Object(34).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(35).Amount = 0
UserList(UserIndex).BancoInvent.Object(35).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(36).Amount = 0
UserList(UserIndex).BancoInvent.Object(36).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(37).Amount = 0
UserList(UserIndex).BancoInvent.Object(37).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(38).Amount = 0
UserList(UserIndex).BancoInvent.Object(38).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(39).Amount = 0
UserList(UserIndex).BancoInvent.Object(39).ObjIndex = 0
UserList(UserIndex).BancoInvent.Object(40).Amount = 0
UserList(UserIndex).BancoInvent.Object(40).ObjIndex = 0

'[/KEVIN]

UserList(UserIndex).Invent.ArmourEqpSlot = 4
UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(4).ObjIndex

UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.Object(4).ObjIndex).WeaponAnim

UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(3).ObjIndex
UserList(UserIndex).Invent.WeaponEqpSlot = 3
'[/KEVIN]



Call SaveUser(UserIndex, CharPath & UCase(Name) & ".chr", True)

Dim LoopX As Integer

For LoopX = 1 To 40
    If UserList(UserIndex).BancoInvent.Object(LoopX).Amount <> 0 Then UserList(UserIndex).BancoInvent.Object(LoopX).Amount = 0
    If UserList(UserIndex).BancoInvent.Object(LoopX).ObjIndex <> 0 Then UserList(UserIndex).BancoInvent.Object(LoopX).ObjIndex = 0
Next LoopX

'Open User
Call ConnectUser(UserIndex, Name, Password)
  
End Sub

Sub CloseSocket(ByVal UserIndex As Integer)
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>

'Call LogTarea("Close Socket")

On Error GoTo errhandler
    

    
    Call aDos.RestarConexion(frmMain.Socket2(UserIndex).PeerAddress)
    
    UserList(UserIndex).ConnID = -1
    GameInputMapArray(UserIndex) = -1
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
    
    '[KEVIN]PA EL MACRO
    Call SendData(ToIndex, UserIndex, 0, "SMT")
    '[/KEVIN]
            
    If UserList(UserIndex).Flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            
            '[KEVIN]******************
            If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
            End If
            
            If UserList(UserIndex).Invent.Anillo2EqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.Anillo2EqpSlot)
            End If
            
            If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
            End If
            '[/KEVIN]***************************
            
            Call CloseUser(UserIndex)
    End If
    
    frmMain.Socket2(UserIndex).Cleanup
    Unload frmMain.Socket2(UserIndex)
    Call ResetUserSlot(UserIndex)
    

Exit Sub

errhandler:
    UserList(UserIndex).ConnID = -1
    GameInputMapArray(UserIndex) = -1
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
'    Unload frmMain.Socket2(UserIndex) OJOOOOOOOOOOOOOOOOO
'    If NumUsers > 0 Then NumUsers = NumUsers - 1
    Call ResetUserSlot(UserIndex)
    
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<< NO TOCAR >>>>>>>>>>>>>>>>>>>>>>
    
End Sub



Sub SendData(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, sndData As String)


On Error Resume Next

Dim Extras As Long
Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer
Dim aux$
Dim dec$
Dim nfile As Integer
Dim ret As Long
Dim Translate As Integer
Dim Encontrados As Integer

sndData = sndData & ENDC



Select Case sndRoute
        
    Case ToMapButIndex
        Call SendMapUsersButIndex(GameInputMapArray(1), SocketHandle(1), sndMap, sndData, sndIndex)
        Exit Sub
        
    Case ToMap
        Call SendMapUsers(GameInputMapArray(1), SocketHandle(1), sndMap, sndData)
        Exit Sub
        
    Case ToIndex
        If UserList(sndIndex).ConnID > -1 Then
             Call send(SocketHandle(sndIndex), sndData, Len(sndData), 0)
        Exit Sub
        End If
   
    Case ToPCArea
        For Y = UserList(sndIndex).Pos.Y - MinYBorder + 1 To UserList(sndIndex).Pos.Y + MinYBorder - 1
            For X = UserList(sndIndex).Pos.X - MinXBorder + 1 To UserList(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).UserIndex > 0 Then
                       If UserList(MapData(sndMap, X, Y).UserIndex).ConnID > -1 Then
                            Call send(SocketHandle(MapData(sndMap, X, Y).UserIndex), sndData, Len(sndData), 0)
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
    
    Case ToNPCArea
        For Y = Npclist(sndIndex).Pos.Y - MinYBorder + 1 To Npclist(sndIndex).Pos.Y + MinYBorder - 1
            For X = Npclist(sndIndex).Pos.X - MinXBorder + 1 To Npclist(sndIndex).Pos.X + MinXBorder - 1
               If InMapBounds(sndMap, X, Y) Then
                    If MapData(sndMap, X, Y).UserIndex > 0 Then
                       If UserList(MapData(sndMap, X, Y).UserIndex).ConnID > -1 Then
                            Call send(SocketHandle(MapData(sndMap, X, Y).UserIndex), sndData, Len(sndData), 0)
                       End If
                    End If
               End If
            Next X
        Next Y
        Exit Sub
        
    Case ToGuildMembers
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) Then
                If UserList(sndIndex).GuildInfo.GuildName = UserList(LoopC).GuildInfo.GuildName Then _
                        'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        Call send(SocketHandle(LoopC), sndData, Len(sndData), 0)
                End If
            End If
        Next LoopC
        Exit Sub
        
    Case ToAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 Then
               If EsDios(UserList(LoopC).Name) Or EsSemiDios(UserList(LoopC).Name) Then
                       'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                        'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                        Call send(SocketHandle(LoopC), sndData, Len(sndData), 0)
               End If
            End If
        Next LoopC
        Exit Sub
        
    Case ToAll
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 Then
                If UserList(LoopC).Flags.UserLogged Then 'Esta logeado como usuario?
                    'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                    'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                    Call send(SocketHandle(LoopC), sndData, Len(sndData), 0)
                End If
            End If
        Next LoopC
        Exit Sub
        
    Case ToAllButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) And (LoopC <> sndIndex) Then
                If UserList(LoopC).Flags.UserLogged Then 'Esta logeado como usuario?
                    'Call AddtoVar(UserList(LoopC).BytesTransmitidosSvr, LenB(sndData), 100000)
                    'frmMain.Socket2(LoopC).Write sndData, Len(sndData)
                    Call send(SocketHandle(LoopC), sndData, Len(sndData), 0)
                End If
            End If
        Next LoopC
        Exit Sub
        
End Select

End Sub
Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
        For X = UserList(Index).Pos.X - MinXBorder + 1 To UserList(Index).Pos.X + MinXBorder - 1

            If MapData(UserList(Index).Pos.Map, X, Y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next X
Next Y
EstaPCarea = False
End Function

Function HayPCarea(Pos As WorldPos) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.Map, X, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
Next Y
HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If MapData(Pos.Map, X, Y).OBJInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next X
Next Y
HayOBJarea = False
End Function

Sub CorregirSkills(ByVal UserIndex As Integer)
Dim k As Integer

For k = 1 To NUMSKILLS
  If UserList(UserIndex).Stats.UserSkills(k) > MAXSKILLPOINTS Then UserList(UserIndex).Stats.UserSkills(k) = MAXSKILLPOINTS
Next

For k = 1 To NUMATRIBUTOS
 If UserList(UserIndex).Stats.UserAtributos(k) > MAXATRIBUTOS Then
    Call SendData(ToIndex, UserIndex, 0, "ERREl personaje tiene atributos invalidos.")
    Exit Sub
 End If
Next k
 
End Sub


Function ValidateChr(ByVal UserIndex As Integer) As Boolean

ValidateChr = UserList(UserIndex).Char.Head <> 0 And _
UserList(UserIndex).Char.Body <> 0 And ValidateSkills(UserIndex)

End Function

Sub ConnectUser(ByVal UserIndex As Integer, Name As String, Password As String)
Dim N As Integer
Dim Slot As Byte

'Reseteamos los FLAGS
UserList(UserIndex).Flags.TimesWalk = 0
UserList(UserIndex).Flags.StartWalk = 0
UserList(UserIndex).Flags.CountSH = 0
UserList(UserIndex).Flags.Escondido = 0
UserList(UserIndex).Flags.TargetNpc = 0
UserList(UserIndex).Flags.TargetNpcTipo = 0
UserList(UserIndex).Flags.TargetObj = 0
UserList(UserIndex).Flags.TargetUser = 0
UserList(UserIndex).Char.FX = 0
UserList(UserIndex).Flags.TimesWalk = 0
UserList(UserIndex).Flags.StartWalk = 0
'[KEVIN]
UserList(UserIndex).Flags.VerSH = False
'[/KEVIN]

'Controlamos no pasar el maximo de usuarios
If NumUsers >= MaxUsers Then
    Call SendData(ToIndex, UserIndex, 0, "ERREl servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If
  
'¿Este IP ya esta conectado?
If AllowMultiLogins = 0 Then
    If CheckForSameIP(UserIndex, frmMain.Socket2(UserIndex).PeerAddress) = True Then
        Call SendData(ToIndex, UserIndex, 0, "ERRNo es posible usar mas de un personaje al mismo tiempo.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
End If

'¿Ya esta conectado el personaje?
If CheckForSameName(UserIndex, Name) = True Then
    Call SendData(ToIndex, UserIndex, 0, "ERRPerdon, un usuario con el mismo nombre se há logoeado.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Existe el personaje?
If FileExist(CharPath & UCase(Name) & ".chr", vbNormal) = False Then
    Call SendData(ToIndex, UserIndex, 0, "ERREl personaje no existe.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'¿Es el passwd valido?
If UCase(Password) <> UCase(GetVar(CharPath & UCase(Name) & ".chr", "INIT", "Password")) Then
    Call SendData(ToIndex, UserIndex, 0, "ERRPassword incorrecto.")
    'Call frmMain.Socket2(UserIndex).Disconnect
    Call CloseSocket(UserIndex)
    Exit Sub
End If

'Cargamos los datos del personaje
Call LoadUserInit(UserIndex, CharPath & UCase(Name) & ".chr")
Call LoadUserStats(UserIndex, CharPath & UCase(Name) & ".chr")
'Call CorregirSkills(UserIndex)

Call CalculaModificadoresClase(UserIndex)
Call CalculaNivelModificador(UserIndex)

If Not ValidateChr(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "ERRError en el personaje.")
    Call CloseSocket(UserIndex)
    Exit Sub
End If

Call LoadUserReputacion(UserIndex, CharPath & UCase(Name) & ".chr")


If UserList(UserIndex).Invent.EscudoEqpSlot = 0 Then UserList(UserIndex).Char.ShieldAnim = NingunEscudo
If UserList(UserIndex).Invent.CascoEqpSlot = 0 Then UserList(UserIndex).Char.CascoAnim = NingunCasco
If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then UserList(UserIndex).Char.WeaponAnim = NingunArma

'[KEVIN] ******************************************************
'TUVE QUE SCARLO POR UN ERROR KE SALTABA Y NO TE DEJABA CONECTAR
'If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
    'Select Case ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).TipoAnillo
                    
            'Case 1
                'Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Agilidad), ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MaxModificador, MAXATRIBUTOS)
                        
            'Case 2
                'Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Fuerza), ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MaxModificador, MAXATRIBUTOS)
                        
            'Case 3
                'Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Carisma), ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MaxModificador, MAXATRIBUTOS)
                
            'Case 5
                'Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MaxModificador, STAT_MAXHIT)
                'Call AddtoVar(UserList(UserIndex).Stats.MinHIT, ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MaxModificador, STAT_MAXHIT)
                        
    'End Select
'End If

'If UserList(UserIndex).Invent.Anillo2EqpObjIndex > 0 Then
    'Select Case ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).TipoAnillo
                    
            'Case 1
                'Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Agilidad), ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).MaxModificador, MAXATRIBUTOS)
                        
            'Case 2
                'Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Fuerza), ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).MaxModificador, MAXATRIBUTOS)
                        
            'Case 3
                'Call AddtoVar(UserList(UserIndex).Stats.UserAtributos(Carisma), ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).MaxModificador, MAXATRIBUTOS)
                
            'Case 5
                'Call AddtoVar(UserList(UserIndex).Stats.MaxHIT, ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).MaxModificador, STAT_MAXHIT)
                'Call AddtoVar(UserList(UserIndex).Stats.MinHIT, ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).MaxModificador, STAT_MAXHIT)
                        
    'End Select
'End If
'[/KEVIN]***************************************************************

Call UpdateUserInv(True, UserIndex, 0)
Call UpdateUserHechizos(True, UserIndex, 0)

'[KEVIN]
If UserList(UserIndex).Flags.Navegando = 1 Then
    Slot = UserList(UserIndex).Invent.BarcoSlot
    If Slot = 0 Then
        UserList(UserIndex).Flags.Navegando = 0
        If UCase(UserList(UserIndex).Hogar) = "NIX" Then
            UserList(UserIndex).Pos = Nix
        ElseIf UCase(UserList(UserIndex).Hogar) = "DAHL" Then
            UserList(UserIndex).Pos = Dahl
        ElseIf UCase(UserList(UserIndex).Hogar) = "BANDERBILL" Then
            UserList(UserIndex).Pos = Banderbill
        ElseIf UCase(UserList(UserIndex).Hogar) = "TERMINUS" Then
            UserList(UserIndex).Pos = Terminus
        ElseIf UCase(UserList(UserIndex).Hogar) = "NUEVAESPERANZA" Then
            UserList(UserIndex).Pos = NuevaEsperanza
        ElseIf UCase(UserList(UserIndex).Hogar) = "NEXON" Then
            UserList(UserIndex).Pos = Nexon
        Else
            UserList(UserIndex).Hogar = "BANDERBILL"
            UserList(UserIndex).Pos = Banderbill
        End If
    Else
       If UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then
        UserList(UserIndex).Flags.Navegando = 0
        If UCase(UserList(UserIndex).Hogar) = "NIX" Then
            UserList(UserIndex).Pos = Nix
        ElseIf UCase(UserList(UserIndex).Hogar) = "DAHL" Then
            UserList(UserIndex).Pos = Dahl
        ElseIf UCase(UserList(UserIndex).Hogar) = "BANDERBILL" Then
            UserList(UserIndex).Pos = Banderbill
        ElseIf UCase(UserList(UserIndex).Hogar) = "TERMINUS" Then
            UserList(UserIndex).Pos = Terminus
        ElseIf UCase(UserList(UserIndex).Hogar) = "NUEVAESPERANZA" Then
            UserList(UserIndex).Pos = NuevaEsperanza
        ElseIf UCase(UserList(UserIndex).Hogar) = "NEXON" Then
            UserList(UserIndex).Pos = Nexon
        Else
            UserList(UserIndex).Hogar = "BANDERBILL"
            UserList(UserIndex).Pos = Banderbill
        End If
       End If
    End If
    '[/KEVIN]
    
    If UserList(UserIndex).Flags.Navegando = 1 Then
        UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.BarcoObjIndex).Ropaje
        UserList(UserIndex).Char.Head = 0
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    End If
End If

'[KEVIN]ESTO LO SAQUE PARA KE SI SALEN PARALIZADOS ENTREN PARALIZADOS
If UserList(UserIndex).Flags.Paralizado = 1 Then
UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
Call SendData(ToIndex, UserIndex, 0, "PARADOK")
End If
'[/KEVIN]


'Posicion de comienzo
If UserList(UserIndex).Pos.Map = 0 Then
    If UCase(UserList(UserIndex).Hogar) = "NIX" Then
            UserList(UserIndex).Pos = Nix
        ElseIf UCase(UserList(UserIndex).Hogar) = "DAHL" Then
            UserList(UserIndex).Pos = Dahl
        ElseIf UCase(UserList(UserIndex).Hogar) = "BANDERBILL" Then
            UserList(UserIndex).Pos = Banderbill
        ElseIf UCase(UserList(UserIndex).Hogar) = "TERMINUS" Then
            UserList(UserIndex).Pos = Terminus
        ElseIf UCase(UserList(UserIndex).Hogar) = "NUEVAESPERANZA" Then
            UserList(UserIndex).Pos = NuevaEsperanza
        ElseIf UCase(UserList(UserIndex).Hogar) = "NEXON" Then
            UserList(UserIndex).Pos = Nexon
        Else
            UserList(UserIndex).Hogar = "BANDERBILL"
            UserList(UserIndex).Pos = Banderbill
        End If
Else
   
   If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex <> 0 Then Call CloseSocket(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex)
End If

'Upadte anti cuelgues
GameInputMapArray(UserIndex) = UserList(UserIndex).Pos.Map

'Nombre de sistema
UserList(UserIndex).Name = Name

UserList(UserIndex).Password = Password
UserList(UserIndex).ip = frmMain.Socket2(UserIndex).PeerAddress
  
'Info
Call SendData(ToIndex, UserIndex, 0, "IU" & UserIndex) 'Enviamos el User index
Call SendData(ToIndex, UserIndex, 0, "CM" & UserList(UserIndex).Pos.Map & "," & MapInfo(UserList(UserIndex).Pos.Map).MapVersion) 'Carga el mapa
Call SendData(ToIndex, UserIndex, 0, "TM" & MapInfo(UserList(UserIndex).Pos.Map).Music)

If Lloviendo Then Call SendData(ToIndex, UserIndex, 0, "LLU")

Call SendData(ToIndex, UserIndex, 0, "||¡¡¡Bienvenido a CUCSIFAE AO!!!. Si necesitas ayuda escribí /ayuda." & FONTTYPE_INFO)
Call SendData(ToIndex, UserIndex, 0, "||Nuestra página es http://www.cucsifae-ao.com.ar" & FONTTYPE_INFO)

Call UpdateUserMap(UserIndex)
Call SendUserStatsBox(UserIndex)
Call EnviarHambreYsed(UserIndex)

If haciendoBK Then
    Call SendData(ToIndex, UserIndex, 0, "BKW")
    Call SendData(ToIndex, UserIndex, 0, "ERRPor favor espera algunos segundos, WorldSave esta ejecutandose.")
End If

'Actualiza el Num de usuarios
If UserIndex > LastUser Then LastUser = UserIndex

NumUsers = NumUsers + 1


MapInfo(UserList(UserIndex).Pos.Map).NumUsers = MapInfo(UserList(UserIndex).Pos.Map).NumUsers + 1

If UserList(UserIndex).Stats.SkillPts > 0 Then
    Call EnviarSkills(UserIndex)
    Call EnviarSubirNivel(UserIndex, UserList(UserIndex).Stats.SkillPts)
End If

If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers

If NumUsers > recordusuarios Then
    Call SendData(ToAll, 0, 0, "||Record de usuarios conectados simultaniamente." & "Hay " & NumUsers & " usuarios." & FONTTYPE_INFO)
    recordusuarios = NumUsers
    Call WriteVar(IniPath & "Server.ini", "INIT", "Record", Str(recordusuarios))
End If

If EsDios(Name) Then
    UserList(UserIndex).Flags.Privilegios = 3
    Call LogGM(UserList(UserIndex).Name, "Se conecto con ip:" & UserList(UserIndex).ip)
ElseIf EsSemiDios(Name) Then
    UserList(UserIndex).Flags.Privilegios = 2
    Call LogGM(UserList(UserIndex).Name, "Se conecto con ip:" & UserList(UserIndex).ip)
ElseIf EsConsejero(Name) Then
    UserList(UserIndex).Flags.Privilegios = 1
    Call LogGM(UserList(UserIndex).Name, "Se conecto con ip:" & UserList(UserIndex).ip)
Else
    UserList(UserIndex).Flags.Privilegios = 0
End If

Set UserList(UserIndex).GuildRef = FetchGuild(UserList(UserIndex).GuildInfo.GuildName)

UserList(UserIndex).Counters.IdleCount = 0

'BUGREPAIR SOLO PARA ARREGLAR EL BUG DE LAS MASCOTAS
Call RepararMascotas(UserIndex)

If UserList(UserIndex).NroMacotas > 0 Then
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(i) > 0 Then
            UserList(UserIndex).MascotasIndex(i) = SpawnNpc(UserList(UserIndex).MascotasType(i), UserList(UserIndex).Pos, True, True)
            
            If UserList(UserIndex).MascotasIndex(i) <= MAXNPCS Then
                  Npclist(UserList(UserIndex).MascotasIndex(i)).MaestroUser = UserIndex
                  Call FollowAmo(UserList(UserIndex).MascotasIndex(i))
            Else
                  UserList(UserIndex).MascotasIndex(i) = 0
            End If
        End If
    Next i
End If


If UserList(UserIndex).Flags.Navegando = 1 Then Call SendData(ToIndex, UserIndex, 0, "NAVEG")

UserList(UserIndex).Flags.Seguro = True

'Crea  el personaje del usuario
Call MakeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

'[KEVIN]
Call RefreshUserListBox
Call LogInUser(UserIndex)
'[/KEVIN]

Call SendData(ToIndex, UserIndex, 0, "IP" & UserList(UserIndex).Char.CharIndex)
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXWARP & "," & 0)
Call SendData(ToIndex, UserIndex, 0, "LOGGED")
UserList(UserIndex).Flags.UserLogged = True

Call SendGuildNews(UserIndex)


Call MostrarNumUsers

N = FreeFile
Open App.Path & "\logs\numusers.txt" For Output As N
Print #N, NumUsers
Close #N

N = FreeFile
'Log
Open App.Path & "\logs\Connect.log" For Append Shared As #N
Print #N, UserList(UserIndex).Name & " ha entrado al juego. UserIndex:" & UserIndex & " " & Time & " " & Date
Close #N

End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)

UserList(UserIndex).Faccion.ArmadaReal = 0
UserList(UserIndex).Faccion.FuerzasCaos = 0
UserList(UserIndex).Faccion.CiudadanosMatados = 0
UserList(UserIndex).Faccion.CriminalesMatados = 0
UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0
UserList(UserIndex).Faccion.RecibioArmaduraReal = 0
UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0
UserList(UserIndex).Faccion.RecibioExpInicialReal = 0
UserList(UserIndex).Faccion.RecompensasCaos = 0
UserList(UserIndex).Faccion.RecompensasReal = 0

End Sub

Sub ResetContadores(ByVal UserIndex As Integer)

UserList(UserIndex).Counters.AGUACounter = 0
UserList(UserIndex).Counters.AttackCounter = 0
UserList(UserIndex).Counters.Ceguera = 0
UserList(UserIndex).Counters.COMCounter = 0
UserList(UserIndex).Counters.Estupidez = 0
UserList(UserIndex).Counters.Frio = 0
UserList(UserIndex).Counters.HPCounter = 0
UserList(UserIndex).Counters.IdleCount = 0
UserList(UserIndex).Counters.Invisibilidad = 0
UserList(UserIndex).Counters.Paralisis = 0
UserList(UserIndex).Counters.Pasos = 0
UserList(UserIndex).Counters.Pena = 0
UserList(UserIndex).Counters.PiqueteC = 0
UserList(UserIndex).Counters.STACounter = 0
UserList(UserIndex).Counters.Veneno = 0

End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)

UserList(UserIndex).Char.Body = 0
UserList(UserIndex).Char.CascoAnim = 0
UserList(UserIndex).Char.CharIndex = 0
UserList(UserIndex).Char.FX = 0
UserList(UserIndex).Char.Head = 0
UserList(UserIndex).Char.loops = 0
UserList(UserIndex).Char.Heading = 0
UserList(UserIndex).Char.loops = 0
UserList(UserIndex).Char.ShieldAnim = 0
UserList(UserIndex).Char.WeaponAnim = 0

End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)

UserList(UserIndex).Name = ""
UserList(UserIndex).modName = ""
UserList(UserIndex).Password = ""
UserList(UserIndex).Desc = ""
UserList(UserIndex).Pos.Map = 0
UserList(UserIndex).Pos.X = 0
UserList(UserIndex).Pos.Y = 0
UserList(UserIndex).ip = ""
UserList(UserIndex).RDBuffer = ""
UserList(UserIndex).Clase = ""
UserList(UserIndex).Email = ""
UserList(UserIndex).Genero = ""
UserList(UserIndex).Hogar = ""
UserList(UserIndex).Raza = ""

UserList(UserIndex).RandKey = 0
UserList(UserIndex).PrevCRC = 0
UserList(UserIndex).PacketNumber = 0

UserList(UserIndex).Stats.Banco = 0
UserList(UserIndex).Stats.ELV = 0
UserList(UserIndex).Stats.ELU = 0
UserList(UserIndex).Stats.Exp = 0
UserList(UserIndex).Stats.Def = 0
UserList(UserIndex).Stats.CriminalesMatados = 0
UserList(UserIndex).Stats.NPCsMuertos = 0
UserList(UserIndex).Stats.UsuariosMatados = 0

End Sub

Sub ResetReputacion(ByVal UserIndex As Integer)

UserList(UserIndex).Reputacion.AsesinoRep = 0
UserList(UserIndex).Reputacion.BandidoRep = 0
UserList(UserIndex).Reputacion.BurguesRep = 0
UserList(UserIndex).Reputacion.LadronesRep = 0
UserList(UserIndex).Reputacion.NobleRep = 0
UserList(UserIndex).Reputacion.PlebeRep = 0
UserList(UserIndex).Reputacion.NobleRep = 0
UserList(UserIndex).Reputacion.Promedio = 0

End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)

UserList(UserIndex).GuildInfo.ClanFundado = ""
UserList(UserIndex).GuildInfo.Echadas = 0
UserList(UserIndex).GuildInfo.EsGuildLeader = 0
UserList(UserIndex).GuildInfo.FundoClan = 0
UserList(UserIndex).GuildInfo.GuildName = ""
UserList(UserIndex).GuildInfo.Solicitudes = 0
UserList(UserIndex).GuildInfo.SolicitudesRechazadas = 0
UserList(UserIndex).GuildInfo.VecesFueGuildLeader = 0
UserList(UserIndex).GuildInfo.YaVoto = 0
UserList(UserIndex).GuildInfo.ClanesParticipo = 0
UserList(UserIndex).GuildInfo.GuildPoints = 0

End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)

UserList(UserIndex).Flags.Comerciando = False
UserList(UserIndex).Flags.Ban = 0
UserList(UserIndex).Flags.Escondido = 0
UserList(UserIndex).Flags.DuracionEfecto = 0
UserList(UserIndex).Flags.NpcInv = 0
UserList(UserIndex).Flags.StatsChanged = 0
UserList(UserIndex).Flags.TargetNpc = 0
UserList(UserIndex).Flags.TargetNpcTipo = 0
UserList(UserIndex).Flags.TargetObj = 0
UserList(UserIndex).Flags.TargetObjMap = 0
UserList(UserIndex).Flags.TargetObjX = 0
UserList(UserIndex).Flags.TargetObjY = 0
UserList(UserIndex).Flags.TargetUser = 0
UserList(UserIndex).Flags.TipoPocion = 0
UserList(UserIndex).Flags.TomoPocion = False
UserList(UserIndex).Flags.Descuento = ""
UserList(UserIndex).Flags.Hambre = 0
UserList(UserIndex).Flags.Sed = 0
UserList(UserIndex).Flags.Descansar = False
UserList(UserIndex).Flags.ModoCombate = False
UserList(UserIndex).Flags.Vuela = 0
UserList(UserIndex).Flags.Navegando = 0
UserList(UserIndex).Flags.Oculto = 0
UserList(UserIndex).Flags.Envenenado = 0
UserList(UserIndex).Flags.Invisible = 0
UserList(UserIndex).Flags.Paralizado = 0
UserList(UserIndex).Flags.Maldicion = 0
UserList(UserIndex).Flags.Bendicion = 0
UserList(UserIndex).Flags.Meditando = 0
UserList(UserIndex).Flags.Privilegios = 0
UserList(UserIndex).Flags.PuedeMoverse = 0
UserList(UserIndex).Flags.PuedeLanzarSpell = 0
UserList(UserIndex).Stats.SkillPts = 0
UserList(UserIndex).Flags.OldBody = 0
UserList(UserIndex).Flags.OldHead = 0
UserList(UserIndex).Flags.AdminInvisible = 0
UserList(UserIndex).Flags.ValCoDe = 0
UserList(UserIndex).Flags.Hechizo = 0

End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)

Dim LoopC As Integer
For LoopC = 1 To MAXUSERHECHIZOS
    UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
Next

End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)

Dim LoopC As Integer

UserList(UserIndex).NroMacotas = 0
    
For LoopC = 1 To MAXMASCOTAS
    UserList(UserIndex).MascotasIndex(LoopC) = 0
    UserList(UserIndex).MascotasType(LoopC) = 0
Next LoopC

End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)

Set UserList(UserIndex).CommandsBuffer = Nothing
Set UserList(UserIndex).GuildRef = Nothing

Call ResetFacciones(UserIndex)
Call ResetContadores(UserIndex)
Call ResetCharInfo(UserIndex)
Call ResetBasicUserInfo(UserIndex)
Call ResetReputacion(UserIndex)
Call ResetGuildInfo(UserIndex)
Call ResetUserFlags(UserIndex)
Call LimpiarInventario(UserIndex)
Call ResetUserSpells(UserIndex)
Call ResetUserPets(UserIndex)

'UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
'UserList(UserIndex).BytesTransmitidosUser = 0
'UserList(UserIndex).BytesTransmitidosSvr = 0





End Sub
Sub CloseUser(ByVal UserIndex As Integer)
'Call LogTarea("CloseUser " & UserIndex)

On Error GoTo errhandler

Dim N As Integer
Dim X As Integer
Dim Y As Integer
Dim LoopC As Integer
Dim Map As Integer
Dim Name As String
Dim Raza As String
Dim Clase As String
Dim i As Integer

Dim aN As Integer

aN = UserList(UserIndex).Flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).Flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).Flags.OldHostil
      Npclist(aN).Flags.AttackedBy = ""
End If

Map = UserList(UserIndex).Pos.Map
X = UserList(UserIndex).Pos.X
Y = UserList(UserIndex).Pos.Y
Name = UCase(UserList(UserIndex).Name)
Raza = UserList(UserIndex).Raza
Clase = UserList(UserIndex).Clase

UserList(UserIndex).Char.FX = 0
UserList(UserIndex).Char.loops = 0
Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
   

UserList(UserIndex).Flags.UserLogged = False

'Le devolvemos el body y head originales
If UserList(UserIndex).Flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)

' Grabamos el personaje del usuario
Call SaveUser(UserIndex, CharPath & Name & ".chr")

'Quitar el dialogo
If MapInfo(Map).NumUsers > 0 Then
    Call SendData(ToMapButIndex, UserIndex, Map, "QDL" & UserList(UserIndex).Char.CharIndex)
End If

'Borrar el personaje
If UserList(UserIndex).Char.CharIndex > 0 Then
    Call EraseUserChar(ToMapButIndex, UserIndex, Map, UserIndex)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).Flags.NPCActive Then _
                Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
    End If
Next i

If UserIndex = LastUser Then
    Do Until UserList(LastUser).Flags.UserLogged
        LastUser = LastUser - 1
        If LastUser < 1 Then Exit Do
    Loop
End If
  
'If NumUsers <> 0 Then
'    NumUsers = NumUsers - 1
'End If

'Update Map Users
MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1

If MapInfo(Map).NumUsers < 0 Then
    MapInfo(Map).NumUsers = 0
End If

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(UserIndex).Name) Then Call Ayuda.Quitar(UserList(UserIndex).Name)

'[KEVIN]
If InscTorneo.Existe(UserList(UserIndex).Name) Then Call InscTorneo.Quitar(UserList(UserIndex).Name)
'[/KEVIN]

Call ResetUserSlot(UserIndex)

Call MostrarNumUsers
'[KEVIN]
Call RefreshUserListBox
Call LogOutUser
'[/KEVIN]

N = FreeFile(1)
Open App.Path & "\logs\Connect.log" For Append Shared As #N
Print #N, Name & " há dejado el juego. " & "User Index:" & UserIndex & " " & Time & " " & Date
Close #N

Exit Sub

errhandler:
Call LogError("Error en CloseUser")


End Sub


Sub HandleData(ByVal UserIndex As Integer, ByVal rdata As String)

'Call LogTarea("Sub HandleData :" & Rdata & " " & UserList(UserIndex).name)

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
Dim pass As String
Dim mapa As Integer
Dim Name As String
Dim ind
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim TempTick As Long
Dim ClientCRC As String
Dim ServerSideCRC As Long

CadenaOriginal = rdata

'¿Tiene un indece valido?
If UserIndex <= 0 Then
    Call CloseSocket(UserIndex)
    Exit Sub
End If

If Left$(rdata, 13) = "gIvEmEvAlcOde" Then
   '<<<<<<<<<<< MODULO PRIVADO DE CADA IMPLEMENTACION >>>>>>
   UserList(UserIndex).Flags.ValCoDe = CInt(RandomNumber(20000, 32000))
   UserList(UserIndex).RandKey = CLng(RandomNumber(0, 99999))
   UserList(UserIndex).PrevCRC = UserList(UserIndex).RandKey
   UserList(UserIndex).PacketNumber = 100
   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   Call SendData(ToIndex, UserIndex, 0, "VAL" & UserList(UserIndex).RandKey & "," & UserList(UserIndex).Flags.ValCoDe)
   Exit Sub
Else
   '<<<<<<<<<<< MODULO PRIVADO DE CADA IMPLEMENTACION >>>>>>
   ClientCRC = ReadField(2, rdata, 126)
   tStr = Left$(rdata, Len(rdata) - Len(ClientCRC) - 1)
   ServerSideCRC = GenCrc(UserList(UserIndex).PrevCRC, tStr)
   If CLng(ClientCRC) <> ServerSideCRC Then Call CloseSocket(UserIndex)
   UserList(UserIndex).PrevCRC = ServerSideCRC
   rdata = tStr
   tStr = ""
   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>
End If

If UserList(UserIndex).Flags.UserLogged Then UserList(UserIndex).Counters.IdleCount = 0
   
   If Not UserList(UserIndex).Flags.UserLogged Then

        Select Case Left$(rdata, 6)
            Case "OLOGIN"
                rdata = Right$(rdata, Len(rdata) - 6)
                Ver = ReadField(3, rdata, 44)
                If VersionOK(Ver) Then
                    tName = ReadField(1, rdata, 44)
                    
                    If Not AsciiValidos(tName) Then
                        Call SendData(ToIndex, UserIndex, 0, "ERRNombre invalido.")
                        Exit Sub
                    End If
                    
                    If Not PersonajeExiste(tName) Then
                        Call SendData(ToIndex, UserIndex, 0, "ERREl personaje no existe.")
                        Exit Sub
                    End If
                    
                    If Not BANCheck(tName) Then
                    
                        '[KEVIN]
                        If ReadField(4, rdata, 44) <> "ANTICHEAT" Then
                            Call LogHackAttemp("IP:" & frmMain.Socket2(UserIndex).PeerAddress & ", " & UserList(UserIndex).Name & ": intentó entrar con un cliente trucho.")
                            Call SendData(ToIndex, UserIndex, 0, "ERREse no es nuestro cliente, has sido anotado en nuestros logs por intentar conectarte con otro cliente.")
                            Call CloseSocket(UserIndex)
                            Exit Sub
                        End If
                        '[/KEVIN]
                        
                        If (UserList(UserIndex).Flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(UserIndex).Flags.ValCoDe) <> CInt(val(ReadField(5, rdata, 44)))) Then
                              Call LogHackAttemp("IP:" & frmMain.Socket2(UserIndex).PeerAddress & " intento crear un bot.")
                              Call CloseSocket(UserIndex)
                              Exit Sub
                        End If
                        
                        Call ConnectUser(UserIndex, tName, ReadField(2, rdata, 44))
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "ERRSe te ha prohibido la entrada a Argentum debido a tu mal comportamiento.")
                    End If
                    
                Else
                     Call SendData(ToIndex, UserIndex, 0, "ERREsta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en nuestra pagina. http://www.cucsifae-ao.com.ar")
                     'Call CloseSocket(UserIndex)
                     Exit Sub
                End If
                Exit Sub
            Case "NLOGIN"
            
                If aClon.MaxPersonajes(frmMain.Socket2(UserIndex).PeerAddress) Then
                        Call SendData(ToIndex, UserIndex, 0, "ERRHas creado demasiados personajes.")
                        Call CloseSocket(UserIndex)
                        Exit Sub
                End If
                
                rdata = Right$(rdata, Len(rdata) - 6)
'                If Not ValidInputNP(rdata) Then Exit Sub
                
                Ver = ReadField(5, rdata, 44)
                If VersionOK(Ver) Then
                     Dim miinteger As Integer
                     miinteger = CInt(val(ReadField(40, rdata, 44)))
                     
                     If (UserList(UserIndex).Flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(UserIndex).Flags.ValCoDe) <> CInt(val(ReadField(40, rdata, 44)))) Then
                         Call LogHackAttemp("IP:" & frmMain.Socket2(UserIndex).PeerAddress & " intento crear un bot.")
                         Call CloseSocket(UserIndex)
                         Exit Sub
                     End If
                     
                     Call ConnectNewUser(UserIndex, ReadField(1, rdata, 44), ReadField(2, rdata, 44), val(ReadField(3, rdata, 44)), ReadField(4, rdata, 44), ReadField(6, rdata, 44), ReadField(7, rdata, 44), _
                     ReadField(8, rdata, 44), ReadField(9, rdata, 44), ReadField(10, rdata, 44), ReadField(11, rdata, 44), ReadField(12, rdata, 44), ReadField(13, rdata, 44), _
                     ReadField(14, rdata, 44), ReadField(15, rdata, 44), ReadField(16, rdata, 44), ReadField(17, rdata, 44), ReadField(18, rdata, 44), ReadField(19, rdata, 44), _
                     ReadField(20, rdata, 44), ReadField(21, rdata, 44), ReadField(22, rdata, 44), ReadField(23, rdata, 44), ReadField(24, rdata, 44), ReadField(25, rdata, 44), _
                     ReadField(26, rdata, 44), ReadField(27, rdata, 44), ReadField(28, rdata, 44), ReadField(29, rdata, 44), ReadField(30, rdata, 44), ReadField(31, rdata, 44), _
                     ReadField(32, rdata, 44), ReadField(33, rdata, 44), ReadField(34, rdata, 44), ReadField(35, rdata, 44), ReadField(36, rdata, 44), ReadField(37, rdata, 44), _
                     ReadField(38, rdata, 44), ReadField(39, rdata, 44))
                Else
                     Call SendData(ToIndex, UserIndex, 0, "!!Esta version del juego es obsoleta, la version correcta es " & ULTIMAVERSION & ". La misma se encuentra disponible en nuestra pagina.")
                     Exit Sub
                End If
                
                Exit Sub
        End Select
    End If
 
Select Case Left$(rdata, 4)
    Case "BORR" ' <<< borra personajes
       On Error GoTo ExitErr1
        rdata = Right$(rdata, Len(rdata) - 4)
        If (UserList(UserIndex).Flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(UserIndex).Flags.ValCoDe) <> CInt(val(ReadField(3, rdata, 44)))) Then
                      Call LogHackAttemp("IP:" & frmMain.Socket2(UserIndex).PeerAddress & " intento borrar un personaje.")
                      Call CloseSocket(UserIndex)
                      Exit Sub
        End If
        Arg1 = ReadField(1, rdata, 44)
        
        If Not AsciiValidos(Arg1) Then Exit Sub
        
        '¿Existe el personaje?
        If Not FileExist(CharPath & UCase(Arg1) & ".chr", vbNormal) Then
            Call CloseSocket(UserIndex)
            Exit Sub
        End If

        '¿Es el passwd valido?
        If UCase(ReadField(2, rdata, 44)) <> UCase(GetVar(CharPath & UCase(Arg1) & ".chr", "INIT", "Password")) Then
            Call CloseSocket(UserIndex)
            Exit Sub
        End If

        'If FileExist(CharPath & UCase(Arg1) & ".chr", vbNormal) Then
            Dim rt$
            rt$ = App.Path & "\ChrBackUp\" & UCase(Arg1) & ".bak"
            If FileExist(rt$, vbNormal) Then Kill rt$
            Name CharPath & UCase(Arg1) & ".chr" As rt$
            Call SendData(ToIndex, UserIndex, 0, "BORROK")
            Exit Sub
ExitErr1:
    Call LogError(Err.Description & " " & rdata)
    Exit Sub
        'End If
End Select

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Si no esta logeado y envia un comando diferente a los
'de arriba cerramos la conexion.
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
If Not UserList(UserIndex).Flags.UserLogged Then
    Call LogHackAttemp("Mesaje enviado sin logearse:" & rdata)
    Call CloseSocket(UserIndex)
    Exit Sub
End If
  


Select Case Left$(rdata, 1)
Dim caca As Long
    Case "M" 'Moverse
        If UserList(UserIndex).Flags.TimesWalk >= 30 Then
            TempTick = GetTickCount And &H7FFFFFFF
            caca = (TempTick - UserList(UserIndex).Flags.StartWalk)
            If caca < 6050 Then
                If TempTick - UserList(UserIndex).Flags.CountSH > 90000 Then
                    UserList(UserIndex).Flags.CountSH = 0
                End If
                If Not UserList(UserIndex).Flags.CountSH = 0 Then
                    caca = 126000 / caca
                    Call LogHackAttemp("Tramposo SH: " & UserList(UserIndex).Name & " , " & caca)
                    
                    '[KEVIN]
                    Dim LoopXXX As Integer
                        For LoopXXX = 1 To LastUser
                            If UserList(LoopXXX).Flags.Privilegios > 1 Then
                                If UserList(LoopXXX).Flags.VerSH = True Then
                                    Call SendData(ToIndex, LoopXXX, 0, "||Servidor: Tramposo SH: " & UserList(UserIndex).Name & " , " & caca & FONTTYPE_CELESTE & ENDC)
                                End If
                            End If
                        Next LoopXXX
                    '[/KEVIN]
                        
                    Call CloseSocket(UserIndex)
                    Exit Sub
                Else
                    UserList(UserIndex).Flags.CountSH = TempTick
                End If
            End If
            UserList(UserIndex).Flags.StartWalk = TempTick
            UserList(UserIndex).Flags.TimesWalk = 0
        End If
        UserList(UserIndex).Flags.TimesWalk = UserList(UserIndex).Flags.TimesWalk + 1
        
        rdata = Right$(rdata, Len(rdata) - 1)
        
        If UserList(UserIndex).Flags.Oculto = 1 Then
            If CheckArmaduraOculto(UserIndex) = False Then
                If UCase(UserList(UserIndex).Clase) <> "LADRON" Then
                    If UCase(UserList(UserIndex).Clase) <> "ASESINO" Then
                        Call SendData(ToIndex, UserIndex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
                        UserList(UserIndex).Flags.Oculto = 0
                        UserList(UserIndex).Flags.Invisible = 0
                        Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "KLBPA" & UserList(UserIndex).Char.CharIndex & ",0")
                        'Exit Sub
                    End If
                End If
            End If
        End If
        
        If Not UserList(UserIndex).Flags.Descansar And Not UserList(UserIndex).Flags.Meditando _
           And UserList(UserIndex).Flags.Paralizado = 0 Then
           Call MoveUserChar(UserIndex, val(rdata))
           Exit Sub
        ElseIf UserList(UserIndex).Flags.Descansar Then
          UserList(UserIndex).Flags.Descansar = False
          Call SendData(ToIndex, UserIndex, 0, "DOK")
          Call SendData(ToIndex, UserIndex, 0, "||Has dejado de descansar." & FONTTYPE_INFO)
          Call MoveUserChar(UserIndex, val(rdata))
          Exit Sub
        ElseIf UserList(UserIndex).Flags.Meditando Then
          UserList(UserIndex).Flags.Meditando = False
          Call SendData(ToIndex, UserIndex, 0, "MEDOK")
          Call SendData(ToIndex, UserIndex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
          UserList(UserIndex).Char.FX = 0
          UserList(UserIndex).Char.loops = 0
          Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
          Call MoveUserChar(UserIndex, val(rdata))
          Exit Sub
        Else
          Call SendData(ToIndex, UserIndex, 0, "||No podes moverte porque estas paralizado." & FONTTYPE_INFO)
          UserList(UserIndex).Flags.CountSH = 0
          Exit Sub
        End If

        Exit Sub
        
    Case ";" 'Hablar
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 1)
        If InStr(rdata, "°") Then
            Exit Sub
        End If
    
        ind = UserList(UserIndex).Char.CharIndex
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & rdata & "°" & Str(ind))
        Exit Sub
    
    Case "-" 'Gritar
        If UserList(UserIndex).Flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
                Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 1)
        If InStr(rdata, "°") Then
            Exit Sub
        End If
        ind = UserList(UserIndex).Char.CharIndex
        Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbRed & "°" & rdata & "°" & Str(ind))
        Exit Sub
        
    Case "\" 'Susurrar al oido
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 1)
        tName = ReadField(1, rdata, 32)
        tIndex = NameIndex(tName)
        If tIndex <> 0 Then
            If Len(rdata) <> Len(tName) Then
                tMessage = Right$(rdata, Len(rdata) - (1 + Len(tName)))
            Else
                tMessage = " "
            End If
            If Not EstaPCarea(UserIndex, tIndex) Then
                Call SendData(ToIndex, UserIndex, 0, "||Estas muy lejos del usuario." & FONTTYPE_INFO)
                Exit Sub
            End If
            ind = UserList(UserIndex).Char.CharIndex
            If InStr(tMessage, "°") Then
                Exit Sub
            End If
            Call SendData(ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbBlue & "°" & tMessage & "°" & Str(ind))
            Call SendData(ToIndex, tIndex, UserList(UserIndex).Pos.Map, "||" & vbBlue & "°" & tMessage & "°" & Str(ind))
            Exit Sub
        End If
        Call SendData(ToIndex, UserIndex, 0, "||Usuario inexistente. " & FONTTYPE_INFO)
        Exit Sub
    
    Exit Sub
End Select

Select Case Left$(rdata, 2)

    Case "TI" 'Tirar item
            If UserList(UserIndex).Flags.Navegando = 1 Then Exit Sub
            rdata = Right$(rdata, Len(rdata) - 2)
            Arg1 = ReadField(1, rdata, 44)
            Arg2 = ReadField(2, rdata, 44)
            If val(Arg1) = FLAGORO Then
                Call TirarOro(val(Arg2), UserIndex)
                Call SendUserStatsBox(UserIndex)
                Exit Sub
            Else
                If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) > 0 Then
                    If UserList(UserIndex).Invent.Object(val(Arg1)).ObjIndex = 0 Then
                            Exit Sub
                    End If
                    Call DropObj(UserIndex, val(Arg1), val(Arg2), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                Else
                    Exit Sub
                End If
            End If
            Exit Sub
    
    Case "LK" ' Lanzar hechizo
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!." & FONTTYPE_INFO)
            Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 2)
        UserList(UserIndex).Flags.Hechizo = val(rdata)
        Exit Sub
    
    Case "LC" 'Click izquierdo
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        X = CInt(Arg1)
        Y = CInt(Arg2)
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
        Exit Sub
    
    Case "RC" 'Click derecho
        rdata = Right$(rdata, Len(rdata) - 2)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
        X = CInt(Arg1)
        Y = CInt(Arg2)
        Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
        Exit Sub
    
    Case "UE"
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!." & FONTTYPE_INFO)
            Exit Sub
        End If

        rdata = Right$(rdata, Len(rdata) - 2)
        Select Case val(rdata)
            Case Robar
                Call SendData(ToIndex, UserIndex, 0, "T01" & Robar)
            Case Magia
                Call SendData(ToIndex, UserIndex, 0, "T01" & Magia)
            Case Domar
                Call SendData(ToIndex, UserIndex, 0, "T01" & Domar)
            Case Ocultarse
                
                If UserList(UserIndex).Flags.Navegando = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||No podes ocultarte si estas navegando." & FONTTYPE_INFO)
                      Exit Sub
                End If
                
                If UserList(UserIndex).Flags.Oculto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Ya estas oculto." & FONTTYPE_INFO)
                      Exit Sub
                End If
                
                Call DoOcultarse(UserIndex)
        End Select
        Exit Sub
    
    Case "AG"
        If UserList(UserIndex).Flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden tomar objetos. " & FONTTYPE_INFO)
                Exit Sub
        End If
        Call GetObj(UserIndex)
        Exit Sub
End Select


Select Case Left$(rdata, 3)
    Case "WLC" 'Click izquierdo en modo trabajo
        rdata = Right$(rdata, Len(rdata) - 3)
        Arg1 = ReadField(1, rdata, 44)
        Arg2 = ReadField(2, rdata, 44)
        Arg3 = ReadField(3, rdata, 44)
        If Arg3 = "" Or Arg2 = "" Or Arg1 = "" Then Exit Sub
        If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(Arg3) Then Exit Sub
        
        X = CInt(Arg1)
        Y = CInt(Arg2)
        tLong = CInt(Arg3)
        
        If UserList(UserIndex).Flags.Muerto = 1 Or _
           UserList(UserIndex).Flags.Descansar Or _
           UserList(UserIndex).Flags.Meditando Or _
           Not InMapBounds(UserList(UserIndex).Pos.Map, X, Y) Then Exit Sub
                          
        
        Select Case tLong
        
        Case Proyectiles
            Dim TU As Integer, tN As Integer
            
            If AntiSH = 0 Then
                'TempTick = GetTickCount And &H7FFFFFFF
                'If TempTick - UserList(UserIndex).Flags.LastSpell < IntervaloUserPuedeCastear Then Exit Sub
                'UserList(UserIndex).Flags.LastSpell = TempTick
                TempTick = GetTickCount And &H7FFFFFFF
                If TempTick - UserList(UserIndex).Flags.LastAtacar < IntervaloUserPuedeAtacar Then Exit Sub
                UserList(UserIndex).Flags.LastAtacar = TempTick
            Else
                If UserList(UserIndex).Counters.ProyectilesCounter > 0 Then Exit Sub
                UserList(UserIndex).Counters.ProyectilesCounter = IntervaloUserPuedeProyectiles
            End If
            
            'Nos aseguramos que este usando un arma de proyectiles
            If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then Exit Sub
            
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then Exit Sub
             
            If UserList(UserIndex).Invent.MunicionEqpObjIndex = 0 Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tenes municiones." & FONTTYPE_INFO)
                    Exit Sub
            End If
             
            'Quitamos stamina
            If UserList(UserIndex).Stats.MinSta >= 10 Then
                 Call QuitarSta(UserIndex, RandomNumber(1, 10))
            Else
                 Call SendData(ToIndex, UserIndex, 0, "||Estas muy cansado para luchar." & FONTTYPE_INFO)
                 Exit Sub
            End If
            
            '[KEVIN]
            If UserList(UserIndex).Flags.Oculto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
                UserList(UserIndex).Flags.Oculto = 0
                UserList(UserIndex).Flags.Invisible = 0
                Call SendData(ToMap, 0, UserList(UserIndex).Pos.Map, "KLBPA" & UserList(UserIndex).Char.CharIndex & ",0")
                'Exit Sub
            End If
            '[/KEVIN]
             
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, Arg1, Arg2)
            
            TU = UserList(UserIndex).Flags.TargetUser
            tN = UserList(UserIndex).Flags.TargetNpc
            
            
            If tN > 0 Then
                If Npclist(tN).Attackable = 0 Then Exit Sub
            Else
                If TU = 0 Then Exit Sub
            End If
            
            If tN > 0 Then Call UsuarioAtacaNpc(UserIndex, tN)
                
            If TU > 0 Then
                If UserList(UserIndex).Flags.Seguro Then
                        If Not Criminal(TU) Then
                                Call SendData(ToIndex, UserIndex, 0, "||No podes atacar ciudadanos, para hacerlo debes desactivar el seguro apretando la tecla S" & FONTTYPE_FIGHT)
                                Exit Sub
                        End If
                End If

                Call UsuarioAtacaUsuario(UserIndex, TU)
            End If
            
            Dim DummyInt As Integer
            DummyInt = UserList(UserIndex).Invent.MunicionEqpSlot
            Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 1)
            If DummyInt < 1 Or DummyInt > MAX_INVENTORY_SLOTS Then Exit Sub
            If UserList(UserIndex).Invent.Object(DummyInt).Amount > 0 Then
                UserList(UserIndex).Invent.Object(DummyInt).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpSlot = DummyInt
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(DummyInt).ObjIndex
                Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                Exit Sub
            Else
                Call UpdateUserInv(False, UserIndex, DummyInt)
                UserList(UserIndex).Invent.MunicionEqpSlot = 0
                UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
                Exit Sub
            End If
            
        Case Magia
            If AntiSH = 0 Then
                TempTick = GetTickCount And &H7FFFFFFF
                If TempTick - UserList(UserIndex).Flags.LastSpell < IntervaloUserPuedeCastear Then Exit Sub
                UserList(UserIndex).Flags.LastSpell = TempTick
            Else
                If UserList(UserIndex).Counters.MagiaCounter > 0 Then Exit Sub
                UserList(UserIndex).Counters.MagiaCounter = IntervaloUserPuedeCastear
            End If
            
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, ReadField(1, rdata, 44), ReadField(2, rdata, 44))
            
            If UserList(UserIndex).Flags.Hechizo > 0 Then
                Call LanzarHechizo(UserList(UserIndex).Flags.Hechizo, UserIndex)
                UserList(UserIndex).Flags.Hechizo = 0
            Else
                Call SendData(ToIndex, UserIndex, 0, "||¡Primero selecciona el hechizo que quieres lanzar!" & FONTTYPE_INFO)
                Exit Sub
            End If
        
        Case Pesca
            TempTick = GetTickCount And &H7FFFFFFF
            If TempTick - UserList(UserIndex).Flags.LastWork < IntervaloUserPuedeTrabajar Then Exit Sub
            UserList(UserIndex).Flags.LastWork = TempTick
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
            
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex = OBJTYPE_CAÑA Or UserList(UserIndex).Invent.HerramientaEqpObjIndex = OBJTYPE_CAÑAN Then '<-----[KEVIN]
                                   
            If HayAgua(UserList(UserIndex).Pos.Map, X, Y) Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_PESCAR)
                Call DoPescar(UserIndex)
                Exit Sub
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No hay agua donde pescar busca un lago, rio o mar." & FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
            
        '[KEVIN](RED DE PESCA)
        Case PescarR
            TempTick = GetTickCount And &H7FFFFFFF
            If TempTick - UserList(UserIndex).Flags.LastWork < IntervaloUserPuedeTrabajar Then Exit Sub
            UserList(UserIndex).Flags.LastWork = TempTick
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> RED_PESCA Then
                    Call CloseSocket(UserIndex)
                    Exit Sub
            End If
                      
            If HayAgua(UserList(UserIndex).Pos.Map, X, Y) Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_PESCAR)
                Call DoPescarR(UserIndex)
                Exit Sub
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No hay agua donde pescar busca un lago, rio o mar." & FONTTYPE_INFO)
                Exit Sub
            End If
        '[/KEVIN]
        Case Robar
            TempTick = GetTickCount And &H7FFFFFFF
            If TempTick - UserList(UserIndex).Flags.LastWork < IntervaloUserPuedeTrabajar Then Exit Sub
            UserList(UserIndex).Flags.LastWork = TempTick
           
           If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                If UserList(UserIndex).Flags.TargetUser > 0 And UserList(UserIndex).Flags.TargetUser <> UserIndex Then
                   If UserList(UserList(UserIndex).Flags.TargetUser).Flags.Muerto = 0 Then
                        wpaux.Map = UserList(UserIndex).Pos.Map
                        wpaux.X = val(ReadField(1, rdata, 44))
                        wpaux.Y = val(ReadField(2, rdata, 44))
                        If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                            Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        Call DoRobar(UserIndex, UserList(UserIndex).Flags.TargetUser)
                        Exit Sub
                   End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||No a quien robarle!." & FONTTYPE_INFO)
                    Exit Sub
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||¡No podes robarle en zonas seguras!." & FONTTYPE_INFO)
                Exit Sub
            End If
            
        Case Talar
            TempTick = GetTickCount And &H7FFFFFFF
            If TempTick - UserList(UserIndex).Flags.LastWork < IntervaloUserPuedeTrabajar Then Exit Sub
            UserList(UserIndex).Flags.LastWork = TempTick
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Deberías equiparte el hacha." & FONTTYPE_INFO)
                Exit Sub
            End If
            
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex = HACHA_LEÑADOR Or UserList(UserIndex).Invent.HerramientaEqpObjIndex = HACHA_LEÑADORN Then '<---[KEVIN]
                    
            auxind = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
            If auxind > 0 Then
                wpaux.Map = UserList(UserIndex).Pos.Map
                wpaux.X = X
                wpaux.Y = Y
                If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                    Exit Sub
                End If
                '¿Hay un arbol donde clickeo?
                If ObjData(auxind).ObjType = OBJTYPE_ARBOLES Then
                    Call SendData(ToPCArea, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SOUND_TALAR)
                    Call DoTalar(UserIndex)
                    Exit Sub
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No hay ningun arbol ahi." & FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        '[KEVIN]
        Case Jardineria
            TempTick = GetTickCount And &H7FFFFFFF
            If TempTick - UserList(UserIndex).Flags.LastWork < IntervaloUserPuedeTrabajar Then Exit Sub
            UserList(UserIndex).Flags.LastWork = TempTick
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Deberías equiparte las tijeras." & FONTTYPE_INFO)
                Exit Sub
            End If
            
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex = TIJERAS_DRUIDA Then
                    
            auxind = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
            If auxind > 0 Then
                wpaux.Map = UserList(UserIndex).Pos.Map
                wpaux.X = X
                wpaux.Y = Y
                If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                    Exit Sub
                End If
                '¿Hay un arbol donde clickeo?
                If ObjData(auxind).ObjType = OBJTYPE_ARBOLES Then
                    'Call SendData(ToPCArea, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SOUND_TALAR)
                    Call DoJardineria(UserIndex)
                    Exit Sub
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||No hay ningun arbol ahi." & FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        '[/KEVIN]
        Case Mineria
            TempTick = GetTickCount And &H7FFFFFFF
            If TempTick - UserList(UserIndex).Flags.LastWork < IntervaloUserPuedeTrabajar Then Exit Sub
            UserList(UserIndex).Flags.LastWork = TempTick
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
            
        If UserList(UserIndex).Invent.HerramientaEqpObjIndex = PIQUETE_MINERO Or UserList(UserIndex).Invent.HerramientaEqpObjIndex = PIQUETE_MINERON Then '<---[KEVIN]
                    
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            
            auxind = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
            If auxind > 0 Then
                wpaux.Map = UserList(UserIndex).Pos.Map
                wpaux.X = X
                wpaux.Y = Y
                If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                    Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                    Exit Sub
                End If
                '¿Hay un yacimiento donde clickeo?
                If ObjData(auxind).ObjType = OBJTYPE_YACIMIENTO Then
                    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SOUND_MINERO)
                    Call DoMineria(UserIndex)
                    Exit Sub
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
                    Exit Sub
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
            
        Case Domar
        TempTick = GetTickCount And &H7FFFFFFF
        If TempTick - UserList(UserIndex).Flags.LastWork < IntervaloUserPuedeTrabajar Then Exit Sub
        UserList(UserIndex).Flags.LastWork = TempTick
          
          Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
          
          If UserList(UserIndex).Flags.TargetNpc > 0 Then
                   If Npclist(UserList(UserIndex).Flags.TargetNpc).Flags.Domable > 0 Then
                        wpaux.Map = UserList(UserIndex).Pos.Map
                        wpaux.X = X
                        wpaux.Y = Y
                        If Distancia(wpaux, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 2 Then
                              Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                              Exit Sub
                        End If
                        Call DoDomar(UserIndex, UserList(UserIndex).Flags.TargetNpc)
                        Exit Sub
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "||No podes domar a esa criatura." & FONTTYPE_INFO)
                        Exit Sub
                    End If
          Else
                 Call SendData(ToIndex, UserIndex, 0, "||No hay ninguna criatura alli!." & FONTTYPE_INFO)
                 Exit Sub
          End If
          
        Case FundirMetal
            TempTick = GetTickCount And &H7FFFFFFF
            If TempTick - UserList(UserIndex).Flags.LastWork < IntervaloUserPuedeTrabajar Then Exit Sub
            UserList(UserIndex).Flags.LastWork = TempTick
            
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            
            If UserList(UserIndex).Flags.TargetObj > 0 Then
                If ObjData(UserList(UserIndex).Flags.TargetObj).ObjType = OBJTYPE_FRAGUA Then
                    Call FundirMineral(UserIndex)
                    Exit Sub
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)
                    Exit Sub
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)
                Exit Sub
            End If
            
        Case Herreria
            TempTick = GetTickCount And &H7FFFFFFF
            If TempTick - UserList(UserIndex).Flags.LastWork < IntervaloUserPuedeTrabajar Then Exit Sub
            UserList(UserIndex).Flags.LastWork = TempTick
            
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            If UserList(UserIndex).Flags.TargetObj > 0 Then
                If ObjData(UserList(UserIndex).Flags.TargetObj).ObjType = OBJTYPE_YUNQUE Then
                    Call EnivarArmasConstruibles(UserIndex)
                    Call EnivarArmadurasConstruibles(UserIndex)
                    Call SendData(ToIndex, UserIndex, 0, "SFH")
                    Exit Sub
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)
                    Exit Sub
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)
                Exit Sub
            End If
        End Select
    Exit Sub
    Case "USA"
        TempTick = GetTickCount And &H7FFFFFFF
        If TempTick - UserList(UserIndex).Flags.LastUsar < FrmInterv.TextPociones.Text Then Exit Sub
        UserList(UserIndex).Flags.LastUsar = TempTick

        rdata = Right$(rdata, Len(rdata) - 3)
        If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) > 0 Then
            If UserList(UserIndex).Invent.Object(val(rdata)).ObjIndex = 0 Then Exit Sub
        Else
            Exit Sub
        End If
        Call UseInvItem(UserIndex, val(rdata))
        Exit Sub
    Case "CNS" ' Construye herreria
        rdata = Right$(rdata, Len(rdata) - 3)
        X = CInt(rdata)
        If X < 1 Then Exit Sub
        If ObjData(X).SkHerreria = 0 Then Exit Sub
        Call HerreroConstruirItem(UserIndex, X)
        Exit Sub
    Case "CNC" ' Construye carpinteria
        rdata = Right$(rdata, Len(rdata) - 3)
        
        X = CInt(ReadField(1, rdata, 32))
        Arg1 = CInt(ReadField(2, rdata, 32))
        If X < 1 Or ObjData(X).SkCarpinteria = 0 Then Exit Sub
        Call CarpinteroConstruirItem(UserIndex, X, Arg1)
        Exit Sub
    '[KEVIN]
    Case "DCI" ' Crea pociones
        rdata = Right$(rdata, Len(rdata) - 3)
        
        X = CInt(ReadField(1, rdata, 32))
        Arg1 = CInt(ReadField(2, rdata, 32))
        If X < 1 Or ObjData(X).SkPociones = 0 Then Exit Sub
        Call DruidaConstruirItem(UserIndex, X, Arg1)
        Exit Sub
    Case "SCR" ' Crea Ropas
        rdata = Right$(rdata, Len(rdata) - 3)
        
        X = CInt(ReadField(1, rdata, 32))
        Arg1 = CInt(ReadField(2, rdata, 32))
        If X < 1 Or ObjData(X).SkSastreria = 0 Then Exit Sub
        Call SastreConstruirItem(UserIndex, X, Arg1)
        Exit Sub
    '[/KEVIN]
    Case "CIG"
        rdata = Right$(rdata, Len(rdata) - 3)
        X = Guilds.Count
        
        If CreateGuild(UserList(UserIndex).Name, UserList(UserIndex).Reputacion.Promedio, UserIndex, rdata) Then
            If X = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||Felicidades has creado el primer clan de Argentum!!!." & FONTTYPE_INFO)
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Felicidades has creado el clan numero " & X + 1 & " de Argentum!!!." & FONTTYPE_INFO)
            End If
            Call SaveGuildsDB
        End If
        Exit Sub
        
    Case "RPU" 'Pedido de actualizacion de la posicion
        Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
        Exit Sub
    
    Case "TAB" 'Entrar o salir modo combate
        If UserList(UserIndex).Flags.ModoCombate Then
            Call SendData(ToIndex, UserIndex, 0, "||Has salido del modo de combate. " & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Has pasado al modo de combate. " & FONTTYPE_INFO)
        End If
        UserList(UserIndex).Flags.ModoCombate = Not UserList(UserIndex).Flags.ModoCombate
        Exit Sub
    
    Case "SEG" 'Activa / desactiva el seguro
        If UserList(UserIndex).Flags.Seguro Then
              Call SendData(ToIndex, UserIndex, 0, "||Has desactivado el seguro. " & FONTTYPE_INFO)
        Else
              Call SendData(ToIndex, UserIndex, 0, "||Has activado el seguro. " & FONTTYPE_INFO)
        End If
        UserList(UserIndex).Flags.Seguro = Not UserList(UserIndex).Flags.Seguro
        Exit Sub
    
End Select

Select Case Left$(rdata, 4)
    Case "YMCA"
            If Len(rdata) = 4 Then
            If UserList(UserIndex).Flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡No podes atacar a nadie porque estas muerto!!. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If Not UserList(UserIndex).Flags.ModoCombate Then
                Call SendData(ToIndex, UserIndex, 0, "||No estas en modo de combate, presiona la tecla ""C"" para pasar al modo combate. " & FONTTYPE_INFO)
            Else
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                        Call SendData(ToIndex, UserIndex, 0, "||No podés usar asi esta arma." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                Call UsuarioAtaca(UserIndex)
            End If
            Exit Sub
            End If
End Select

If UCase$(Left$(rdata, 6)) = "/CMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    '[KEVIN]Hablar a alguien del clan
    If UserList(UserIndex).Flags.Muerto = 1 Then
        Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
        Exit Sub
    End If
            
    If UserList(UserIndex).GuildInfo.GuildName <> "" Then
        Call SendData(ToGuildMembers, UserIndex, 0, "||" & UserList(UserIndex).Name & ": " & rdata & FONTTYPE_YELLOW & ENDC)
        Exit Sub
    End If
    '[/KEVIN]
End If

If UCase(rdata) = "/ONLINECLAN" Then
        If UserList(UserIndex).GuildInfo.GuildName <> "" Then

            For LoopC = 1 To LastUser
                If UserList(LoopC).GuildInfo.GuildName = UserList(UserIndex).GuildInfo.GuildName Then
                    tStr = tStr & UserList(LoopC).Name & ", "
                End If
            Next LoopC
    
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
        
        End If

        Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/ACEPTO" Then
    rdata = Right$(rdata, Len(rdata) - 7)
    '[KEVIN]
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
            Exit Sub
        End If
            
        If UserList(UserIndex).Flags.TargetUser <> 0 Then
            If UserList(UserIndex).Flags.Pareja = UserList(UserIndex).Flags.TargetUser Then
                If UserList(UserIndex).Genero <> UserList(UserList(UserIndex).Flags.TargetUser).Genero Then
                    UserList(UserIndex).Stats.Matrimonio = UserList(UserList(UserIndex).Flags.TargetUser).Name
                    UserList(UserList(UserIndex).Flags.TargetUser).Stats.Matrimonio = UserList(UserIndex).Name
                    Call SendData(ToIndex, UserList(UserIndex).Flags.TargetUser, 0, "||" & UserList(UserIndex).Name & " ha aceptado, te has casado!" & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "||Te has casado con " & UserList(UserList(UserIndex).Flags.TargetUser).Name & "!" & FONTTYPE_INFO)
                    Exit Sub
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Debes cliquear a la persona que te lo ha propuesto!" & FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call SendData(ToIndex, UserList(UserIndex).Flags.TargetUser, 0, "||No has cliqueado a nadie" & FONTTYPE_INFO)
            Exit Sub
        End If
    '[/KEVIN]
End If

If UCase$(Left$(rdata, 9)) = "/DIVORCIO" Then
    rdata = Right$(rdata, Len(rdata) - 9)
    '[KEVIN]
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Flags.TargetUser <> 0 Then
            If UserList(UserIndex).Stats.Matrimonio = UserList(UserList(UserIndex).Flags.TargetUser).Name Then
                Call SendData(ToIndex, UserIndex, 0, "||" & UserList(UserList(UserIndex).Flags.TargetUser).Name & " ha recivido tu petición." & FONTTYPE_INFO)
                Call SendData(ToIndex, UserList(UserIndex).Flags.TargetUser, 0, "||" & UserList(UserIndex).Name & " te ha pedido el divorcio. Si deseas aceptar hasle click y typea /FIRMAR DIVORCIO" & FONTTYPE_INFO)
                UserList(UserList(UserIndex).Flags.TargetUser).Flags.Divorcio = UserIndex
                Exit Sub
            End If
        Else
            Call SendData(ToIndex, UserList(UserIndex).Flags.TargetUser, 0, "||No has cliqueado a nadie" & FONTTYPE_INFO)
            Exit Sub
        End If
    '[/KEVIN]
End If

If UCase$(Left$(rdata, 11)) = "/CASAMIENTO" Then
    rdata = Right$(rdata, Len(rdata) - 11)
    '[KEVIN]
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Flags.TargetUser <> 0 Then
            If UserList(UserIndex).Genero <> UserList(UserList(UserIndex).Flags.TargetUser).Genero Then
                If UserList(UserList(UserIndex).Flags.TargetUser).Stats.Matrimonio = "" Then
                    Call SendData(ToIndex, UserIndex, 0, "||" & UserList(UserList(UserIndex).Flags.TargetUser).Name & " ha recivido tu petición." & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserList(UserIndex).Flags.TargetUser, 0, "||" & UserList(UserIndex).Name & " te ha propuesto matrimonio. Si deseas aceptar hasle click y typea /ACEPTO" & FONTTYPE_INFO)
                    UserList(UserList(UserIndex).Flags.TargetUser).Flags.Pareja = UserIndex
                    Exit Sub
                End If
            End If
        Else
            Call SendData(ToIndex, UserList(UserIndex).Flags.TargetUser, 0, "||No has cliqueado a nadie" & FONTTYPE_INFO)
            Exit Sub
        End If
    '[/KEVIN]
End If

If UCase$(Left$(rdata, 16)) = "/FIRMAR DIVORCIO" Then
    rdata = Right$(rdata, Len(rdata) - 16)
    '[KEVIN]
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Flags.TargetUser <> 0 Then
            If UserList(UserIndex).Flags.Divorcio = UserList(UserIndex).Flags.TargetUser Then
                    UserList(UserIndex).Stats.Matrimonio = ""
                    UserList(UserList(UserIndex).Flags.TargetUser).Stats.Matrimonio = ""
                    Call SendData(ToIndex, UserList(UserIndex).Flags.TargetUser, 0, "||Te has divorciado!" & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "||Te has divorciado!" & FONTTYPE_INFO)
                    Exit Sub
            Else
                Call SendData(ToIndex, UserIndex, 0, "||Debes cliquear a la persona que te lo ha pedido!" & FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call SendData(ToIndex, UserList(UserIndex).Flags.TargetUser, 0, "||No has cliqueado a nadie" & FONTTYPE_INFO)
            Exit Sub
        End If
    '[/KEVIN]
End If
  
Select Case UCase$(rdata)
    Case "/GM"
        If Not Ayuda.Existe(UserList(UserIndex).Name) Then
            Call SendData(ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
            Call Ayuda.Push(rdata, UserList(UserIndex).Name)
        Else
            Call Ayuda.Quitar(UserList(UserIndex).Name)
            Call Ayuda.Push(rdata, UserList(UserIndex).Name)
            Call SendData(ToIndex, UserIndex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
        End If
        Exit Sub
        
    '[KEVIN]
    Case "/INSCRIBIR"
    
        If EsNewbie(UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "||Los newbies no pueden participar." & FONTTYPE_INFO)
            Exit Sub
        End If
    
        If InscAbierta = False Then
            Call SendData(ToIndex, UserIndex, 0, "||No hay inscripción abierta para ningún torneo o las inscripciones han sido cerradas." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Stats.ELV < MinLevel Then
            Call SendData(ToIndex, UserIndex, 0, "||No tienes el nivel necesario para entrar al torneo, el nivel mínimo requerido es: " & MinLevel & "." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not InscTorneo.Existe(UserList(UserIndex).Name) Then
            Call SendData(ToIndex, UserIndex, 0, "||Te has inscripto al torneo, ahora solo debes al GM." & FONTTYPE_INFO)
            Call InscTorneo.Push(rdata, UserList(UserIndex).Name)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Ya estás inscripto para el torneo." & FONTTYPE_INFO)
        End If
        Exit Sub
    
    Case "/DESINSCRIBIR"
    
        If EsNewbie(UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "||Los newbies no pueden participar." & FONTTYPE_INFO)
            Exit Sub
        End If
    
        If InscAbierta = False Then
            Call SendData(ToIndex, UserIndex, 0, "||No hay inscripción abierta para ningún torneo." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not InscTorneo.Existe(UserList(UserIndex).Name) Then
            Call SendData(ToIndex, UserIndex, 0, "||No te has inscripto a ningún torneo." & FONTTYPE_INFO)
        Else
            Call InscTorneo.Quitar(UserList(UserIndex).Name)
            Call SendData(ToIndex, UserIndex, 0, "||Has sido removido del Torneo." & FONTTYPE_INFO)
        End If
        Exit Sub
    
    '[/KEVIN]
        
    Case "ACTUALIZAR"
        Call SendData(ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
        Exit Sub
    Case "/ONLINE"
        For LoopC = 1 To LastUser
            If (UserList(LoopC).Name <> "") Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
        Call SendData(ToIndex, UserIndex, 0, "||Número de usuarios: " & NumUsers & FONTTYPE_INFO)
        Exit Sub
    Case "/SALIR"
    
        '[KEVIN]*************************
        If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
        End If
        
        If UserList(UserIndex).Invent.Anillo2EqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.Anillo2EqpSlot)
        End If
        
        If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
        End If
        
        Call SendData(ToIndex, UserIndex, 0, "SMT")
        
        '[/KEVIN]********************************
    
        Call SendData(ToIndex, UserIndex, 0, "FINOK")
        Exit Sub
    Case "/FUNDARCLAN"
        If UserList(UserIndex).GuildInfo.FundoClan = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||Ya has fundado un clan, solo se puede fundar uno por personaje." & FONTTYPE_INFO)
            Exit Sub
        End If
        If CanCreateGuild(UserIndex) Then
            Call SendData(ToIndex, UserIndex, 0, "SHOWFUN" & FONTTYPE_INFO)
        End If
        Exit Sub
    Case "GLINFO"
        If UserList(UserIndex).GuildInfo.EsGuildLeader = 1 Then
                    Call SendGuildLeaderInfo(UserIndex)
        Else
                    Call SendGuildsList(UserIndex)
        End If
        Exit Sub
    Case "/BALANCE"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).Flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).Flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).Flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 3 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                  Exit Sub
        End If
        If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO _
        Or UserList(UserIndex).Flags.Muerto = 1 Then Exit Sub
        If FileExist(CharPath & UCase(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
              Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
              CloseSocket (UserIndex)
              Exit Sub
        End If
        Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
        Exit Sub
    Case "/QUIETO" ' << Comando a mascotas
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).Flags.Muerto = 1 Then
                      Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
         End If
         'Se asegura que el target es un npc
         If UserList(UserIndex).Flags.TargetNpc = 0 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
         End If
         If Distancia(Npclist(UserList(UserIndex).Flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
         End If
         If Npclist(UserList(UserIndex).Flags.TargetNpc).MaestroUser <> _
            UserIndex Then Exit Sub
         Npclist(UserList(UserIndex).Flags.TargetNpc).Movement = ESTATICO
         Call Expresar(UserList(UserIndex).Flags.TargetNpc, UserIndex)
         Exit Sub
    Case "/ACOMPAÑAR"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).Flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).Flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).Flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
        End If
        If Npclist(UserList(UserIndex).Flags.TargetNpc).MaestroUser <> _
          UserIndex Then Exit Sub
        Call FollowAmo(UserList(UserIndex).Flags.TargetNpc)
        Call Expresar(UserList(UserIndex).Flags.TargetNpc, UserIndex)
        Exit Sub
    Case "/ENTRENAR"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).Flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).Flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).Flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
        End If
        If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> NPCTYPE_ENTRENADOR Then Exit Sub
        Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).Flags.TargetNpc)
        Exit Sub
    Case "/DESCANSAR"
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        If HayOBJarea(UserList(UserIndex).Pos, FOGATA) Then
                Call SendData(ToIndex, UserIndex, 0, "DOK")
                If Not UserList(UserIndex).Flags.Descansar Then
                    Call SendData(ToIndex, UserIndex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                Else
                    Call SendData(ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                End If
                UserList(UserIndex).Flags.Descansar = Not UserList(UserIndex).Flags.Descansar
        Else
                If UserList(UserIndex).Flags.Descansar Then
                    Call SendData(ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                    
                    UserList(UserIndex).Flags.Descansar = False
                    Call SendData(ToIndex, UserIndex, 0, "DOK")
                    Exit Sub
                End If
                Call SendData(ToIndex, UserIndex, 0, "||No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)
        End If
        Exit Sub
    Case "/MEDITAR"
        If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes meditar cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        Call SendData(ToIndex, UserIndex, 0, "MEDOK")
        If Not UserList(UserIndex).Flags.Meditando Then
           Call SendData(ToIndex, UserIndex, 0, "||Comenzas a meditar." & FONTTYPE_INFO)
        Else
           Call SendData(ToIndex, UserIndex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
        End If
        UserList(UserIndex).Flags.Meditando = Not UserList(UserIndex).Flags.Meditando
        If UserList(UserIndex).Flags.Meditando Then
            UserList(UserIndex).Char.loops = LoopAdEternum
            'MODIFICADO (AGREGUE MAS FXS DE MEDITAR)
            If UserList(UserIndex).Stats.ELV < 8 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARCHICO & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARCHICO
            ElseIf UserList(UserIndex).Stats.ELV < 15 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARCHICOVIOLETA & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARCHICOVIOLETA
            ElseIf UserList(UserIndex).Stats.ELV < 23 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARMEDIANO & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARMEDIANO
            ElseIf UserList(UserIndex).Stats.ELV < 30 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARMEDIANOVERDE & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARMEDIANOVERDE
            ElseIf UserList(UserIndex).Stats.ELV < 38 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARGRANDE & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARGRANDE
            ElseIf UserList(UserIndex).Stats.ELV < 46 Then
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARGRANDEDORADO & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARGRANDEDORADO
            Else
                Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXMEDITARAZUL & "," & LoopAdEternum)
                UserList(UserIndex).Char.FX = FXMEDITARAZUL
            End If
            '******------------*************---------
        Else
            UserList(UserIndex).Char.FX = 0
            UserList(UserIndex).Char.loops = 0
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
        End If
        Exit Sub
    Case "/RESUCITAR"
       'Se asegura que el target es un npc
       If UserList(UserIndex).Flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> 1 _
       Or UserList(UserIndex).Flags.Muerto <> 1 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 10 Then
           Call SendData(ToIndex, UserIndex, 0, "||El sacerdote no puede resucitarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       If FileExist(CharPath & UCase(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
           Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
           CloseSocket (UserIndex)
           Exit Sub
       End If
       Call RevivirUsuario(UserIndex)
       Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás sido resucitado!!" & FONTTYPE_INFO)
       Exit Sub
    Case "/CURAR"
       'Se asegura que el target es un npc
       If UserList(UserIndex).Flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> 1 _
       Or UserList(UserIndex).Flags.Muerto <> 0 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 10 Then
           Call SendData(ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
       Call SendUserStatsBox(val(UserIndex))
       Call SendData(ToIndex, UserIndex, 0, "||¡¡Hás sido curado!!" & FONTTYPE_INFO)
       Exit Sub
    Case "/AYUDA"
       Call SendHelp(UserIndex)
       Exit Sub
    'Case "/GRABAR"
        'Call SaveUser(UserIndex, CharPath & UCase(UserList(UserIndex).Name) & ".chr")
        'Call SendData(ToIndex, UserIndex, 0, "||Caracter grabado." & FONTTYPE_INFO)
        'Exit Sub
    Case "/EST"
        Call SendUserStatsTxt(UserIndex, UserIndex)
        Exit Sub
    Case "ATRI"
        Call EnviarAtrib(UserIndex)
        Exit Sub
    Case "FAMA"
        Call EnviarFama(UserIndex)
        Exit Sub
    Case "ESKI"
        Call EnviarSkills(UserIndex)
        Exit Sub
    Case "/COMERCIAR"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).Flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
        End If
        '¿El target es un NPC valido?
        If UserList(UserIndex).Flags.TargetNpc > 0 Then
              '¿El NPC puede comerciar?
              If Npclist(UserList(UserIndex).Flags.TargetNpc).Comercia = 0 Then
                 If Len(Npclist(UserList(UserIndex).Flags.TargetNpc).Desc) > 0 Then Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                 Exit Sub
              End If
              If Distancia(Npclist(UserList(UserIndex).Flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 3 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                  Exit Sub
              End If
              'Iniciamos la rutina pa' comerciar.
              Call IniciarCOmercioNPC(UserIndex)
         '[Alejo]
        ElseIf UserList(UserIndex).Flags.TargetUser > 0 Then
            'Comercio con otro usuario
            'Puede comerciar ?
            If UserList(UserList(UserIndex).Flags.TargetUser).Flags.Muerto = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "||¡¡No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            'soy yo ?
            If UserList(UserIndex).Flags.TargetUser = UserIndex Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                Exit Sub
            End If
            'ta muy lejos ?
            If Distancia(UserList(UserList(UserIndex).Flags.TargetUser).Pos, UserList(UserIndex).Pos) > 3 Then
                Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del usuario." & FONTTYPE_INFO)
                Exit Sub
            End If
            'Ya ta comerciando ? es con migo o con otro ?
            If UserList(UserList(UserIndex).Flags.TargetUser).Flags.Comerciando = True And _
                UserList(UserList(UserIndex).Flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call SendData(ToIndex, UserIndex, 0, "||No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
                Exit Sub
            End If
            'inicializa unas variables...
            UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).Flags.TargetUser
            UserList(UserIndex).ComUsu.Cant = 0
            UserList(UserIndex).ComUsu.Objeto = 0
            UserList(UserIndex).ComUsu.Acepto = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).Flags.TargetUser)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
        End If
        Exit Sub
    '[/Alejo]
    '#####################################TEMPORAL#################################
    Case "/COFRE"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).Flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
        End If
        If ObjData(UserList(UserIndex).Flags.TargetObj).ObjType = OBJTYPE_CONTENEDORES Then
              Call IniciarDepCofre(UserIndex)
        Else
          Call SendData(ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el cofre." & FONTTYPE_INFO)
        End If
        Exit Sub
    '###############################################################################
    '[KEVIN]------------------------------------------
    Case "/BOVEDA"
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).Flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
        End If
        '¿El target es un NPC valido?
        If UserList(UserIndex).Flags.TargetNpc > 0 Then
              If Distancia(Npclist(UserList(UserIndex).Flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 3 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                  Exit Sub
              End If
              If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype = 4 Then
                Call IniciarDeposito(UserIndex)
              Else
                Exit Sub
              End If
        Else
          Call SendData(ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
        End If
        Exit Sub
    '[/KEVIN]------------------------------------
    '[Alejo]
    Case "FINCOM"
        'User sale del modo COMERCIO
        UserList(UserIndex).Flags.Comerciando = False
        Call SendData(ToIndex, UserIndex, 0, "FINCOMOK")
        Exit Sub
        Case "FINCOMUSU"
        'Sale modo comercio Usuario
        If UserList(UserIndex).ComUsu.DestUsu > 0 And _
            UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
            Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha dejado de comerciar con vos." & FONTTYPE_TALK)
            Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
        End If
        
        Call FinComerciarUsu(UserIndex)
        Exit Sub
    '[KEVIN]---------------------------------------
    '******************************************************
    Case "FINBAN"
        'User sale del modo BANCO
        UserList(UserIndex).Flags.Comerciando = False
        Call SendData(ToIndex, UserIndex, 0, "FINBANOK")
        Exit Sub
    '-------------------------------------------------------
    '[/KEVIN]***********************************************
    '[KEVIN]---------------------------------------
    '******************************************************
    Case "FINCOF"
        'User sale del modo COFRE
        UserList(UserIndex).Flags.Comerciando = False
        Call SendData(ToIndex, UserIndex, 0, "FINCOFOK")
        Exit Sub
    '-------------------------------------------------------
    '[/KEVIN]***********************************************
    Case "COMUSUOK"
        'Aceptar el cambio
        Call AceptarComercioUsu(UserIndex)
        Exit Sub
    Case "COMUSUNO"
        'Rechazar el cambio
        If UserList(UserIndex).ComUsu.DestUsu > 0 Then
            Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha rechazado tu oferta." & FONTTYPE_TALK)
            Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
        End If
        Call SendData(ToIndex, UserIndex, 0, "||Has rechazado la oferta del otro usuario." & FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)
        Exit Sub
    '[/Alejo]
    '[KEVIN]
    Case "/ENLISTAR"
        'Se asegura que el target es un npc
       If UserList(UserIndex).Flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       If UserList(UserIndex).Flags.Muerto <> 0 Then Exit Sub
       
       Select Case Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype
            Case 5
                If UserList(UserIndex).Raza <> "Humano" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un humano lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 10
                If UserList(UserIndex).Raza <> "Elfo" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un elfo lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 11
                If UserList(UserIndex).Raza <> "Elfo Oscuro" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un Elfo Oscuro lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 12
                If UserList(UserIndex).Raza <> "Gnomo" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un Gnomo lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 13
                If UserList(UserIndex).Raza <> "Enano" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un Enano lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 14
                If UserList(UserIndex).Raza <> "Orco" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un Orco lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case Else
                Exit Sub
        End Select
    
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       If Npclist(UserList(UserIndex).Flags.TargetNpc).Flags.Faccion = 0 Then
              Call EnlistarArmadaReal(UserIndex)
       Else
              Call EnlistarCaos(UserIndex)
       End If
       
       Exit Sub
    '[/KEVIN]
    '[KEVIN]
    Case "/AVENTURA"
        'Se asegura que el target es un npc
       If UserList(UserIndex).Flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> 8 _
       Or UserList(UserIndex).Flags.Muerto <> 0 Then Exit Sub
       
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "||Estás demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       Call HacerQuest(UserIndex)
       
       Exit Sub
    '[/KEVIN]
    '[KEVIN]
    Case "/REWARD"
       'Se asegura que el target es un npc
       If UserList(UserIndex).Flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> 8 _
       Or UserList(UserIndex).Flags.Muerto <> 0 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "||Estás demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       Call RecibirRecompensaQuest(UserIndex)
       
       Exit Sub
    '[/KEVIN]
    '[KEVIN]
    Case "/INFOQ"
       'Se asegura que el target es un npc
       If UserList(UserIndex).Flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> 8 _
       Or UserList(UserIndex).Flags.Muerto <> 0 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "||Estás demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       Call SendInfoQuest(UserIndex)
       
       Exit Sub
    '[/KEVIN]
    '[KEVIN]
    Case "/MERINDO"
       'Se asegura que el target es un npc
       If UserList(UserIndex).Flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> 8 _
       Or UserList(UserIndex).Flags.Muerto <> 0 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "||Estás demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       Call UserSeRinde(UserIndex)
       
       Exit Sub
    '[/KEVIN]
    '[KEVIN]
    Case "/ADIVINA"
       'Se asegura que el target es un npc
       If UserList(UserIndex).Flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       If UserList(UserIndex).Flags.Muerto <> 0 Then Exit Sub
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "||Estás demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       Call CheckNpcAmigo(UserIndex)
       
       Exit Sub
    '[/KEVIN]
    Case "/INFORMACION"
       'Se asegura que el target es un npc
       If UserList(UserIndex).Flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       If UserList(UserIndex).Flags.Muerto <> 0 Then Exit Sub
       '[KEVIN]
       Select Case Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype
            Case 5
                If UserList(UserIndex).Raza <> "Humano" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un humano lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 10
                If UserList(UserIndex).Raza <> "Elfo" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un elfo lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 11
                If UserList(UserIndex).Raza <> "Elfo Oscuro" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un Elfo Oscuro lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 12
                If UserList(UserIndex).Raza <> "Gnomo" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un Gnomo lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 13
                If UserList(UserIndex).Raza <> "Enano" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un Enano lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 14
                If UserList(UserIndex).Raza <> "Orco" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un Orco lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case Else
                Exit Sub
        End Select
       '[/KEVIN]
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       
       If Npclist(UserList(UserIndex).Flags.TargetNpc).Flags.Faccion = 0 Then
            If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
       Else
            If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las fuerzas del caos!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
       End If
       Exit Sub
    Case "/RECOMPENSA"
       'Se asegura que el target es un npc
       If UserList(UserIndex).Flags.TargetNpc = 0 Then
           Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
           Exit Sub
       End If
       If UserList(UserIndex).Flags.Muerto <> 0 Then Exit Sub
       '[KEVIN]
       Select Case Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype
            Case 5
                If UserList(UserIndex).Raza <> "Humano" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un humano lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 10
                If UserList(UserIndex).Raza <> "Elfo" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un elfo lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 11
                If UserList(UserIndex).Raza <> "Elfo Oscuro" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un Elfo Oscuro lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 12
                If UserList(UserIndex).Raza <> "Gnomo" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un Gnomo lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 13
                If UserList(UserIndex).Raza <> "Enano" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un Enano lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case 14
                If UserList(UserIndex).Raza <> "Orco" Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu no eres un Orco lárgate de aquí!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
                End If
            Case Else
                Exit Sub
        End Select
       '[/KEVIN]
       If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 4 Then
           Call SendData(ToIndex, UserIndex, 0, "||Estás demasiado lejos." & FONTTYPE_INFO)
           Exit Sub
       End If
       If Npclist(UserList(UserIndex).Flags.TargetNpc).Flags.Faccion = 0 Then
            If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call RecompensaArmadaReal(UserIndex)
       Else
            If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las fuerzas del caos!!!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                Exit Sub
            End If
            Call RecompensaCaos(UserIndex)
       End If
       Exit Sub
End Select

Select Case UCase$(Left$(rdata, 4))
    Case "INFS" 'Informacion del hechizo
            rdata = Right$(rdata, Len(rdata) - 4)
            If val(rdata) > 0 And val(rdata) < MAXUSERHECHIZOS + 1 Then
                Dim H As Integer
                H = UserList(UserIndex).Stats.UserHechizos(val(rdata))
                If H > 0 And H < NumeroHechizos + 1 Then
                    Call SendData(ToIndex, UserIndex, 0, "||%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "||Nombre:" & Hechizos(H).Nombre & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "||Descripcion:" & Hechizos(H).Desc & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "||Skill requerido: " & Hechizos(H).MinSkill & " de magia." & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "||Mana necesario: " & Hechizos(H).ManaRequerido & FONTTYPE_INFO)
                    Call SendData(ToIndex, UserIndex, 0, "||%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%" & FONTTYPE_INFO)
                End If
            Else
                Call SendData(ToIndex, UserIndex, 0, "||¡Primero selecciona el hechizo.!" & FONTTYPE_INFO)
            End If
            Exit Sub
   Case "EQUI"
            If UserList(UserIndex).Flags.Muerto = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
            End If
            rdata = Right$(rdata, Len(rdata) - 4)
            If val(rdata) <= MAX_INVENTORY_SLOTS And val(rdata) > 0 Then
                 If UserList(UserIndex).Invent.Object(val(rdata)).ObjIndex = 0 Then Exit Sub
            Else
                Exit Sub
            End If
            Call EquiparInvItem(UserIndex, val(rdata))
            Exit Sub
    Case "CHEA" 'Cambiar Heading ;-)
        rdata = Right$(rdata, Len(rdata) - 4)
        If val(rdata) > 0 And val(rdata) < 5 Then
            UserList(UserIndex).Char.Heading = rdata
            Call ChangeUserChar(ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
        Exit Sub
    Case "SKSE" 'Modificar skills
        Dim i As Integer
        Dim sumatoria As Integer
        Dim incremento As Integer
        rdata = Right$(rdata, Len(rdata) - 4)
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rdata, 44))
            
            '[KEVIN]
            If HaciendoBKPJS = True Then
                Exit Sub
            End If
            '[/KEVIN]
            
            If incremento < 0 Then
                'Call SendData(ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
                UserList(UserIndex).Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            sumatoria = sumatoria + incremento
        Next i
        
        If sumatoria > UserList(UserIndex).Stats.SkillPts Then
            'UserList(UserIndex).Flags.AdministrativeBan = 1
            'Call SendData(ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
            Call LogHackAttemp(UserList(UserIndex).Name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        For i = 1 To NUMSKILLS
            incremento = val(ReadField(i, rdata, 44))
            UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts - incremento
            UserList(UserIndex).Stats.UserSkills(i) = UserList(UserIndex).Stats.UserSkills(i) + incremento
            If UserList(UserIndex).Stats.UserSkills(i) > 100 Then UserList(UserIndex).Stats.UserSkills(i) = 100
        Next i
        Exit Sub
    Case "ENTR" 'Entrena hombre!
        
        If UserList(UserIndex).Flags.TargetNpc = 0 Then Exit Sub
        
        If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> 3 Then Exit Sub
        
        rdata = Right$(rdata, Len(rdata) - 4)
        
        If Npclist(UserList(UserIndex).Flags.TargetNpc).Mascotas < MAXMASCOTASENTRENADOR Then
            If val(rdata) > 0 And val(rdata) < Npclist(UserList(UserIndex).Flags.TargetNpc).NroCriaturas + 1 Then
                    Dim SpawnedNpc As Integer
                    SpawnedNpc = SpawnNpc(Npclist(UserList(UserIndex).Flags.TargetNpc).Criaturas(val(rdata)).NpcIndex, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos, True, False)
                    If SpawnedNpc <= MAXNPCS Then
                        Npclist(SpawnedNpc).MaestroNpc = UserList(UserIndex).Flags.TargetNpc
                        Npclist(UserList(UserIndex).Flags.TargetNpc).Mascotas = Npclist(UserList(UserIndex).Flags.TargetNpc).Mascotas + 1
                    End If
            End If
        Else
            Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & "No puedo traer mas criaturas, mata las existentes!" & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
        End If
        
        Exit Sub
    Case "COMP"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).Flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                   Exit Sub
         End If
         '¿El target es un NPC valido?
         If UserList(UserIndex).Flags.TargetNpc > 0 Then
               '¿El NPC puede comerciar?
               If Npclist(UserList(UserIndex).Flags.TargetNpc).Comercia = 0 Then
                   Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                   Exit Sub
               End If
         Else
           Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 5)
         'User compra el item del slot rdata
         Call NPCVentaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)), UserList(UserIndex).Flags.TargetNpc)
         Exit Sub
    '[KEVIN]*********************************************************************
    '------------------------------------------------------------------------------------
    Case "RETI"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).Flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                   Exit Sub
         End If
         '¿El target es un NPC valido?
         If UserList(UserIndex).Flags.TargetNpc > 0 Then
               '¿Es el banquero?
               If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> 4 Then
                   Exit Sub
               End If
         Else
           Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 5)
         'User retira el item del slot rdata
         Call UserRetiraItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
    '-----------------------------------------------------------------------------------
    '[/KEVIN]****************************************************************************
    '[KEVIN]*********************************************************************
    '------------------------------------------------------------------------------------
    Case "CRET"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).Flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                   Exit Sub
         End If
        
        If ObjData(UserList(UserIndex).Flags.TargetObj).ObjType <> OBJTYPE_CONTENEDORES Then Exit Sub

         rdata = Right$(rdata, Len(rdata) - 5)
         'User retira el item del slot rdata
         Call UserCRetiraItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
    '-----------------------------------------------------------------------------------
    '[/KEVIN]****************************************************************************
    Case "VEND"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).Flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                   Exit Sub
         End If
         '¿El target es un NPC valido?
         If UserList(UserIndex).Flags.TargetNpc > 0 Then
               '¿El NPC puede comerciar?
               If Npclist(UserList(UserIndex).Flags.TargetNpc).Comercia = 0 Then
                   Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & Str(Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex))
                   Exit Sub
               End If
         Else
           Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 5)
         'User compra el item del slot rdata
         Call NPCCompraItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
    '[KEVIN]-------------------------------------------------------------------------
    '****************************************************************************************
    Case "DEPO"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).Flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                   Exit Sub
         End If
         '¿El target es un NPC valido?
         If UserList(UserIndex).Flags.TargetNpc > 0 Then
               '¿El NPC puede comerciar?
               If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> 4 Then
                   Exit Sub
               End If
         Else
           Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 5)
         'User deposita el item del slot rdata
         Call UserDepositaItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
    '****************************************************************************************
    '[/KEVIN]---------------------------------------------------------------------------------
    '[KEVIN]-------------------------------------------------------------------------
    '****************************************************************************************
    Case "CDEP"
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).Flags.Muerto = 1 Then
                   Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                   Exit Sub
         End If

        If ObjData(UserList(UserIndex).Flags.TargetObj).ObjType <> OBJTYPE_CONTENEDORES Then Exit Sub

         rdata = Right$(rdata, Len(rdata) - 5)
         'User deposita el item del slot rdata
         Call UserDepositaCItem(UserIndex, val(ReadField(1, rdata, 44)), val(ReadField(2, rdata, 44)))
         Exit Sub
    '****************************************************************************************
    '[/KEVIN]---------------------------------------------------------------------------------
End Select

Select Case UCase$(Left$(rdata, 5))
    Case "DEMSG"
        If UserList(UserIndex).Flags.TargetObj > 0 Then
        rdata = Right$(rdata, Len(rdata) - 5)
        Dim f As String, Titu As String, msg As String, f2 As String
        f = App.Path & "\foros\"
        f = f & UCase(ObjData(UserList(UserIndex).Flags.TargetObj).ForoID) & ".for"
        '[KEVIN]
        Titu = UserList(UserIndex).Name & ": " & ReadField(1, rdata, 176)
        '[/KEVIN]
        msg = ReadField(2, rdata, 176)
        Dim n2 As Integer, loopme As Integer
        If FileExist(f, vbNormal) Then
            Dim num As Integer
            num = val(GetVar(f, "INFO", "CantMSG"))
            If num > MAX_MENSAJES_FORO Then
                For loopme = 1 To num
                    Kill App.Path & "\foros\" & UCase(ObjData(UserList(UserIndex).Flags.TargetObj).ForoID) & loopme & ".for"
                Next
                Kill App.Path & "\foros\" & UCase(ObjData(UserList(UserIndex).Flags.TargetObj).ForoID) & ".for"
                num = 0
            End If
            n2 = FreeFile
            f2 = Left$(f, Len(f) - 4)
            f2 = f2 & num + 1 & ".for"
            Open f2 For Output As n2
            Print #n2, Titu
            Print #n2, msg
            Call WriteVar(f, "INFO", "CantMSG", num + 1)
        Else
            n2 = FreeFile
            f2 = Left$(f, Len(f) - 4)
            f2 = f2 & "1" & ".for"
            Open f2 For Output As n2
            Print #n2, Titu
            Print #n2, msg
            Call WriteVar(f, "INFO", "CantMSG", 1)
        End If
        Close #n2
        End If
        Exit Sub
    Case "/BUG "
        N = FreeFile
        Open App.Path & "\BUGS\BUGs.log" For Append Shared As N
        Print #N,
        Print #N,
        Print #N, "########################################################################"
        Print #N, "########################################################################"
        Print #N, "Usuario:" & UserList(UserIndex).Name & "  Fecha:" & Date & "    Hora:" & Time
        Print #N, "########################################################################"
        Print #N, "BUG:"
        Print #N, Right$(rdata, Len(rdata) - 5)
        Print #N, "########################################################################"
        Print #N, "########################################################################"
        Print #N,
        Print #N,
        Close #N
        Exit Sub
End Select


Select Case UCase$(Left$(rdata, 6))
    Case "/DESC "
        rdata = Right$(rdata, Len(rdata) - 6)
        If Not AsciiValidos(rdata) Then
            Call SendData(ToIndex, UserIndex, 0, "||La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(UserIndex).Desc = rdata
        Call SendData(ToIndex, UserIndex, 0, "||La descripcion a cambiado." & FONTTYPE_INFO)
        Exit Sub
    Case "DESCOD" 'Informacion del hechizo
            rdata = Right$(rdata, Len(rdata) - 6)
            Call UpdateCodexAndDesc(rdata, UserIndex)
            Exit Sub
    Case "/VOTO "
            rdata = Right$(rdata, Len(rdata) - 6)
            Call ComputeVote(UserIndex, rdata)
            Exit Sub
            
 End Select
 
'[Alejo]
Select Case UCase$(Left$(rdata, 7))
Case "OFRECER"
        rdata = Right$(rdata, Len(rdata) - 7)
        Arg1 = ReadField(1, rdata, Asc(","))
        Arg2 = ReadField(2, rdata, Asc(","))

        If val(Arg1) <= 0 Or val(Arg2) <= 0 Then
            Exit Sub
        End If
        If UserList(UserList(UserIndex).ComUsu.DestUsu).Flags.UserLogged = False Then
            'sigue vivo el usuario ?
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        Else
            'esta vivo ?
            If UserList(UserList(UserIndex).ComUsu.DestUsu).Flags.Muerto = 1 Then
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            End If
            '//Tiene la cantidad que ofrece ??//'
            If val(Arg1) = FLAGORO Then
                'oro
                If val(Arg2) > UserList(UserIndex).Stats.GLD Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                    Exit Sub
                End If
            Else
                'inventario
                If val(Arg2) > UserList(UserIndex).Invent.Object(val(Arg1)).Amount Then
                    Call SendData(ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            UserList(UserIndex).ComUsu.Objeto = val(Arg1)
            UserList(UserIndex).ComUsu.Cant = val(Arg2)
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex Then
                'Es el primero que ofrece algo ?
                Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR." & FONTTYPE_TALK)
                UserList(UserList(UserIndex).ComUsu.DestUsu).Flags.TargetUser = UserIndex
            Else
                '[CORREGIDO]
                If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = True Then
                    'NO NO NO vos te estas pasando de listo...
                    UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False
                    Call SendData(ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).Name & " ha cambiado su oferta." & FONTTYPE_TALK)
                End If
                '[/CORREGIDO]
                'Es la ofrenda de respuesta :)
                Call EnviarObjetoTransaccion(UserList(UserIndex).ComUsu.DestUsu)
            End If
        End If
        Exit Sub
End Select
'[/Alejo]


Select Case UCase$(Left$(rdata, 8))
    Case "ACEPPEAT"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call AcceptPeaceOffer(UserIndex, rdata)
        Exit Sub
    Case "PEACEOFF"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call RecievePeaceOffer(UserIndex, rdata)
        Exit Sub
    Case "PEACEDET"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendPeaceRequest(UserIndex, rdata)
        Exit Sub
    Case "ENVCOMEN"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendPeticion(UserIndex, rdata)
        Exit Sub
    Case "ENVPROPP"
        Call SendPeacePropositions(UserIndex)
        Exit Sub
    Case "DECGUERR"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DeclareWar(UserIndex, rdata)
        Exit Sub
    Case "DECALIAD"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DeclareAllie(UserIndex, rdata)
        Exit Sub
    Case "NEWWEBSI"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SetNewURL(UserIndex, rdata)
        Exit Sub
    Case "ACEPTARI"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call AcceptClanMember(UserIndex, rdata)
        Exit Sub
    Case "RECHAZAR"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call DenyRequest(UserIndex, rdata)
        Exit Sub
    Case "ECHARCLA"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call EacharMember(UserIndex, rdata)
        Exit Sub
    Case "/PASSWD "
        rdata = Right$(rdata, Len(rdata) - 8)
        If Len(rdata) < 6 Then
             Call SendData(ToIndex, UserIndex, 0, "||El password debe tener al menos 6 caracteres." & FONTTYPE_INFO)
        Else
             Call SendData(ToIndex, UserIndex, 0, "||El password ha sido cambiado." & FONTTYPE_INFO)
             UserList(UserIndex).Password = rdata
        End If
        Exit Sub
    Case "ACTGNEWS"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call UpdateGuildNews(rdata, UserIndex)
        Exit Sub
    Case "1HRINFO<"
        rdata = Right$(rdata, Len(rdata) - 8)
        Call SendCharInfo(rdata, UserIndex)
        Exit Sub
End Select


Select Case UCase$(Left$(rdata, 9))
    Case "SOLICITUD"
         rdata = Right$(rdata, Len(rdata) - 9)
         Call SolicitudIngresoClan(UserIndex, rdata)
         Exit Sub
    Case "/RETIRAR " 'RETIRA ORO EN EL BANCO
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).Flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
         End If
         'Se asegura que el target es un npc
         If UserList(UserIndex).Flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
              Exit Sub
         End If
         rdata = Right$(rdata, Len(rdata) - 9)
         If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO _
         Or UserList(UserIndex).Flags.Muerto = 1 Then Exit Sub
         If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 10 Then
              Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
              Exit Sub
         End If
         If FileExist(CharPath & UCase(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
              Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
              CloseSocket (UserIndex)
              Exit Sub
         End If
         If val(rdata) > 0 And val(rdata) <= UserList(UserIndex).Stats.Banco Then
              UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(rdata)
              UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rdata)
              Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
         Else
              Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
         End If
         Call SendUserStatsBox(val(UserIndex))
         Exit Sub
End Select


Select Case UCase$(Left$(rdata, 11))
    Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
        '¿Esta el user muerto? Si es asi no puede comerciar
        If UserList(UserIndex).Flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
        End If
        'Se asegura que el target es un npc
        If UserList(UserIndex).Flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
              Exit Sub
        End If
        If Distancia(Npclist(UserList(UserIndex).Flags.TargetNpc).Pos, UserList(UserIndex).Pos) > 10 Then
                  Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
        End If
        rdata = Right$(rdata, Len(rdata) - 11)
        If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO _
        Or UserList(UserIndex).Flags.Muerto = 1 Then Exit Sub
        If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 10 Then
              Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
              Exit Sub
        End If
        If CLng(val(rdata)) > 0 And CLng(val(rdata)) <= UserList(UserIndex).Stats.GLD Then
              UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(rdata)
              UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rdata)
              Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
        Else
              Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
        End If
        Call SendUserStatsBox(val(UserIndex))
        Exit Sub
  Case "CLANDETAILS"
        rdata = Right$(rdata, Len(rdata) - 11)
        Call SendGuildDetails(UserIndex, rdata)
        Exit Sub
End Select

Select Case UCase$(Left$(rdata, 12))
'[KEVIN]
    Case "/TRANSFERIR " 'TRANSFERIR ORO A LA CUENTA DE OTRO
         rdata = Right$(rdata, Len(rdata) - 12)
         Arg1 = val(ReadField(1, rdata, 32))
         Arg2 = ReadField(2, rdata, 45)
         tIndex = NameIndex(Arg2)
         '¿Esta el user muerto? Si es asi no puede comerciar
         If UserList(UserIndex).Flags.Muerto = 1 Then
                  Call SendData(ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                  Exit Sub
         End If
         'Se asegura que el target es un npc
         If UserList(UserIndex).Flags.TargetNpc = 0 Then
              Call SendData(ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
              Exit Sub
         End If
         
         If Npclist(UserList(UserIndex).Flags.TargetNpc).NPCtype <> NPCTYPE_BANQUERO _
         Or UserList(UserIndex).Flags.Muerto = 1 Then Exit Sub
         If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).Flags.TargetNpc).Pos) > 10 Then
              Call SendData(ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
              Exit Sub
         End If
         If FileExist(CharPath & UCase(UserList(UserIndex).Name) & ".chr", vbNormal) = False Then
              Call SendData(ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
              CloseSocket (UserIndex)
              Exit Sub
         End If
         
         '¿Existe el pj?
         If FileExist(CharPath & UCase(Arg2) & ".chr", vbNormal) = False Then
              Call SendData(ToIndex, UserIndex, 0, "||El personaje no existe." & FONTTYPE_INFO)
              Exit Sub
         Else
            '¿Está online?
            If tIndex > 0 Then
               If val(Arg1) > 0 And val(Arg1) <= UserList(UserIndex).Stats.Banco Then
                    UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(Arg1)
                    UserList(tIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(Arg1)
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Has transferido " & Arg1 & " monedas de oro a la cuenta de " & UserList(tIndex).Name & "." & "°" & Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
               Else
                    Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
               End If
            Else
               If val(Arg1) > 0 And val(Arg1) <= UserList(UserIndex).Stats.Banco Then
                   Dim TAccount As Long
                   TAccount = val(GetVar(CharPath & UCase(Arg2) & ".chr", "STATS", "BANCO"))
                   TAccount = TAccount + val(Arg1)
                   Call WriteVar(CharPath & UCase(Arg2) & ".chr", "STATS", "BANCO", Str(TAccount))
                   UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(Arg1)
                   Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Has transferido " & Arg1 & " monedas de oro a la cuenta de " & Arg2 & "." & "°" & Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
               Else
                   Call SendData(ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).Flags.TargetNpc).Char.CharIndex & FONTTYPE_INFO)
               End If
            End If
         End If
         
         Exit Sub
    '[/KEVIN]
End Select

'EN CASO DE ROBO EL PJ SERVERSHUTDOWN PUEDE APAGAR EL SERVER
If UCase$(Left$(rdata, 7)) = "\APAGAR" Then
    rdata = Right$(rdata, Len(rdata) - 7)
    
    If UserList(UserIndex).Name <> "SERVERSHUTDOWN" Then Exit Sub
    
    End
    
    Exit Sub
End If


'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
 If UserList(UserIndex).Flags.Privilegios = 0 Then Exit Sub
'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<


'<<<<<<<<<<<<<<<<<<<< Consejeros <<<<<<<<<<<<<<<<<<<<

'HORA
If UCase$(Left$(rdata, 5)) = "/HORA" Then
    Call LogGM(UserList(UserIndex).Name, "Hora.")
    rdata = Right$(rdata, Len(rdata) - 5)
    Call SendData(ToAll, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_INFO)
    Exit Sub
End If

'¿Donde esta?
If UCase$(Left$(rdata, 7)) = "/DONDE " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    Call SendData(ToIndex, UserIndex, 0, "||Ubicacion  " & UserList(tIndex).Name & ": " & UserList(tIndex).Pos.Map & ", " & UserList(tIndex).Pos.X & ", " & UserList(tIndex).Pos.Y & "." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "/Donde")
    Exit Sub
End If

'INFO DE USER
If UCase(Left$(rdata, 6)) = "/INFO " Then
    Call LogGM(UserList(UserIndex).Name, rdata)
    
    rdata = Right$(rdata, Len(rdata) - 6)
    
    tIndex = NameIndex(rdata)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If

    SendUserStatsTxt UserIndex, tIndex
    Exit Sub
End If

'INV DEL USER
If UCase(Left$(rdata, 5)) = "/INV " Then
    Call LogGM(UserList(UserIndex).Name, rdata)
    
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = NameIndex(rdata)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If

    SendUserInvTxt UserIndex, tIndex
    Exit Sub
End If

'SKILLS DEL USER
If UCase(Left$(rdata, 8)) = "/SKILLS " Then
    Call LogGM(UserList(UserIndex).Name, rdata)
    
    rdata = Right$(rdata, Len(rdata) - 8)
    
    tIndex = NameIndex(rdata)
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If

    SendUserSkillsTxt UserIndex, tIndex
    Exit Sub
End If

'Nro de enemigos
If UCase$(Left$(rdata, 6)) = "/NENE " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    If MapaValido(val(rdata)) Then
        Call SendData(ToIndex, UserIndex, 0, "NENE" & NPCHostiles(val(rdata)))
        Call LogGM(UserList(UserIndex).Name, "Numero enemigos en mapa " & rdata)
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/REVIVIR " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    Name = rdata
    If UCase$(Name) <> "YO" Then
        tIndex = NameIndex(Name)
    Else
        tIndex = UserIndex
    End If
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    UserList(tIndex).Flags.Muerto = 0
    UserList(tIndex).Stats.MinHP = UserList(tIndex).Stats.MaxHP
    Call DarCuerpoDesnudo(tIndex)
    Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, val(tIndex), UserList(tIndex).Char.Body, UserList(tIndex).OrigChar.Head, UserList(tIndex).Char.Heading, UserList(tIndex).Char.WeaponAnim, UserList(tIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Call SendUserStatsBox(val(tIndex))
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " te há resucitado." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "Resucito a " & UserList(tIndex).Name)
    Exit Sub
End If

'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<< SemiDioses <<<<<<<<<<<<<<<<<<<<<<<<
If UserList(UserIndex).Flags.Privilegios < 2 Then
    Exit Sub
End If

If UCase(rdata) = "/ONLINEGM" Then
        For LoopC = 1 To LastUser
            If (UserList(LoopC).Name <> "") And UserList(LoopC).Flags.Privilegios <> 0 Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, UserIndex, 0, "||No hay GMs Online" & FONTTYPE_INFO)
        End If
        Exit Sub
End If

'Bloquear
If UCase$(Left$(rdata, 5)) = "/BLOQ" Then
    Call LogGM(UserList(UserIndex).Name, "/BLOQ")
    rdata = Right$(rdata, Len(rdata) - 5)
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 0 Then
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 1
        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 1)
    Else
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 0
        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 0)
    End If
    Exit Sub
End If


If UCase(Left$(rdata, 8)) = "/CARCEL " Then
    
    rdata = Right$(rdata, Len(rdata) - 8)
    
    Name = ReadField(1, rdata, 32)
    i = val(ReadField(1, rdata, 32))
    Name = Right$(rdata, Len(rdata) - (Len(Name) + 1))
    
    tIndex = NameIndex(Name)
    
'    If UCase(Name) = "MORGOLOCK" Then Exit Sub
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).Flags.Privilegios > UserList(UserIndex).Flags.Privilegios Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes encarcelar a alguien con jerarquia mayor a la tuya." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If i > 30 Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes encarcelar por mas de 30 minutos." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    UserList(tIndex).Flags.EnTorneo = 0
    Call Encarcelar(tIndex, i, UserList(UserIndex).Name)
    
    Exit Sub
End If

If UCase(Left$(rdata, 9)) = "/SHOW SOS" Then
    Dim m As String
    For N = 1 To Ayuda.Longitud
        m = Ayuda.VerElemento(N)
        Call SendData(ToIndex, UserIndex, 0, "RSOS" & m)
    Next N
    Call SendData(ToIndex, UserIndex, 0, "MSOS")
    Exit Sub
End If

'[KEVIN]
If UCase(Left$(rdata, 8)) = "/VERINSC" Then
    Dim mm As String
    For N = 1 To InscTorneo.Longitud
        mm = InscTorneo.VerElemento(N)
        Call SendData(ToIndex, UserIndex, 0, "IRSOS" & mm)
    Next N
    Call SendData(ToIndex, UserIndex, 0, "IMSOS")
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
        Call SendData(ToIndex, UserIndex, 0, "||Debes escribir el nivel mínimo de los pjs a entrar" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    MinLevel = rdata
        
    If MinLevel < 13 Then
        Call SendData(ToIndex, UserIndex, 0, "||Los newbies no pueden participar de los torneos" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    InscAbierta = True
    Call SendData(ToIndex, UserIndex, 0, "||Has abierto la inscripción." & FONTTYPE_INFO)
    Exit Sub
End If

If UCase(Left$(rdata, 7)) = "/CERRAR" Then
    rdata = Right$(rdata, Len(rdata) - 7)
    InscAbierta = False
    MinLevel = 1
    Call SendData(ToIndex, UserIndex, 0, "||Has cerrado la inscripción." & FONTTYPE_INFO)
    Exit Sub
End If
'[/KEVIN]

If UCase(Left$(rdata, 7)) = "SOSDONE" Then
    rdata = Right$(rdata, Len(rdata) - 7)
    Call Ayuda.Quitar(rdata)
    Exit Sub
End If

'PERDON
If UCase$(Left$(rdata, 8)) = "/PERDON " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    tIndex = NameIndex(rdata)
    If tIndex > 0 Then
        
        If EsNewbie(tIndex) Then
                Call VolverCiudadano(tIndex)
        Else
                Call LogGM(UserList(UserIndex).Name, "Intento perdonar un personaje de nivel avanzado.")
                Call SendData(ToIndex, UserIndex, 0, "||Solo se permite perdonar newbies." & FONTTYPE_INFO)
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
    If UCase(rdata) = "NEB" Then Exit Sub
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).Flags.Privilegios > UserList(UserIndex).Flags.Privilegios Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes echar a alguien con jerarquia mayor a la tuya." & FONTTYPE_INFO)
        Exit Sub
    End If
        
    Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(tIndex).Name & "." & FONTTYPE_INFO)
    Call CloseSocket(tIndex)
    Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(tIndex).Name)
    Exit Sub
End If


If UCase$(Left$(rdata, 5)) = "/BAN " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    tIndex = NameIndex(ReadField(2, rdata, Asc("@")))
    Name = ReadField(1, rdata, Asc("@"))
    
    If UCase(Name) = "NEB" Then Exit Sub
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(tIndex).Flags.Privilegios > UserList(UserIndex).Flags.Privilegios Then
        Call SendData(ToIndex, UserIndex, 0, "||No podes banear a al alguien de mayor jerarquia." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call LogBan(tIndex, UserIndex, Name)
    
    Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " echo a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
    Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " Banned a " & UserList(tIndex).Name & "." & FONTTYPE_FIGHT)
    
    'Ponemos el flag de ban a 1
    UserList(tIndex).Flags.Ban = 1
    
    If UserList(tIndex).Flags.Privilegios > 0 Then
            UserList(UserIndex).Flags.Ban = 1
            Call CloseSocket(UserIndex)
            Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & " banned by the server por bannear un Administrador." & FONTTYPE_FIGHT)
    End If
    
    Call LogGM(UserList(UserIndex).Name, "Echo a " & UserList(tIndex).Name)
    Call LogGM(UserList(UserIndex).Name, "BAN a " & UserList(tIndex).Name)
    Call CloseSocket(tIndex)
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/UNBAN " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    
    If FileExist(CharPath & UCase(rdata) & ".chr", vbNormal) = False Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe.")
        Exit Sub
    End If
    
    Call UnBan(rdata)
    Call LogGM(UserList(UserIndex).Name, "/UNBAN a " & rdata)
    Call SendData(ToIndex, UserIndex, 0, "||" & rdata & " unbanned." & FONTTYPE_INFO)
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
        tIndex = NameIndex(Name)
    Else
        tIndex = UserIndex
    End If
    X = val(ReadField(3, rdata, 32))
    Y = val(ReadField(4, rdata, 32))
    If Not InMapBounds(mapa, X, Y) Then Exit Sub
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    '[KEVIN]
    UserList(tIndex).Flags.EnTorneo = 0
    '[/KEVIN]
    Call WarpUserChar(tIndex, mapa, X, Y, True)
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " transportado." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "Transporto a " & UserList(tIndex).Name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y)
    Exit Sub
End If

'[KEVIN]
'Ver logs de sh en la consola
If UCase$(Left$(rdata, 6)) = "/VERSH" Then
    rdata = Right$(rdata, Len(rdata) - 6)

    If UserList(UserIndex).Flags.VerSH = False Then
        UserList(UserIndex).Flags.VerSH = True
        Call SendData(ToIndex, UserIndex, 0, "||VerSh ON." & FONTTYPE_CELESTE)
    Else
        UserList(UserIndex).Flags.VerSH = False
        Call SendData(ToIndex, UserIndex, 0, "||VerSh OFF." & FONTTYPE_CELESTE)
    End If

    Exit Sub
End If
'[/KEVIN]

'SEGUIR
If UCase$(rdata) = "/SEGUIR" Then
    If UserList(UserIndex).Flags.TargetNpc > 0 Then
        Call DoFollow(UserList(UserIndex).Flags.TargetNpc, UserList(UserIndex).Name)
    End If
    Exit Sub
End If

'IR A
If UCase$(Left$(rdata, 5)) = "/IRA " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    

    Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y + 1, True)
    
    
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " se ha trasportado hacia donde te encontras." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "/IRA " & UserList(tIndex).Name & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y)
    Exit Sub
End If

If UCase$(rdata) = "/TELEPLOC" Then
    Call WarpUserChar(UserIndex, UserList(UserIndex).Flags.TargetMap, UserList(UserIndex).Flags.TargetX, UserList(UserIndex).Flags.TargetY, True)
    Call LogGM(UserList(UserIndex).Name, "/TELEPLOC a x:" & UserList(UserIndex).Flags.TargetX & " Y:" & UserList(UserIndex).Flags.TargetY & " Map:" & UserList(UserIndex).Pos.Map)
    Exit Sub
End If

'Summon
If UCase$(Left$(rdata, 5)) = "/SUM " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " há sido trasportado." & FONTTYPE_INFO)
    Call WarpUserChar(tIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, True)
    
    Call LogGM(UserList(UserIndex).Name, "/SUM " & UserList(tIndex).Name & " Map:" & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.X & " Y:" & UserList(UserIndex).Pos.Y)
    Exit Sub
End If

'[KEVIN]
'Summon pa Torneo
If UCase$(Left$(rdata, 8)) = "/SUMTNO " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    
    tIndex = NameIndex(rdata)
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||El jugador no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " há sido trasportado." & FONTTYPE_INFO)
    Call WarpUserChar(tIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, True)
    
    UserList(tIndex).Flags.EnTorneo = 1
    
    '<<<< Paralisis >>>>
    If UserList(tIndex).Flags.Paralizado = 1 Then
        UserList(tIndex).Flags.Paralizado = 0
        Call SendData(ToIndex, tIndex, 0, "PARADOK")
    End If

    '<<< Ceguera >>>
    If UserList(tIndex).Flags.Ceguera = 1 Then
        UserList(tIndex).Flags.Ceguera = 0
        Call SendData(ToIndex, tIndex, 0, "NSEGUE")
    End If

    '<<<< Estupidez >>>>
    If UserList(tIndex).Flags.Estupidez = 1 Then
        UserList(tIndex).Flags.Estupidez = 0
        Call SendData(ToIndex, tIndex, 0, "NESTUP")
    End If

    '<<<< Invisibilidad >>>>
    If UserList(tIndex).Flags.Invisible = 1 Then
        UserList(tIndex).Counters.Invisibilidad = 0
        UserList(tIndex).Flags.Invisible = 0
        Call SendData(ToMap, 0, UserList(tIndex).Pos.Map, "KLBPA" & UserList(tIndex).Char.CharIndex & ",0")
    End If

    Call LogGM(UserList(UserIndex).Name, "/SUMTNO " & UserList(UserIndex).Name & " Map:" & UserList(UserIndex).Pos.Map & " X:" & UserList(UserIndex).Pos.X & " Y:" & UserList(UserIndex).Pos.Y)
    Exit Sub
End If

'Pone el flag de torneo en 0 de todos los usuarios
If UCase$(Left$(rdata, 7)) = "/RFLAGS" Then
    rdata = Right$(rdata, Len(rdata) - 7)
    
    Dim LoopU As Integer
    For LoopU = 1 To LastUser
        If UserList(LoopU).Flags.EnTorneo = 1 Then UserList(LoopU).Flags.EnTorneo = 0
    Next LoopU
    
    MinLevel = 1
    
    Exit Sub
End If
'[/KEVIN]

'Crear criatura
If UCase$(Left$(rdata, 3)) = "/CC" Then
   Call EnviarSpawnList(UserIndex)
   Exit Sub
End If

'Spawn!!!!!
If UCase$(Left$(rdata, 3)) = "SPA" Then
    rdata = Right$(rdata, Len(rdata) - 3)
    
    If val(rdata) > 0 And val(rdata) < UBound(SpawnList) + 1 Then _
          Call SpawnNpc(SpawnList(val(rdata)).NpcIndex, UserList(UserIndex).Pos, True, False)
          
          Call LogGM(UserList(UserIndex).Name, "Sumoneo " & SpawnList(val(rdata)).NpcName)
          
    Exit Sub
End If

'Haceme invisible vieja!
If UCase$(rdata) = "/INVISIBLE" Then
    Call DoAdminInvisible(UserIndex)
    Call LogGM(UserList(UserIndex).Name, "/INVISIBLE")
    Exit Sub
End If

If UCase(rdata) = "/LLUVIA" Then
    Lloviendo = Not Lloviendo
    Call SendData(ToAll, 0, 0, "LLU")
    Exit Sub
End If

'Resetea el inventario
If UCase$(rdata) = "/RESETINV" Then
    rdata = Right$(rdata, Len(rdata) - 9)
    If UserList(UserIndex).Flags.TargetNpc = 0 Then Exit Sub
    Call ResetNpcInv(UserList(UserIndex).Flags.TargetNpc)
    Call LogGM(UserList(UserIndex).Name, "/RESETINV " & Npclist(UserList(UserIndex).Flags.TargetNpc).Name)
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
    Call LogGM(UserList(UserIndex).Name, "Mensaje Broadcast:" & rdata)
    If rdata <> "" Then
        Call SendData(ToAll, 0, 0, "||" & UserList(UserIndex).Name & ": " & rdata & FONTTYPE_TALK & ENDC)
    End If
    Exit Sub
End If

'[KEVIN]
If UCase$(Left$(rdata, 9)) = "/SERVMSG " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    Call LogGM(UserList(UserIndex).Name, "Mensaje al Servidor:" & rdata)
    If rdata <> "" Then
        frmMain.rtfGmMsg.SelText = UserList(UserIndex).Name & ": " & rdata & vbCrLf
        Call SendData(ToIndex, UserIndex, 0, "||Has enviado un mensaje al server" & FONTTYPE_INFO)
    End If
    Exit Sub
End If
'[/KEVIN]

'[KEVIN]
If UCase$(Left$(rdata, 6)) = "/GMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(UserIndex).Name, "Mensaje a los GMS:" & rdata)
    If rdata <> "" Then
        Call SendData(ToAdmins, 0, 0, "||" & UserList(UserIndex).Name & ": " & rdata & FONTTYPE_ORANGE & ENDC)
    End If
    Exit Sub
End If
'[/KEVIN]

If UCase(Left$(rdata, 9)) = "/DOBACKUP" Then
    Call DoBackUp
    Exit Sub
End If

'Ip del nick
If UCase$(Left$(rdata, 8)) = "/IPNICK " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    tIndex = NameIndex(UCase(rdata))
    If tIndex > 0 Then
       Call SendData(ToIndex, UserIndex, 0, "||El ip de " & rdata & " es " & UserList(tIndex).ip & FONTTYPE_INFO)
    End If
    Exit Sub
End If

'[KEVIN]
If UCase$(Left$(rdata, 8)) = "$USKILL " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    
    tIndex = NameIndex(UCase(rdata))
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(ToIndex, UserIndex, 0, "PNSKILL" & UserList(tIndex).Stats.UserSkills(1) & "," & UserList(tIndex).Stats.UserSkills(2) & "," & UserList(tIndex).Stats.UserSkills(3) & "," & UserList(tIndex).Stats.UserSkills(4) & "," & UserList(tIndex).Stats.UserSkills(5) & "," & UserList(tIndex).Stats.UserSkills(6) & "," & UserList(tIndex).Stats.UserSkills(7) & "," & UserList(tIndex).Stats.UserSkills(8) & "," & UserList(tIndex).Stats.UserSkills(9) & "," & UserList(tIndex).Stats.UserSkills(10) & "," & UserList(tIndex).Stats.UserSkills(11) & "," & UserList(tIndex).Stats.UserSkills(12) & "," & UserList(tIndex).Stats.UserSkills(13) & "," & UserList(tIndex).Stats.UserSkills(14) & "," & UserList(tIndex).Stats.UserSkills(15) & "," & UserList(tIndex).Stats.UserSkills(16) & "," & UserList(tIndex).Stats.UserSkills(17) & "," & UserList(tIndex).Stats.UserSkills(18) & "," & UserList(tIndex).Stats.UserSkills(19) & "," & UserList(tIndex).Stats.UserSkills(20) & "," & UserList(tIndex).Stats.UserSkills(21) _
    & "," & UserList(tIndex).Stats.UserSkills(22) & "," & UserList(tIndex).Stats.UserSkills(23))
    
    Exit Sub
End If

If UCase(rdata) = "/TOOL" Then

Call SendData(ToIndex, UserIndex, 0, "HGM")
    
Exit Sub
End If
'[/KEVIN]

'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<< Dioses >>>>>>>>>>>>>>>>>>>>>>>>
If UserList(UserIndex).Flags.Privilegios < 3 Then
    Exit Sub
End If

'Destruir
If UCase$(Left$(rdata, 5)) = "/DEST" Then
    Call LogGM(UserList(UserIndex).Name, "/DEST")
    rdata = Right$(rdata, Len(rdata) - 5)
    Call EraseObj(ToMap, UserIndex, UserList(UserIndex).Pos.Map, 10000, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
    Exit Sub
End If

'Bloquear
If UCase$(Left$(rdata, 5)) = "/BLOQ" Then
    Call LogGM(UserList(UserIndex).Name, "/BLOQ")
    rdata = Right$(rdata, Len(rdata) - 5)
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 0 Then
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 1
        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 1)
    Else
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked = 0
        Call Bloquear(ToMap, UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, 0)
    End If
    Exit Sub
End If

'Quitar NPC
If UCase$(rdata) = "/MATA" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    If UserList(UserIndex).Flags.TargetNpc = 0 Then Exit Sub
    Call QuitarNPC(UserList(UserIndex).Flags.TargetNpc)
    Call LogGM(UserList(UserIndex).Name, "/MATA " & Npclist(UserList(UserIndex).Flags.TargetNpc).Name)
    Exit Sub
End If

'Quita todos los NPCs del area
If UCase$(rdata) = "/MASSKILL" Then
    For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
            For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                    If MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(UserList(UserIndex).Pos.Map, X, Y).NpcIndex)
            Next X
    Next Y
    Call LogGM(UserList(UserIndex).Name, "/MASSKILL")
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
    Call LogGM(UserList(UserIndex).Name, "Mensaje de sistema:" & rdata)
    Call SendData(ToAll, 0, 0, "!!" & UserList(UserIndex).Name & ": " & rdata & ENDC)
    
    Exit Sub
End If

'Crear criatura, toma directamente el indice
If UCase$(Left$(rdata, 5)) = "/ACC " Then
   rdata = Right$(rdata, Len(rdata) - 5)
   Call SpawnNpc(val(rdata), UserList(UserIndex).Pos, True, False)
   Exit Sub
End If

'Crear criatura con respawn, toma directamente el indice
If UCase$(Left$(rdata, 6)) = "/RACC " Then
   rdata = Right$(rdata, Len(rdata) - 6)
   Call SpawnNpc(val(rdata), UserList(UserIndex).Pos, True, True)
   Exit Sub
End If

If UCase$(rdata) = "/GUARDARPERS" Then
    Call GrabarInfoUsuarios
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
    If UserList(UserIndex).Flags.Navegando = 1 Then
        UserList(UserIndex).Flags.Navegando = 0
    Else
        UserList(UserIndex).Flags.Navegando = 1
    End If
    Exit Sub
End If

'Apagamos
If UCase$(rdata) = "/APAGAR" Then
    If UCase(UserList(UserIndex).Name) <> "NEB" Then
        Call LogGM(UserList(UserIndex).Name, "¡¡¡Intento apagar el server!!!")
        Exit Sub
    End If
    'Log
    mifile = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #mifile
    Print #mifile, Date & " " & Time & " server apagado por " & UserList(UserIndex).Name & ". "
    Close #mifile
    Unload frmMain
    Exit Sub
End If

'CONDENA
If UCase$(Left$(rdata, 8)) = "/CONDEN " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    tIndex = NameIndex(rdata)
    If tIndex > 0 Then Call VolverCriminal(tIndex)
    Exit Sub
End If

'[KEVIN]
If UCase$(Left$(rdata, 7)) = "/RAJAR " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(UCase(rdata))
    If tIndex > 0 Then
        If UserList(tIndex).Faccion.ArmadaReal > 0 Then
            Call ExpulsarFaccionReal(tIndex)
        ElseIf UserList(tIndex).Faccion.FuerzasCaos > 0 Then
            Call ExpulsarCaos(tIndex)
        End If
    End If
    Exit Sub
End If
'[/KEVIN]

'[KEVIN]
If UCase$(Left$(rdata, 10)) = "/CREAROBJ " Then
    rdata = Right$(rdata, Len(rdata) - 10)
    Dim ObjC As Obj
    
    ObjC.ObjIndex = ReadField(1, rdata, 32)
    ObjC.Amount = ReadField(2, rdata, 32)
    
    MeterItemEnInventario UserIndex, ObjC
    
    Call LogGM(UserList(UserIndex).Name, "Creó " & ObjC.Amount & " unidades del objeto " & ObjC.ObjIndex)
    Exit Sub
End If
'[/KEVIN]


'[KEVIN]
If UCase$(rdata) = "/BOOTALL" Then
    
    Dim BootAll As Integer
    
    For BootAll = 1 To LastUser
        If BootAll <> UserIndex Then
            CloseSocket (BootAll)
        End If
    Next BootAll
    
    Call LogGM(UserList(UserIndex).Name, "¡¡¡ECHO A TODOS!!!")
    
    Exit Sub
End If
'[/KEVIN]

'[KEVIN]MODIFICADO
'MODIFICA CARACTER
If UCase(Left$(rdata, 5)) = "/MOD " Then
    Call LogGM(UserList(UserIndex).Name, rdata)
    rdata = Right$(rdata, Len(rdata) - 5)
    tIndex = NameIndex(ReadField(1, rdata, 32))
    Arg1 = ReadField(2, rdata, 32)
    Arg2 = ReadField(3, rdata, 32)
    Arg3 = ReadField(4, rdata, 32)
    Arg4 = ReadField(5, rdata, 32)
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
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
            Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, val(Arg2), UserList(tIndex).Char.Head, UserList(tIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            Exit Sub
        Case "HEAD"
            Call ChangeUserChar(ToMap, 0, UserList(tIndex).Pos.Map, tIndex, UserList(tIndex).Char.Body, val(Arg2), UserList(tIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
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
            Call SendData(ToIndex, UserIndex, 0, "||Comando no permitido." & FONTTYPE_INFO)
            Exit Sub
    End Select

    Exit Sub
End If
'[/KEVIN]

If UCase(Left$(rdata, 9)) = "/DOBACKUP" Then
    Call DoBackUp
    Exit Sub
End If


If UCase(Left$(rdata, 11)) = "/BORRAR SOS" Then
    Call Ayuda.Reset
    Exit Sub
End If

'[KEVIN]
If UCase(Left$(rdata, 11)) = "/BORRAR INS" Then
    Call InscTorneo.Reset
    Exit Sub
End If

If UCase(Left$(rdata, 9)) = "/SHOW INT" Then
    Call frmMain.mnuMostrar_Click
    Exit Sub
End If

If UCase(rdata) = "/PASSDAY" Then
    Call DayElapsed
    Exit Sub
End If

'[KEVIN]
If UCase$(Left$(rdata, 7)) = "/CURAR " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    tIndex = NameIndex(UCase(rdata))
    
    If tIndex <= 0 Then
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
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
        Call SendData(ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(ToIndex, UserIndex, 0, "PNSKILL" & UserList(tIndex).Stats.UserSkills(1) & "," & UserList(tIndex).Stats.UserSkills(2) & "," & UserList(tIndex).Stats.UserSkills(3) & "," & UserList(tIndex).Stats.UserSkills(4) & "," & UserList(tIndex).Stats.UserSkills(5) & "," & UserList(tIndex).Stats.UserSkills(6) & "," & UserList(tIndex).Stats.UserSkills(7) & "," & UserList(tIndex).Stats.UserSkills(8) & "," & UserList(tIndex).Stats.UserSkills(9) & "," & UserList(tIndex).Stats.UserSkills(10) & "," & UserList(tIndex).Stats.UserSkills(11) & "," & UserList(tIndex).Stats.UserSkills(12) & "," & UserList(tIndex).Stats.UserSkills(13) & "," & UserList(tIndex).Stats.UserSkills(14) & "," & UserList(tIndex).Stats.UserSkills(15) & "," & UserList(tIndex).Stats.UserSkills(16) & "," & UserList(tIndex).Stats.UserSkills(17) & "," & UserList(tIndex).Stats.UserSkills(18) & "," & UserList(tIndex).Stats.UserSkills(19) & "," & UserList(tIndex).Stats.UserSkills(20) & "," & UserList(tIndex).Stats.UserSkills(21) _
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
        Call SendData(ToIndex, UserIndex, 0, "||Debes colocar un número menor a 5000!!")
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
If UCase(rdata) = "/GMTOOL" Then

Call SendData(ToIndex, UserIndex, 0, "GMT")
    
Exit Sub
End If

'Agregar un Dios
If UCase$(Left$(rdata, 13)) = "/AGREGARDIOS " Then
    rdata = Right$(rdata, Len(rdata) - 13)
    Dim Temp, GmNum, NumGMs As Integer
    Arg1 = rdata
    
    If UCase(UserList(UserIndex).Name) <> "NEB" Then Exit Sub
    
    If FileExist(CharPath & UCase(Arg1) & ".chr", vbNormal) = False Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe.")
        Exit Sub
    End If
    
    NumGMs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
    For GmNum = 1 To NumGMs
        If UCase(Arg1) = UCase(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & GmNum)) Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario ya es dios.")
            Exit Sub
        End If
    Next GmNum
    
    Temp = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses")) + 1
    Call WriteVar(IniPath & "Server.ini", "INIT", "Dioses", Str(Temp))
    Call WriteVar(IniPath & "Server.ini", "Dioses", "Dios" & Temp, Arg1)
    Call SendData(ToIndex, UserIndex, 0, "||Has agregado a " & Arg1 & " a la lista de los Dioses." & FONTTYPE_INFO)
    
    tIndex = NameIndex(Arg1)
    If tIndex > 0 Then
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " te agregó a la lista de los Dioses." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "Agregó a " & rdata & " a la lista de los Dioses")
    End If
    Exit Sub
End If

'Agregar un Semi Dios
If UCase$(Left$(rdata, 13)) = "/AGREGARSEMI " Then
    rdata = Right$(rdata, Len(rdata) - 13)
    Dim NumTemp, GmsNum, NumSGMs As Integer
    Arg1 = rdata
    
    If UCase(UserList(UserIndex).Name) <> "NEB" Then Exit Sub
    
    If FileExist(CharPath & UCase(Arg1) & ".chr", vbNormal) = False Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe.")
        Exit Sub
    End If
    
    NumSGMs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
    For GmsNum = 1 To NumSGMs
        If UCase(Arg1) = UCase(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & GmsNum)) Then
            Call SendData(ToIndex, UserIndex, 0, "||El usuario ya es SemiDios.")
            Exit Sub
        End If
    Next GmsNum
    
    NumTemp = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses")) + 1
    Call WriteVar(IniPath & "Server.ini", "INIT", "SemiDioses", Str(NumTemp))
    Call WriteVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & NumTemp, Arg1)
    Call SendData(ToIndex, UserIndex, 0, "||Has agregado a " & Arg1 & " a la lista de los SemiDioses." & FONTTYPE_INFO)
    
    tIndex = NameIndex(Arg1)
    If tIndex > 0 Then
    Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " te agregó a la lista de los SemiDioses." & FONTTYPE_INFO)
    Call LogGM(UserList(UserIndex).Name, "Agregó a " & rdata & " a la lista de los SemiDioses")
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
            Call SendData(ToIndex, UserIndex, 0, "||Has removido a " & Arg1 & " de la lista de Dioses." & FONTTYPE_INFO)
            tIndex = NameIndex(Arg1)
            If tIndex > 0 Then
                Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " te a removido de la lista de Dioses." & FONTTYPE_INFO)
                'Call modName(tIndex)
            End If
            Call LogGM(UserList(UserIndex).Name, "Echó a " & rdata & " de la lista de los Dioses")
            Exit Sub
        End If
    Next GmGNum
    Call SendData(ToIndex, UserIndex, 0, "||" & Arg1 & " no fue encontrado en la lista de los dioses." & FONTTYPE_INFO)
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
            Call SendData(ToIndex, UserIndex, 0, "||Has removido a " & Arg1 & " de la lista de Dioses." & FONTTYPE_INFO)
            tIndex = NameIndex(Arg1)
            If tIndex > 0 Then
                Call SendData(ToIndex, tIndex, 0, "||" & UserList(UserIndex).Name & " te a removido de la lista de SemiDioses." & FONTTYPE_INFO)
                'Call modName(tIndex)
            End If
            Call LogGM(UserList(UserIndex).Name, "Echó a " & rdata & " de la lista de los SemiDioses")
            Exit Sub
        End If
    Next GmNum2
    Call SendData(ToIndex, UserIndex, 0, "||" & Arg1 & " no fue encontrado en la lista de los SemiDioses." & FONTTYPE_INFO)
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
    
    For Y = UserList(UserIndex).Pos.Y - MinYBorder + 1 To UserList(UserIndex).Pos.Y + MinYBorder - 1
            For X = UserList(UserIndex).Pos.X - MinXBorder + 1 To UserList(UserIndex).Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                    If MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex > 0 Then Call WarpUserChar(MapData(UserList(UserIndex).Pos.Map, X, Y).UserIndex, mapa, X2, Y2, True)
            Next X
    Next Y
    
    Call LogGM(UserList(UserIndex).Name, "Hizo un teleport en masa al " & "Mapa" & mapa & " X:" & X2 & " Y:" & Y2)
    Exit Sub

End If

'Destrabar Usuario
If UCase$(Left$(rdata, 11)) = "/DESTRABAR " Then
    rdata = Right$(rdata, Len(rdata) - 11)
    mapa = val(ReadField(2, rdata, Asc("-")))
    
    If Not MapaValido(mapa) Then
        Call SendData(ToIndex, UserIndex, 0, "||Mapa Incorrecto.")
        Exit Sub
    End If
    
    Name = ReadField(1, rdata, 45)
    X = val(ReadField(3, rdata, 45))
    Y = val(ReadField(4, rdata, 45))
    
    If FileExist(CharPath & UCase(Name) & ".chr", vbNormal) = False Then
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe.")
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
        Call SendData(ToIndex, UserIndex, 0, "||El usuario no existe.")
        Exit Sub
    End If
    
    Call WriteVar(CharPath & UCase(Nme) & ".chr", "INIT", "Password", MD5String(Pasd))
    Call LogGM(UserList(UserIndex).Name, "Le cambio el password a " & Nme)
    Exit Sub
End If

'[/KEVIN]

Exit Sub

errorhandler:
 Call LogError("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(UserIndex).Name & "UI:" & UserIndex)
 Call CloseSocket(UserIndex)

End Sub

Sub ReloadSokcet()

On Error GoTo errhandler

    frmMain.Socket1.Cleanup
    Call ConfigListeningSocket(frmMain.Socket1, puerto)


Exit Sub
errhandler:
    Call LogError("Error en CheckSocketState," & Err.Description)

End Sub

Public Function CheckArmaduraOculto(ByVal UserIndex As Integer) As Boolean

Dim LoopC As Integer

For LoopC = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(LoopC).ObjIndex = OCULTAR_RAZAALTA Or _
    UserList(UserIndex).Invent.Object(LoopC).ObjIndex = OCULTAR_RAZABAJA Then
        If UserList(UserIndex).Invent.Object(LoopC).Equipped = 1 Then
            CheckArmaduraOculto = True
            Exit Function
        End If
    End If
Next LoopC

CheckArmaduraOculto = False
        
End Function
