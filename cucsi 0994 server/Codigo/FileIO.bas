Attribute VB_Name = "ES"
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

Public Sub CargarSpawnList()

    Dim N As Integer, LoopC As Integer
    N = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(N) As tCriaturasEntrenador
    For LoopC = 1 To N
        SpawnList(LoopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
        SpawnList(LoopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC


End Sub

Function EsDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
For WizNum = 1 To NumWizs
    If UCase(Name) = UCase(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum)) Then
        EsDios = True
        Exit Function
    End If
Next WizNum
EsDios = False
End Function

Function EsSemiDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
For WizNum = 1 To NumWizs
    If UCase(Name) = UCase(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum)) Then
        EsSemiDios = True
        Exit Function
    End If
Next WizNum
EsSemiDios = False
End Function

Function EsConsejero(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))
For WizNum = 1 To NumWizs
    If UCase(Name) = UCase(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum)) Then
        EsConsejero = True
        Exit Function
    End If
Next WizNum
EsConsejero = False
End Function
Public Function TxtDimension(ByVal Name As String) As Long
Dim N As Integer, cad As String, Tam As Long
N = FreeFile(1)
Open Name For Input As #N
Tam = 0
Do While Not EOF(N)
    Tam = Tam + 1
    Line Input #N, cad
Loop
Close N
TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()
ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
Dim N As Integer, i As Integer
N = FreeFile(1)
Open DatPath & "NombresInvalidos.txt" For Input As #N

For i = 1 To UBound(ForbidenNames)
    Line Input #N, ForbidenNames(i)
Next i

Close N

End Sub
Public Sub CargarHechizos()
On Error GoTo errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."

Dim Hechizo As Integer

'obtiene el numero de hechizos
NumeroHechizos = val(GetVar(DatPath & "Hechizos.dat", "INIT", "NumeroHechizos"))
ReDim Hechizos(1 To NumeroHechizos) As tHechizo

frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumeroHechizos
frmCargando.cargar.Value = 0

'Llena la lista
For Hechizo = 1 To NumeroHechizos

    Hechizos(Hechizo).Nombre = GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Nombre")
    Hechizos(Hechizo).Desc = GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Desc")
    Hechizos(Hechizo).PalabrasMagicas = GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "PalabrasMagicas")
    
    Hechizos(Hechizo).HechizeroMsg = GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "HechizeroMsg")
    Hechizos(Hechizo).TargetMsg = GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "TargetMsg")
    Hechizos(Hechizo).PropioMsg = GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "PropioMsg")
    
    Hechizos(Hechizo).Tipo = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Tipo"))
    Hechizos(Hechizo).WAV = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "WAV"))
    Hechizos(Hechizo).FXgrh = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Fxgrh"))
    
    Hechizos(Hechizo).loops = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Loops"))
    
    Hechizos(Hechizo).Resis = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Resis"))
    
    Hechizos(Hechizo).SubeHP = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeHP"))
    Hechizos(Hechizo).MinHP = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinHP"))
    Hechizos(Hechizo).MaxHP = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxHP"))
    
    Hechizos(Hechizo).SubeMana = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeMana"))
    Hechizos(Hechizo).MiMana = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinMana"))
    Hechizos(Hechizo).MaMana = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxMana"))
    
    Hechizos(Hechizo).SubeSta = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeSta"))
    Hechizos(Hechizo).MinSta = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinSta"))
    Hechizos(Hechizo).MaxSta = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxSta"))
    
    Hechizos(Hechizo).SubeHam = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeHam"))
    Hechizos(Hechizo).MinHam = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinHam"))
    Hechizos(Hechizo).MaxHam = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxHam"))
    
    Hechizos(Hechizo).SubeSed = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeSed"))
    Hechizos(Hechizo).MinSed = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinSed"))
    Hechizos(Hechizo).MaxSed = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxSed"))
    
    Hechizos(Hechizo).SubeAgilidad = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeAG"))
    Hechizos(Hechizo).MinAgilidad = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinAG"))
    Hechizos(Hechizo).MaxAgilidad = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxAG"))
    
    Hechizos(Hechizo).SubeFuerza = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeFU"))
    Hechizos(Hechizo).MinFuerza = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinFU"))
    Hechizos(Hechizo).MaxFuerza = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxFU"))
    
    Hechizos(Hechizo).SubeCarisma = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "SubeCA"))
    Hechizos(Hechizo).MinCarisma = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinCA"))
    Hechizos(Hechizo).MaxCarisma = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MaxCA"))
    
    
    Hechizos(Hechizo).Invisibilidad = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Invisibilidad"))
    Hechizos(Hechizo).Paraliza = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Paraliza"))
    Hechizos(Hechizo).RemoverParalisis = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "RemoverParalisis"))
    
    Hechizos(Hechizo).CuraVeneno = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "CuraVeneno"))
    Hechizos(Hechizo).Envenena = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Envenena"))
    Hechizos(Hechizo).Maldicion = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Maldicion"))
    Hechizos(Hechizo).RemoverMaldicion = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "RemoverMaldicion"))
    Hechizos(Hechizo).Bendicion = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Bendicion"))
    Hechizos(Hechizo).Revivir = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Revivir"))
    
    Hechizos(Hechizo).Ceguera = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Ceguera"))
    Hechizos(Hechizo).Estupidez = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Estupidez"))
    
    Hechizos(Hechizo).Invoca = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Invoca"))
    Hechizos(Hechizo).NumNpc = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "NumNpc"))
    Hechizos(Hechizo).Cant = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Cant"))
    
    
    Hechizos(Hechizo).Materializa = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Materializa"))
    Hechizos(Hechizo).ItemIndex = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "ItemIndex"))
    
    Hechizos(Hechizo).MinSkill = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "MinSkill"))
    Hechizos(Hechizo).ManaRequerido = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "ManaRequerido"))
    
    Hechizos(Hechizo).Target = val(GetVar(DatPath & "hechizos.dat", "Hechizo" & Hechizo, "Target"))
    frmCargando.cargar.Value = frmCargando.cargar.Value + 1
Next
Exit Sub

errhandler:
 MsgBox "Error cargando hechizos.dat"
End Sub
'[KEVIN]************************************************************
'***************************************************************************
'***************************************************************************
Public Sub CargarQuests()
On Error GoTo errhandler

Dim Quest As Integer

'obtiene el numero de quest
NumeroQuests = val(GetVar(DatPath & "Quests.dat", "INIT", "NumQuests"))
ReDim Quests(1 To NumeroQuests) As tQuest

'Llena la lista
For Quest = 1 To NumeroQuests

    Quests(Quest).Coordenadas1 = GetVar(DatPath & "Quests.dat", "Quest" & Quest, "Coordenadas1")
    Quests(Quest).Coordenadas2 = GetVar(DatPath & "Quests.dat", "Quest" & Quest, "Coordenadas2")
    Quests(Quest).Coordenadas3 = GetVar(DatPath & "Quests.dat", "Quest" & Quest, "Coordenadas3")
    Quests(Quest).Coordenadas4 = GetVar(DatPath & "Quests.dat", "Quest" & Quest, "Coordenadas4")
    Quests(Quest).Coordenadas5 = GetVar(DatPath & "Quests.dat", "Quest" & Quest, "Coordenadas5")
    Quests(Quest).CriaturaIndex = val(GetVar(DatPath & "Quests.dat", "Quest" & Quest, "CriaturaIndex"))
    Quests(Quest).DaExp = val(GetVar(DatPath & "Quests.dat", "Quest" & Quest, "DaExp"))
    Quests(Quest).DaObj = val(GetVar(DatPath & "Quests.dat", "Quest" & Quest, "DaObj"))
    Quests(Quest).DaOro = val(GetVar(DatPath & "Quests.dat", "Quest" & Quest, "DaOro"))
    Quests(Quest).Exp = val(GetVar(DatPath & "Quests.dat", "Quest" & Quest, "Exp"))
    Quests(Quest).Obj = val(GetVar(DatPath & "Quests.dat", "Quest" & Quest, "Obj"))
    Quests(Quest).Oro = val(GetVar(DatPath & "Quests.dat", "Quest" & Quest, "Oro"))
    Quests(Quest).Objetivo = val(GetVar(DatPath & "Quests.dat", "Quest" & Quest, "Objetivo"))
    Quests(Quest).NPCs = val(GetVar(DatPath & "Quests.dat", "Quest" & Quest, "Npcs"))
    Quests(Quest).UsUaRiOs = val(GetVar(DatPath & "Quests.dat", "Quest" & Quest, "Usuarios"))
    Quests(Quest).Ciudadanos = val(GetVar(DatPath & "Quests.dat", "Quest" & Quest, "Ciudadanos"))
    Quests(Quest).Criminales = val(GetVar(DatPath & "Quests.dat", "Quest" & Quest, "Criminales"))
    Quests(Quest).AmigoNpc = val(GetVar(DatPath & "Quests.dat", "Quest" & Quest, "AmigoNpc"))
    
Next
Exit Sub

errhandler:
 MsgBox "Error cargando Quests.dat"
End Sub
'[/KEVIN]

'[KEVIN]************************************************************
'***************************************************************************
'***************************************************************************
Public Sub CargarCofres()
On Error GoTo errhandler

Dim Cofre As Integer
Dim LoopC As Integer
Dim ln As String

'obtiene el numero de quest
NumeroCofres = val(GetVar(DatPath & "Cofres.dat", "INIT", "NumCofres"))
ReDim Cofres(1 To NumeroCofres) As CofreInventario

'Llena la lista
For Cofre = 1 To NumeroCofres

Cofres(Cofre).NroItems = val(GetVar(DatPath & "Cofres.dat", "Cofre" & Cofre, "NroItems"))

'Lista de objetos del cofre
For LoopC = 1 To MAX_COFREINVENTORY_SLOTS
    ln = GetVar(DatPath & "Cofres.dat", "Cofre" & Cofre, "Obj" & LoopC)
    Cofres(Cofre).Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
    Cofres(Cofre).Object(LoopC).Amount = val(ReadField(2, ln, 45))
Next LoopC
    
Next
Exit Sub

errhandler:
 MsgBox "Error cargando Cofres.dat"
End Sub
'[/KEVIN]

Public Sub DoBackUp()
'Call LogTarea("Sub DoBackUp")
haciendoBK = True


Call SendData(ToAll, 0, 0, "BKW")

Call SaveGuildsDB
Call LimpiarMundo
Call GrabarInfoUsuarios
'[KEVIN]
Call SaveCofres
'[/KEVIN]
Call WorldSave

Call SendData(ToAll, 0, 0, "BKW")

haciendoBK = False

End Sub


Public Sub SaveMapData(ByVal N As Integer)

'Call LogTarea("Sub SaveMapData N:" & n)

Dim LoopC As Integer
Dim TempInt As Integer
Dim Y As Integer
Dim X As Integer
Dim SaveAs As String

SaveAs = App.Path & "\WorldBackUP\Map" & N & ".map"

If FileExist(SaveAs, vbNormal) Then
    Kill SaveAs
End If

If FileExist(Left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) Then
    Kill Left$(SaveAs, Len(SaveAs) - 4) & ".inf"
End If

'Open .map file
Open SaveAs For Binary As #1
Seek #1, 1
SaveAs = Left$(SaveAs, Len(SaveAs) - 4)
SaveAs = SaveAs & ".inf"
'Open .inf file
Open SaveAs For Binary As #2
Seek #2, 1
'map Header
        
Put #1, , MapInfo(N).MapVersion
Put #1, , MiCabecera
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt
Put #1, , TempInt

'inf Header
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt
Put #2, , TempInt

'Write .map file
For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        
        '.map file
        Put #1, , MapData(N, X, Y).Blocked
        
        For LoopC = 1 To 4
            Put #1, , MapData(N, X, Y).Graphic(LoopC)
        Next LoopC
        
        'Lugar vacio para futuras expansiones
        Put #1, , MapData(N, X, Y).trigger
        
        Put #1, , TempInt
        
        '.inf file
        'Tile exit
        Put #2, , MapData(N, X, Y).TileExit.Map
        Put #2, , MapData(N, X, Y).TileExit.X
        Put #2, , MapData(N, X, Y).TileExit.Y
        
        'NPC
        If MapData(N, X, Y).NpcIndex > 0 Then
            Put #2, , Npclist(MapData(N, X, Y).NpcIndex).Numero
        Else
            Put #2, , 0
        End If
        'Object
        
        If MapData(N, X, Y).OBJInfo.ObjIndex > 0 Then
            If ObjData(MapData(N, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_FOGATA Then
                MapData(N, X, Y).OBJInfo.ObjIndex = 0
                MapData(N, X, Y).OBJInfo.Amount = 0
            End If
'            If ObjData(MapData(n, X, Y).OBJInfo.ObjIndex).ObjType = OBJTYPE_MANCHAS Then
'                MapData(n, X, Y).OBJInfo.ObjIndex = 0
'                MapData(n, X, Y).OBJInfo.Amount = 0
'            End If
        End If
        
        Put #2, , MapData(N, X, Y).OBJInfo.ObjIndex
        Put #2, , MapData(N, X, Y).OBJInfo.Amount
        
        'Empty place holders for future expansion
        Put #2, , TempInt
        Put #2, , TempInt
        
    Next X
Next Y

'Close .map file
Close #1

'Close .inf file
Close #2

'write .dat file
SaveAs = Left$(SaveAs, Len(SaveAs) - 4) & ".dat"
Call WriteVar(SaveAs, "Mapa" & N, "Name", MapInfo(N).Name)
Call WriteVar(SaveAs, "Mapa" & N, "MusicNum", MapInfo(N).Music)
Call WriteVar(SaveAs, "Mapa" & N, "StartPos", MapInfo(N).StartPos.Map & "-" & MapInfo(N).StartPos.X & "-" & MapInfo(N).StartPos.Y)

Call WriteVar(SaveAs, "Mapa" & N, "Terreno", MapInfo(N).Terreno)
Call WriteVar(SaveAs, "Mapa" & N, "Zona", MapInfo(N).Zona)
Call WriteVar(SaveAs, "Mapa" & N, "Restringir", MapInfo(N).Restringir)
Call WriteVar(SaveAs, "Mapa" & N, "BackUp", Str(MapInfo(N).BackUp))

If MapInfo(N).Pk Then
    Call WriteVar(SaveAs, "Mapa" & N, "pk", "0")
Else
    Call WriteVar(SaveAs, "Mapa" & N, "pk", "1")
End If

End Sub

Sub LoadArmasHerreria()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
Next lc


End Sub

Sub LoadArmadurasHerreria()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
Next lc

End Sub

Sub LoadObjCarpintero()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

ReDim Preserve ObjCarpintero(1 To N) As Integer

For lc = 1 To N
    ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
Next lc

End Sub
'[KEVIN]
'************************************************************************
'************************************************************************
Sub LoadObjDruida()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ObjDruida.dat", "INIT", "NumObjs"))

ReDim Preserve ObjDruida(1 To N) As Integer

For lc = 1 To N
    ObjDruida(lc) = val(GetVar(DatPath & "ObjDruida.dat", "Obj" & lc, "Index"))
Next lc
'[/KEVIN]
End Sub
'[KEVIN]
'************************************************************************
'************************************************************************
Sub LoadObjSastre()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ObjSastre.dat", "INIT", "NumObjs"))

ReDim Preserve ObjSastre(1 To N) As Integer

For lc = 1 To N
    ObjSastre(lc) = val(GetVar(DatPath & "ObjSastre.dat", "Obj" & lc, "Index"))
Next lc
'[/KEVIN]
End Sub

Sub LoadOBJData()

'Call LogTarea("Sub LoadOBJData")

On Error GoTo errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer

'obtiene el numero de obj
NumObjDatas = val(GetVar(DatPath & "Obj.dat", "INIT", "NumObjs"))

frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumObjDatas
frmCargando.cargar.Value = 0

FrmStat.ProgressBar1.Min = 0
FrmStat.ProgressBar1.max = NumObjDatas
FrmStat.ProgressBar1.Value = 0


ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
        
    ObjData(Object).Name = GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Name")
    
    ObjData(Object).GrhIndex = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "GrhIndex"))
    
    ObjData(Object).ObjType = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "ObjType"))
    ObjData(Object).SubTipo = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Subtipo"))
    
    ObjData(Object).Newbie = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Newbie"))
    
    If ObjData(Object).SubTipo = OBJTYPE_ESCUDO Then
        ObjData(Object).ShieldAnim = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Anim"))
        ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
        ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
        ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
        ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
    End If
    
    If ObjData(Object).SubTipo = OBJTYPE_CASCO Then
        ObjData(Object).CascoAnim = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Anim"))
        ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
        ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
        ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
        ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
    End If
    
    ObjData(Object).Ropaje = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "HechizoIndex"))
    
    If ObjData(Object).ObjType = OBJTYPE_WEAPON Then
            ObjData(Object).WeaponAnim = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Apuñala"))
            ObjData(Object).Envenena = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHIT"))
            ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Caos"))
            ObjData(Object).proyectil = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Municiones"))
            '[KEVIN]
            ObjData(Object).DosManos = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "DosManos"))
            ObjData(Object).DosArmas = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "DosArmas"))
            '[/KEVIN]
            
    End If
    
    If ObjData(Object).ObjType = OBJTYPE_ARMOUR Then
            ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Caos"))
    End If
    
    If ObjData(Object).ObjType = OBJTYPE_HERRAMIENTAS Then
            ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
    End If
    
    If ObjData(Object).ObjType = OBJTYPE_INSTRUMENTOS Then
        ObjData(Object).Snd1 = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SND1"))
        ObjData(Object).Snd2 = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SND2"))
        ObjData(Object).Snd3 = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SND3"))
        ObjData(Object).MinInt = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinInt"))
    End If
    
    ObjData(Object).LingoteIndex = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingoteIndex"))
    
    If ObjData(Object).ObjType = 31 Or ObjData(Object).ObjType = 23 Then
        ObjData(Object).MinSkill = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinSkill"))
    End If
    
    ObjData(Object).MineralIndex = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).MaxHP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHP"))
  
    
    ObjData(Object).Mujer = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Mujer"))
    ObjData(Object).Hombre = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Hombre"))
    
    ObjData(Object).MinHam = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHam"))
    ObjData(Object).MinSed = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinAgu"))
    
    
    ObjData(Object).MinDef = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MAXDEF"))
    
    '[KEVIN]
    ObjData(Object).Vendible = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Vendible"))
    ObjData(Object).DefMagia = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "DefMagia"))
    ObjData(Object).MagiaMinDef = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MAGIAMINDEF"))
    ObjData(Object).MagiaMaxDef = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MAGIAMAXDEF"))
    '[/KEVIN]
    
    ObjData(Object).Respawn = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "ReSpawn"))
    
    ObjData(Object).RazaEnana = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "RazaEnana"))
    
    '[KEVIN]
    ObjData(Object).RazaOrca = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "RazaOrca"))
    '[/KEVIN]
    
    ObjData(Object).Valor = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Valor"))
    
    '[KEVIN]
    ObjData(Object).Materiales = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Materiales"))
    ObjData(Object).Materiales2 = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Materiales2"))
    '[/KEVIN]
    
    ObjData(Object).Crucial = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Crucial"))
    
    ObjData(Object).Cerrada = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "abierta"))
    
    If ObjData(Object).Cerrada = 1 Then
            ObjData(Object).Llave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Llave"))
            ObjData(Object).Clave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Clave"))
    End If
    
    
    If ObjData(Object).ObjType = OBJTYPE_PUERTAS Or ObjData(Object).ObjType = OBJTYPE_BOTELLAVACIA Or ObjData(Object).ObjType = OBJTYPE_BOTELLALLENA Then
        ObjData(Object).IndexAbierta = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "IndexAbierta"))
        ObjData(Object).IndexCerrada = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "IndexCerrada"))
        ObjData(Object).IndexCerradaLlave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "IndexCerradaLlave"))
    End If
    
    '[KEVIN]
    If ObjData(Object).ObjType = OBJTYPE_CONTENEDORES Then
        ObjData(Object).IndexAbierta = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "IndexAbierta"))
        ObjData(Object).IndexCerradaLlave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "IndexCerradaLlave"))
        ObjData(Object).Llave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Llave"))
        ObjData(Object).Clave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Clave"))
        ObjData(Object).CofreNro = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "NroCofre"))
    End If
    '[/KEVIN]
    
    'Puertas y llaves
    ObjData(Object).Clave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Clave"))
    
    ObjData(Object).texto = GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Agarrable"))
    ObjData(Object).ForoID = GetVar(DatPath & "Obj.dat", "OBJ" & Object, "ID")
    
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        ObjData(Object).ClaseProhibida(i) = GetVar(DatPath & "Obj.dat", "OBJ" & Object, "CP" & i)
    Next
            
    ObjData(Object).Resistencia = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Resistencia"))
    
    'Pociones
    If ObjData(Object).ObjType = 11 Then
        ObjData(Object).TipoPocion = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "TipoPocion"))
        ObjData(Object).MaxModificador = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxModificador"))
        ObjData(Object).MinModificador = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinModificador"))
        ObjData(Object).DuracionEfecto = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "DuracionEfecto"))
    End If
    
    '[KEVIN]
    If ObjData(Object).ObjType = OBJTYPE_ANILLO Then
        ObjData(Object).TipoAnillo = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "TipoAnillo"))
        ObjData(Object).MaxModificador = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxModificador"))
        ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
        ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
        ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
        ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
    End If
    '[/KEVIN]

    ObjData(Object).SkCarpinteria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkCarpinteria"))
    '[KEVIN]
    ObjData(Object).SkPociones = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkPociones"))
    '[/KEVIN]
    
    '[KEVIN]
    ObjData(Object).SkSastreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkSastreria"))
    '[/KEVIN]
    
    If ObjData(Object).SkCarpinteria > 0 Then _
        ObjData(Object).Madera = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Madera"))
        
    '[KEVIN]
    If ObjData(Object).SkPociones > 0 Then _
        ObjData(Object).Raices = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Raices"))
    '[/KEVIN]
    
    '[KEVIN]
    If ObjData(Object).SkSastreria > 0 Then
        ObjData(Object).PielLobo = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "PielLobo"))
        ObjData(Object).PielOsoPardo = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "PielOsoPardo"))
        ObjData(Object).PielOsoPolar = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "PielOsoPolar"))
    End If
    '[/KEVIN]
    
    If ObjData(Object).ObjType = OBJTYPE_BARCOS Then
            ObjData(Object).MaxHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHIT"))
    End If
    
    If ObjData(Object).ObjType = OBJTYPE_FLECHAS Then
            ObjData(Object).MaxHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHIT"))
    End If
    
    'Bebidas
    ObjData(Object).MinSta = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinST"))
    
    frmCargando.cargar.Value = frmCargando.cargar.Value + 1
    FrmStat.ProgressBar1.Value = FrmStat.ProgressBar1.Value + 1

    
    DoEvents
Next Object

FrmStat.Visible = False

Exit Sub

errhandler:
    MsgBox "error cargando objetos"


End Sub

Sub LoadOBJData_Nuevo()

'Call LogTarea("Sub LoadOBJData")

On Error GoTo errhandler
'On Error GoTo 0

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer

Dim A As Long, s As Long

A = INICarga(DatPath & "Obj.dat")
Call INIConf(A, 0, "", 0)

'obtiene el numero de obj
'NumObjDatas = val(GetVar(DatPath & "Obj.dat", "INIT", "NumObjs"))
s = INIBuscarSeccion(A, "INIT")
NumObjDatas = INIDarClaveInt(A, s, "NumOBJs")

frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumObjDatas
frmCargando.cargar.Value = 0


ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
    s = INIBuscarSeccion(A, "OBJ" & Object)
    
    'ObjData(Object).Name = GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Name")
    ObjData(Object).Name = INIDarClaveStr(A, s, "Name")
    
    'ObjData(Object).GrhIndex = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "GrhIndex"))
    ObjData(Object).GrhIndex = INIDarClaveInt(A, s, "GrhIndex")
    
    'ObjData(Object).ObjType = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "ObjType"))
    'ObjData(Object).SubTipo = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Subtipo"))
    
    ObjData(Object).ObjType = INIDarClaveInt(A, s, "ObjType")
    ObjData(Object).SubTipo = INIDarClaveInt(A, s, "Subtipo")
    
    'ObjData(Object).Newbie = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Newbie"))
    ObjData(Object).Newbie = INIDarClaveInt(A, s, "Newbie")
    
    If ObjData(Object).SubTipo = OBJTYPE_ESCUDO Then
'        ObjData(Object).ShieldAnim = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Anim"))
'        ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
'        ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
'        ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
'        ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
        ObjData(Object).ShieldAnim = INIDarClaveInt(A, s, "Anim")
        ObjData(Object).LingH = INIDarClaveInt(A, s, "LingH")
        ObjData(Object).LingP = INIDarClaveInt(A, s, "LingP")
        ObjData(Object).LingO = INIDarClaveInt(A, s, "LingO")
        ObjData(Object).SkHerreria = INIDarClaveInt(A, s, "SkHerreria")
    End If
    
    If ObjData(Object).SubTipo = OBJTYPE_CASCO Then
'        ObjData(Object).CascoAnim = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Anim"))
'        ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
'        ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
'        ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
'        ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
        ObjData(Object).CascoAnim = INIDarClaveInt(A, s, "Anim")
        ObjData(Object).LingH = INIDarClaveInt(A, s, "LingH")
        ObjData(Object).LingP = INIDarClaveInt(A, s, "LingP")
        ObjData(Object).LingO = INIDarClaveInt(A, s, "LingO")
        ObjData(Object).SkHerreria = INIDarClaveInt(A, s, "SkHerreria")
    End If
    
'    ObjData(Object).Ropaje = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "NumRopaje"))
'    ObjData(Object).HechizoIndex = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "HechizoIndex"))
    ObjData(Object).Ropaje = INIDarClaveInt(A, s, "NumRopaje")
    ObjData(Object).HechizoIndex = INIDarClaveInt(A, s, "HechizoIndex")
    
    If ObjData(Object).ObjType = OBJTYPE_WEAPON Then
'            ObjData(Object).WeaponAnim = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Anim"))
'            ObjData(Object).Apuñala = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Apuñala"))
'            ObjData(Object).Envenena = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Envenena"))
'            ObjData(Object).MaxHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxHIT"))
'            ObjData(Object).MinHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHIT"))
'            ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
'            ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
'            ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
'            ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
'            ObjData(Object).Real = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Real"))
'            ObjData(Object).Caos = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Caos"))
'            ObjData(Object).proyectil = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Proyectil"))
'            ObjData(Object).Municion = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Municiones"))
            
            ObjData(Object).WeaponAnim = INIDarClaveInt(A, s, "Anim")
            ObjData(Object).Apuñala = INIDarClaveInt(A, s, "Apuñala")
            ObjData(Object).Envenena = INIDarClaveInt(A, s, "Envenena")
            ObjData(Object).MaxHIT = INIDarClaveInt(A, s, "MaxHIT")
            ObjData(Object).MinHIT = INIDarClaveInt(A, s, "MinHIT")
            ObjData(Object).LingH = INIDarClaveInt(A, s, "LingH")
            ObjData(Object).LingP = INIDarClaveInt(A, s, "LingP")
            ObjData(Object).LingO = INIDarClaveInt(A, s, "LingO")
            ObjData(Object).SkHerreria = INIDarClaveInt(A, s, "SkHerreria")
            ObjData(Object).Real = INIDarClaveInt(A, s, "Real")
            ObjData(Object).Caos = INIDarClaveInt(A, s, "Caos")
            ObjData(Object).proyectil = INIDarClaveInt(A, s, "Proyectil")
            ObjData(Object).Municion = INIDarClaveInt(A, s, "Municiones")
            '[KEVIN]
            ObjData(Object).DosManos = INIDarClaveInt(A, s, "DosManos")
            ObjData(Object).DosArmas = INIDarClaveInt(A, s, "DosArmas")
            '[/KEVIN]
    End If
    
    If ObjData(Object).ObjType = OBJTYPE_ARMOUR Then
'            ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
'            ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
'            ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
'            ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
'            ObjData(Object).Real = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Real"))
'            ObjData(Object).Caos = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Caos"))
            ObjData(Object).LingH = INIDarClaveInt(A, s, "LingH")
            ObjData(Object).LingP = INIDarClaveInt(A, s, "LingP")
            ObjData(Object).LingO = INIDarClaveInt(A, s, "LingO")
            ObjData(Object).SkHerreria = INIDarClaveInt(A, s, "SkHerreria")
            ObjData(Object).Real = INIDarClaveInt(A, s, "Real")
            ObjData(Object).Caos = INIDarClaveInt(A, s, "Caos")
    End If
    
    If ObjData(Object).ObjType = OBJTYPE_HERRAMIENTAS Then
'            ObjData(Object).LingH = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingH"))
'            ObjData(Object).LingP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingP"))
'            ObjData(Object).LingO = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingO"))
'            ObjData(Object).SkHerreria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkHerreria"))
            ObjData(Object).LingH = INIDarClaveInt(A, s, "LingH")
            ObjData(Object).LingP = INIDarClaveInt(A, s, "LingP")
            ObjData(Object).LingO = INIDarClaveInt(A, s, "LingO")
            ObjData(Object).SkHerreria = INIDarClaveInt(A, s, "SkHerreria")
    End If
    
    If ObjData(Object).ObjType = OBJTYPE_INSTRUMENTOS Then
'        ObjData(Object).Snd1 = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SND1"))
'        ObjData(Object).Snd2 = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SND1"))
'        ObjData(Object).Snd3 = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SND3"))
'        ObjData(Object).MinInt = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinInt"))
        ObjData(Object).Snd1 = INIDarClaveInt(A, s, "SND1")
        ObjData(Object).Snd2 = INIDarClaveInt(A, s, "SND2")
        ObjData(Object).Snd3 = INIDarClaveInt(A, s, "SND3")
        ObjData(Object).MinInt = INIDarClaveInt(A, s, "MinInt")
    End If
    
    'ObjData(Object).LingoteIndex = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "LingoteIndex"))
    ObjData(Object).LingoteIndex = INIDarClaveInt(A, s, "LingoteIndex")
    
    If ObjData(Object).ObjType = 31 Or ObjData(Object).ObjType = 23 Then
        'ObjData(Object).MinSkill = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinSkill"))
        ObjData(Object).MinSkill = INIDarClaveInt(A, s, "MinSkill")
    End If
    
'    ObjData(Object).MineralIndex = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MineralIndex"))
'
'    ObjData(Object).MaxHP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxHP"))
'    ObjData(Object).MinHP = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHP"))
'
'
'    ObjData(Object).Mujer = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Mujer"))
'    ObjData(Object).Hombre = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Hombre"))
'
'    ObjData(Object).MinHam = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHam"))
'    ObjData(Object).MinSed = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinAgu"))
    
    ObjData(Object).MineralIndex = INIDarClaveInt(A, s, "MineralIndex")
    
    ObjData(Object).MaxHP = INIDarClaveInt(A, s, "MaxHP")
    ObjData(Object).MinHP = INIDarClaveInt(A, s, "MinHP")
  
    ObjData(Object).Mujer = INIDarClaveInt(A, s, "Mujer")
    ObjData(Object).Hombre = INIDarClaveInt(A, s, "Hombre")
    
    ObjData(Object).MinHam = INIDarClaveInt(A, s, "MinHam")
    ObjData(Object).MinSed = INIDarClaveInt(A, s, "MinAgu")

    
'    ObjData(Object).MinDef = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MINDEF"))
'    ObjData(Object).MaxDef = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MAXDEF"))
'
'    ObjData(Object).Respawn = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "ReSpawn"))
'
'    ObjData(Object).RazaEnana = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "RazaEnana"))
'
'    ObjData(Object).Valor = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Valor"))
'
'    ObjData(Object).Crucial = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Crucial"))
'
'    ObjData(Object).Cerrada = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "abierta"))

    ObjData(Object).MinDef = INIDarClaveInt(A, s, "MINDEF")
    ObjData(Object).MaxDef = INIDarClaveInt(A, s, "MAXDEF")
    
    '[KEVIN]
    ObjData(Object).Vendible = INIDarClaveInt(A, s, "Vendible")
    ObjData(Object).DefMagia = INIDarClaveInt(A, s, "DefMagia")
    ObjData(Object).MagiaMinDef = INIDarClaveInt(A, s, "MAGIAMINDEF")
    ObjData(Object).MagiaMaxDef = INIDarClaveInt(A, s, "MAGIAMAXDEF")
    '[/KEVIN]
    
    ObjData(Object).Respawn = INIDarClaveInt(A, s, "ReSpawn")
    
    ObjData(Object).RazaEnana = INIDarClaveInt(A, s, "RazaEnana")
    
    '[KEVIN]
    ObjData(Object).RazaOrca = INIDarClaveInt(A, s, "RazaOrca")
    '[/KEVIN]
    
    ObjData(Object).Valor = INIDarClaveInt(A, s, "Valor")
    
    '[KEVIN]
    ObjData(Object).Materiales = INIDarClaveInt(A, s, "Materiales")
    ObjData(Object).Materiales2 = INIDarClaveInt(A, s, "Materiales2")
    '[/KEVIN]
    
    ObjData(Object).Crucial = INIDarClaveInt(A, s, "Crucial")
    
    ObjData(Object).Cerrada = INIDarClaveInt(A, s, "abierta")

    If ObjData(Object).Cerrada = 1 Then
'            ObjData(Object).Llave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Llave"))
'            ObjData(Object).Clave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Clave"))
            ObjData(Object).Llave = INIDarClaveInt(A, s, "Llave")
            ObjData(Object).Clave = INIDarClaveInt(A, s, "Clave")
    End If
    
    
    If ObjData(Object).ObjType = OBJTYPE_PUERTAS Or ObjData(Object).ObjType = OBJTYPE_BOTELLAVACIA Or ObjData(Object).ObjType = OBJTYPE_BOTELLALLENA Then
'        ObjData(Object).IndexAbierta = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "IndexAbierta"))
'        ObjData(Object).IndexCerrada = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "IndexCerrada"))
'        ObjData(Object).IndexCerradaLlave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "IndexCerradaLlave"))
        ObjData(Object).IndexAbierta = INIDarClaveInt(A, s, "IndexAbierta")
        ObjData(Object).IndexCerrada = INIDarClaveInt(A, s, "IndexCerrada")
        ObjData(Object).IndexCerradaLlave = INIDarClaveInt(A, s, "IndexCerradaLlave")
    End If
    
    '[KEVIN]
    If ObjData(Object).ObjType = OBJTYPE_CONTENEDORES Then
        ObjData(Object).IndexAbierta = INIDarClaveInt(A, s, "IndexAbierta")
        ObjData(Object).IndexCerradaLlave = INIDarClaveInt(A, s, "IndexCerradaLlave")
        ObjData(Object).Llave = INIDarClaveInt(A, s, "Llave")
        ObjData(Object).Clave = INIDarClaveInt(A, s, "Clave")
        ObjData(Object).CofreNro = INIDarClaveInt(A, s, "NroCofre")
    End If
    '[/KEVIN]
    
    'Puertas y llaves
'    ObjData(Object).Clave = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Clave"))
'
'    ObjData(Object).texto = GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Texto")
'    ObjData(Object).GrhSecundario = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "VGrande"))
'
'    ObjData(Object).Agarrable = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Agarrable"))
'    ObjData(Object).ForoID = GetVar(DatPath & "Obj.dat", "OBJ" & Object, "ID")
    ObjData(Object).Clave = INIDarClaveInt(A, s, "Clave")
    
    ObjData(Object).texto = INIDarClaveStr(A, s, "Texto")
    ObjData(Object).GrhSecundario = INIDarClaveInt(A, s, "VGrande")
    
    ObjData(Object).Agarrable = INIDarClaveInt(A, s, "Agarrable")
    ObjData(Object).ForoID = INIDarClaveStr(A, s, "ID")

    
    Dim i As Integer
    For i = 1 To NUMCLASES
        'ObjData(Object).ClaseProhibida(i) = GetVar(DatPath & "Obj.dat", "OBJ" & Object, "CP" & i)
        ObjData(Object).ClaseProhibida(i) = INIDarClaveStr(A, s, "CP" & i)
    Next
            
    'ObjData(Object).Resistencia = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Resistencia"))
    ObjData(Object).Resistencia = INIDarClaveInt(A, s, "Resistencia")
    
    'Pociones
    If ObjData(Object).ObjType = 11 Then
'        ObjData(Object).TipoPocion = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "TipoPocion"))
'        ObjData(Object).MaxModificador = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxModificador"))
'        ObjData(Object).MinModificador = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinModificador"))
'        ObjData(Object).DuracionEfecto = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "DuracionEfecto"))
        ObjData(Object).TipoPocion = INIDarClaveInt(A, s, "TipoPocion")
        ObjData(Object).MaxModificador = INIDarClaveInt(A, s, "MaxModificador")
        ObjData(Object).MinModificador = INIDarClaveInt(A, s, "MinModificador")
        ObjData(Object).DuracionEfecto = INIDarClaveInt(A, s, "DuracionEfecto")
    
    End If
    
    '[KEVIN]
    If ObjData(Object).ObjType = OBJTYPE_ANILLO Then
        ObjData(Object).TipoAnillo = INIDarClaveInt(A, s, "TipoAnillo")
        ObjData(Object).MaxModificador = INIDarClaveInt(A, s, "MaxModificador")
        ObjData(Object).LingH = INIDarClaveInt(A, s, "LingH")
        ObjData(Object).LingP = INIDarClaveInt(A, s, "LingP")
        ObjData(Object).LingO = INIDarClaveInt(A, s, "LingO")
        ObjData(Object).SkHerreria = INIDarClaveInt(A, s, "SkHerreria")
    End If
    '[/KEVIN]

'    ObjData(Object).SkCarpinteria = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "SkCarpinteria"))
    ObjData(Object).SkCarpinteria = INIDarClaveInt(A, s, "SkCarpinteria")
    
    '[KEVIN]
    ObjData(Object).SkPociones = INIDarClaveInt(A, s, "SkPociones")
    ObjData(Object).SkSastreria = INIDarClaveInt(A, s, "SkSastreria")
    '[/KEVIN]
    
    If ObjData(Object).SkCarpinteria > 0 Then
        'ObjData(Object).Madera = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "Madera"))
        ObjData(Object).Madera = INIDarClaveInt(A, s, "Madera")
    End If
    
    '[KEVIN]
    If ObjData(Object).SkPociones > 0 Then _
        ObjData(Object).Raices = INIDarClaveInt(A, s, "Raices")
    '[/KEVIN]
    
    '[KEVIN]
    If ObjData(Object).SkSastreria > 0 Then
        ObjData(Object).PielLobo = INIDarClaveInt(A, s, "PielLobo")
        ObjData(Object).PielOsoPardo = INIDarClaveInt(A, s, "PielOsoPardo")
        ObjData(Object).PielOsoPolar = INIDarClaveInt(A, s, "PielOsoPolar")
    End If
    '[/KEVIN]
    
    If ObjData(Object).ObjType = OBJTYPE_BARCOS Then
'            ObjData(Object).MaxHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxHIT"))
'            ObjData(Object).MinHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHIT"))
            ObjData(Object).MaxHIT = INIDarClaveInt(A, s, "MaxHIT")
            ObjData(Object).MinHIT = INIDarClaveInt(A, s, "MinHIT")
    End If
    
    If ObjData(Object).ObjType = OBJTYPE_FLECHAS Then
'            ObjData(Object).MaxHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MaxHIT"))
'            ObjData(Object).MinHIT = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinHIT"))
            ObjData(Object).MaxHIT = INIDarClaveInt(A, s, "MaxHIT")
            ObjData(Object).MinHIT = INIDarClaveInt(A, s, "MinHIT")
    End If
    
    'Bebidas
    'ObjData(Object).MinSta = val(GetVar(DatPath & "Obj.dat", "OBJ" & Object, "MinST"))
    ObjData(Object).MinSta = INIDarClaveInt(A, s, "MinST")
    
    frmCargando.cargar.Value = frmCargando.cargar.Value + 1

    
    'DoEvents
Next Object


Call INIDescarga(A)

Exit Sub

errhandler:

Call INIDescarga(A)

    MsgBox "error cargando objetos: " & Err.Number & " : " & Err.Description


End Sub


Sub LoadUserStats(UserIndex As Integer, UserFile As String)



Dim LoopC As Integer

For LoopC = 1 To NUMATRIBUTOS
  UserList(UserIndex).Stats.UserAtributos(LoopC) = GetVar(UserFile, "ATRIBUTOS", "AT" & LoopC)
  UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) = UserList(UserIndex).Stats.UserAtributos(LoopC)
Next

For LoopC = 1 To NUMSKILLS
  UserList(UserIndex).Stats.UserSkills(LoopC) = val(GetVar(UserFile, "SKILLS", "SK" & LoopC))
Next

For LoopC = 1 To MAXUSERHECHIZOS
  UserList(UserIndex).Stats.UserHechizos(LoopC) = val(GetVar(UserFile, "Hechizos", "H" & LoopC))
Next

UserList(UserIndex).Stats.GLD = val(GetVar(UserFile, "STATS", "GLD"))
UserList(UserIndex).Stats.Banco = val(GetVar(UserFile, "STATS", "BANCO"))

UserList(UserIndex).Stats.MET = val(GetVar(UserFile, "STATS", "MET"))
UserList(UserIndex).Stats.MaxHP = val(GetVar(UserFile, "STATS", "MaxHP"))
UserList(UserIndex).Stats.MinHP = val(GetVar(UserFile, "STATS", "MinHP"))

UserList(UserIndex).Stats.FIT = val(GetVar(UserFile, "STATS", "FIT"))
UserList(UserIndex).Stats.MinSta = val(GetVar(UserFile, "STATS", "MinSTA"))
UserList(UserIndex).Stats.MaxSta = val(GetVar(UserFile, "STATS", "MaxSTA"))

UserList(UserIndex).Stats.MaxMAN = val(GetVar(UserFile, "STATS", "MaxMAN"))
UserList(UserIndex).Stats.MinMAN = val(GetVar(UserFile, "STATS", "MinMAN"))

UserList(UserIndex).Stats.MaxHIT = val(GetVar(UserFile, "STATS", "MaxHIT"))
UserList(UserIndex).Stats.MinHIT = val(GetVar(UserFile, "STATS", "MinHIT"))

'[KEVIN]
UserList(UserIndex).Stats.MaxHitBK = UserList(UserIndex).Stats.MaxHIT
UserList(UserIndex).Stats.MinHitBK = UserList(UserIndex).Stats.MinHIT
UserList(UserIndex).Stats.Matrimonio = GetVar(UserFile, "STATS", "Matrimonio")
'[/KEVIN]

UserList(UserIndex).Stats.MaxAGU = val(GetVar(UserFile, "STATS", "MaxAGU"))
UserList(UserIndex).Stats.MinAGU = val(GetVar(UserFile, "STATS", "MinAGU"))

UserList(UserIndex).Stats.MaxHam = val(GetVar(UserFile, "STATS", "MaxHAM"))
UserList(UserIndex).Stats.MinHam = val(GetVar(UserFile, "STATS", "MinHAM"))

UserList(UserIndex).Stats.SkillPts = val(GetVar(UserFile, "STATS", "SkillPtsLibres"))

UserList(UserIndex).Stats.Exp = val(GetVar(UserFile, "STATS", "EXP"))
UserList(UserIndex).Stats.ELU = val(GetVar(UserFile, "STATS", "ELU"))
UserList(UserIndex).Stats.ELV = val(GetVar(UserFile, "STATS", "ELV"))


UserList(UserIndex).Stats.UsuariosMatados = val(GetVar(UserFile, "MUERTES", "UserMuertes"))
UserList(UserIndex).Stats.CriminalesMatados = val(GetVar(UserFile, "MUERTES", "CrimMuertes"))
UserList(UserIndex).Stats.NPCsMuertos = val(GetVar(UserFile, "MUERTES", "NpcsMuertes"))

End Sub

Sub LoadUserReputacion(UserIndex As Integer, UserFile As String)

UserList(UserIndex).Reputacion.AsesinoRep = val(GetVar(UserFile, "REP", "Asesino"))
UserList(UserIndex).Reputacion.BandidoRep = val(GetVar(UserFile, "REP", "Dandido"))
UserList(UserIndex).Reputacion.BurguesRep = val(GetVar(UserFile, "REP", "Burguesia"))
UserList(UserIndex).Reputacion.LadronesRep = val(GetVar(UserFile, "REP", "Ladrones"))
UserList(UserIndex).Reputacion.NobleRep = val(GetVar(UserFile, "REP", "Nobles"))
UserList(UserIndex).Reputacion.PlebeRep = val(GetVar(UserFile, "REP", "Plebe"))
UserList(UserIndex).Reputacion.Promedio = val(GetVar(UserFile, "REP", "Promedio"))

End Sub


Sub LoadUserInit(UserIndex As Integer, UserFile As String)


Dim LoopC As Integer
Dim LoopD As Integer
Dim ln As String
Dim ln2 As String

UserList(UserIndex).Faccion.ArmadaReal = val(GetVar(UserFile, "FACCIONES", "EjercitoReal"))
UserList(UserIndex).Faccion.FuerzasCaos = val(GetVar(UserFile, "FACCIONES", "EjercitoCaos"))
UserList(UserIndex).Faccion.CiudadanosMatados = val(GetVar(UserFile, "FACCIONES", "CiudMatados"))
UserList(UserIndex).Faccion.CriminalesMatados = val(GetVar(UserFile, "FACCIONES", "CrimMatados"))
UserList(UserIndex).Faccion.RecibioArmaduraCaos = val(GetVar(UserFile, "FACCIONES", "rArCaos"))
UserList(UserIndex).Faccion.RecibioArmaduraReal = val(GetVar(UserFile, "FACCIONES", "rArReal"))
UserList(UserIndex).Faccion.RecibioExpInicialCaos = val(GetVar(UserFile, "FACCIONES", "rExCaos"))
UserList(UserIndex).Faccion.RecibioExpInicialReal = val(GetVar(UserFile, "FACCIONES", "rExReal"))
UserList(UserIndex).Faccion.RecompensasCaos = val(GetVar(UserFile, "FACCIONES", "recCaos"))
UserList(UserIndex).Faccion.RecompensasReal = val(GetVar(UserFile, "FACCIONES", "recReal"))

'[KEVIN]
UserList(UserIndex).Quest.Quest = val(GetVar(UserFile, "QUEST", "NroQuest"))
UserList(UserIndex).Quest.EnQuest = val(GetVar(UserFile, "QUEST", "EnQuest"))
UserList(UserIndex).Quest.RealizoQuest = val(GetVar(UserFile, "QUEST", "RealizoQuest"))
UserList(UserIndex).Quest.Recompensa = val(GetVar(UserFile, "QUEST", "Recompensa"))
'[/KEVIN]

UserList(UserIndex).Flags.Muerto = val(GetVar(UserFile, "FLAGS", "Muerto"))
UserList(UserIndex).Flags.Escondido = val(GetVar(UserFile, "FLAGS", "Escondido"))

UserList(UserIndex).Flags.Hambre = val(GetVar(UserFile, "FLAGS", "Hambre"))
UserList(UserIndex).Flags.Sed = val(GetVar(UserFile, "FLAGS", "Sed"))
UserList(UserIndex).Flags.Desnudo = val(GetVar(UserFile, "FLAGS", "Desnudo"))

UserList(UserIndex).Flags.Envenenado = val(GetVar(UserFile, "FLAGS", "Envenenado"))
UserList(UserIndex).Flags.Paralizado = val(GetVar(UserFile, "FLAGS", "Paralizado"))
UserList(UserIndex).Flags.Navegando = val(GetVar(UserFile, "FLAGS", "Navegando"))

'[KEVIN]
UserList(UserIndex).Flags.EnTorneo = 0
'[/KEVIN]

UserList(UserIndex).Counters.Pena = val(GetVar(UserFile, "COUNTERS", "Pena"))

UserList(UserIndex).Email = GetVar(UserFile, "CONTACTO", "Email")

UserList(UserIndex).Genero = GetVar(UserFile, "INIT", "Genero")
UserList(UserIndex).Clase = GetVar(UserFile, "INIT", "Clase")
UserList(UserIndex).Raza = GetVar(UserFile, "INIT", "Raza")
UserList(UserIndex).Hogar = GetVar(UserFile, "INIT", "Hogar")
UserList(UserIndex).Char.Heading = val(GetVar(UserFile, "INIT", "Heading"))


UserList(UserIndex).OrigChar.Head = val(GetVar(UserFile, "INIT", "Head"))
UserList(UserIndex).OrigChar.Body = val(GetVar(UserFile, "INIT", "Body"))
UserList(UserIndex).OrigChar.WeaponAnim = val(GetVar(UserFile, "INIT", "Arma"))
UserList(UserIndex).OrigChar.ShieldAnim = val(GetVar(UserFile, "INIT", "Escudo"))
UserList(UserIndex).OrigChar.CascoAnim = val(GetVar(UserFile, "INIT", "Casco"))
UserList(UserIndex).OrigChar.Heading = SOUTH

If UserList(UserIndex).Flags.Muerto = 0 Then
        UserList(UserIndex).Char = UserList(UserIndex).OrigChar
Else

'[KEVIN]
    If Criminal(UserIndex) Then
        UserList(UserIndex).Char.Body = iCuerpoMuertoCrimi
        UserList(UserIndex).Char.Head = iCabezaMuertoCrimi
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    Else
        UserList(UserIndex).Char.Body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    End If
End If
'[/KEVIN]

UserList(UserIndex).Desc = GetVar(UserFile, "INIT", "Desc")


UserList(UserIndex).Pos.Map = val(ReadField(1, GetVar(UserFile, "INIT", "Position"), 45))
UserList(UserIndex).Pos.X = val(ReadField(2, GetVar(UserFile, "INIT", "Position"), 45))
UserList(UserIndex).Pos.Y = val(ReadField(3, GetVar(UserFile, "INIT", "Position"), 45))

UserList(UserIndex).Invent.NroItems = GetVar(UserFile, "Inventory", "CantidadItems")

'[KEVIN]--------------------------------------------------------------------
'***********************************************************************************
UserList(UserIndex).BancoInvent.NroItems = GetVar(UserFile, "BancoInventory", "CantidadItems")

'Lista de objetos del banco
For LoopD = 1 To MAX_BANCOINVENTORY_SLOTS
    ln2 = GetVar(UserFile, "BancoInventory", "Obj" & LoopD)
    UserList(UserIndex).BancoInvent.Object(LoopD).ObjIndex = val(ReadField(1, ln2, 45))
    UserList(UserIndex).BancoInvent.Object(LoopD).Amount = val(ReadField(2, ln2, 45))
Next LoopD
'------------------------------------------------------------------------------------
'[/KEVIN]*****************************************************************************

'Lista de objetos
For LoopC = 1 To MAX_INVENTORY_SLOTS
    ln = GetVar(UserFile, "Inventory", "Obj" & LoopC)
    UserList(UserIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
    UserList(UserIndex).Invent.Object(LoopC).Equipped = val(ReadField(3, ln, 45))
Next LoopC

'Obtiene el indice-objeto del arma
UserList(UserIndex).Invent.WeaponEqpSlot = val(GetVar(UserFile, "Inventory", "WeaponEqpSlot"))
If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
    UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del armadura
UserList(UserIndex).Invent.ArmourEqpSlot = val(GetVar(UserFile, "Inventory", "ArmourEqpSlot"))
If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.ArmourEqpSlot).ObjIndex
    UserList(UserIndex).Flags.Desnudo = 0
Else
    UserList(UserIndex).Flags.Desnudo = 1
End If

'Obtiene el indice-objeto del escudo
UserList(UserIndex).Invent.EscudoEqpSlot = val(GetVar(UserFile, "Inventory", "EscudoEqpSlot"))
If UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.EscudoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del casco
UserList(UserIndex).Invent.CascoEqpSlot = val(GetVar(UserFile, "Inventory", "CascoEqpSlot"))
If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.CascoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto barco
UserList(UserIndex).Invent.BarcoSlot = val(GetVar(UserFile, "Inventory", "BarcoSlot"))
If UserList(UserIndex).Invent.BarcoSlot > 0 Then
    UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.BarcoSlot).ObjIndex
End If

'Obtiene el indice-objeto municion
UserList(UserIndex).Invent.MunicionEqpSlot = val(GetVar(UserFile, "Inventory", "MunicionSlot"))
If UserList(UserIndex).Invent.MunicionEqpSlot > 0 Then
    UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).ObjIndex
End If

'[KEVIN] *-*-**-*-*-*-*-****************************************
'Obtiene el indice-objeto Herramienta
UserList(UserIndex).Invent.HerramientaEqpSlot = val(GetVar(UserFile, "Inventory", "HerramientaSlot"))
If UserList(UserIndex).Invent.MunicionEqpSlot > 0 Then
    UserList(UserIndex).Invent.HerramientaEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.HerramientaEqpSlot).ObjIndex
End If

'TUVE QUE SACARLO PORKE SALTABA ERROR Y NO SE PODIA CONECTAR
'Obtiene el indice-objeto anillo
'UserList(UserIndex).Invent.AnilloEqpSlot = val(GetVar(UserFile, "Inventory", "AnilloSlot"))
'If UserList(UserIndex).Invent.AnilloEqpSlot > 0 Then
    'UserList(UserIndex).Invent.AnilloEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.AnilloEqpSlot).ObjIndex
'End If

'Obtiene el indice-objeto anillo 2
'UserList(UserIndex).Invent.Anillo2EqpSlot = val(GetVar(UserFile, "Inventory", "Anillo2Slot"))
'If UserList(UserIndex).Invent.Anillo2EqpSlot > 0 Then
    'UserList(UserIndex).Invent.Anillo2EqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.Anillo2EqpSlot).ObjIndex
'End If
'[/KEVIN]****************---------------------------**********************


UserList(UserIndex).NroMacotas = val(GetVar(UserFile, "Mascotas", "NroMascotas"))

'Lista de objetos
For LoopC = 1 To MAXMASCOTAS
    UserList(UserIndex).MascotasType(LoopC) = val(GetVar(UserFile, "Mascotas", "Mas" & LoopC))
Next LoopC

UserList(UserIndex).GuildInfo.FundoClan = val(GetVar(UserFile, "Guild", "FundoClan"))
UserList(UserIndex).GuildInfo.EsGuildLeader = val(GetVar(UserFile, "Guild", "EsGuildLeader"))
UserList(UserIndex).GuildInfo.Echadas = val(GetVar(UserFile, "Guild", "Echadas"))
UserList(UserIndex).GuildInfo.Solicitudes = val(GetVar(UserFile, "Guild", "Solicitudes"))
UserList(UserIndex).GuildInfo.SolicitudesRechazadas = val(GetVar(UserFile, "Guild", "SolicitudesRechazadas"))
UserList(UserIndex).GuildInfo.VecesFueGuildLeader = val(GetVar(UserFile, "Guild", "VecesFueGuildLeader"))
UserList(UserIndex).GuildInfo.YaVoto = val(GetVar(UserFile, "Guild", "YaVoto"))
UserList(UserIndex).GuildInfo.ClanesParticipo = val(GetVar(UserFile, "Guild", "ClanesParticipo"))
UserList(UserIndex).GuildInfo.GuildPoints = val(GetVar(UserFile, "Guild", "GuildPts"))

UserList(UserIndex).GuildInfo.ClanFundado = GetVar(UserFile, "Guild", "ClanFundado")
UserList(UserIndex).GuildInfo.GuildName = GetVar(UserFile, "Guild", "GuildName")

End Sub





Function GetVar(file As String, Main As String, Var As String) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space(5000) ' This tells the computer how long the longest string can be
  
  
getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), file
  
GetVar = RTrim(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()

'Call LogTarea("Sub CargarBackUp")

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."

Dim Map As Integer
Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim TempInt As Integer
Dim SaveAs As String
Dim npcfile As String
Dim Porc As Long
Dim FileNamE As String
Dim c$
    
On Error GoTo man

 
NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumMaps
frmCargando.cargar.Value = 0

MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")

ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
ReDim MapInfo(1 To NumMaps) As MapInfo
  
For Map = 1 To NumMaps
    
    FileNamE = App.Path & "\WorldBackUp\Map" & Map & ".map"
    
    If FileExist(FileNamE, vbNormal) Then
        Open App.Path & "\WorldBackUp\Map" & Map & ".map" For Binary As #1
        Open App.Path & "\WorldBackUp\Map" & Map & ".inf" For Binary As #2
        c$ = App.Path & "\WorldBackUp\Map" & Map & ".dat"
    Else
        Open App.Path & MapPath & "Mapa" & Map & ".map" For Binary As #1
        Open App.Path & MapPath & "Mapa" & Map & ".inf" For Binary As #2
        c$ = App.Path & MapPath & "Mapa" & Map & ".dat"
    End If
    
        Seek #1, 1
        Seek #2, 1
        'map Header
        Get #1, , MapInfo(Map).MapVersion
        Get #1, , MiCabecera
        Get #1, , TempInt
        Get #1, , TempInt
        Get #1, , TempInt
        Get #1, , TempInt
        'inf Header
        Get #2, , TempInt
        Get #2, , TempInt
        Get #2, , TempInt
        Get #2, , TempInt
        Get #2, , TempInt
        'Load arrays
        For Y = YMinMapSize To YMaxMapSize
            For X = XMinMapSize To XMaxMapSize
        
                    '.dat file
                    Get #1, , MapData(Map, X, Y).Blocked
                    
                    'Get GRH number
                    For LoopC = 1 To 4
                        Get #1, , MapData(Map, X, Y).Graphic(LoopC)
                    Next LoopC
                    
                    'Space holder for future expansion
                    Get #1, , MapData(Map, X, Y).trigger
                    Get #1, , TempInt
                    
                                        
                    '.inf file
                    Get #2, , MapData(Map, X, Y).TileExit.Map
                    Get #2, , MapData(Map, X, Y).TileExit.X
                    Get #2, , MapData(Map, X, Y).TileExit.Y
                    
                    'Get and make NPC
                    Get #2, , MapData(Map, X, Y).NpcIndex
                    If MapData(Map, X, Y).NpcIndex > 0 Then
                        MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                        'Si el npc debe hacer respawn en la pos
                        'original la guardamos
                        
                        If Npclist(MapData(Map, X, Y).NpcIndex).Numero > 499 Then
                            npcfile = DatPath & "NPCs-HOSTILES.dat"
                        Else
                            npcfile = DatPath & "NPCs.dat"
                        End If
                        
                        Dim fl As Byte
                        fl = val(GetVar(npcfile, "NPC" & Npclist(MapData(Map, X, Y).NpcIndex).Numero, "PosOrig"))
                        If fl = 1 Then
                            Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                            Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                            Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                        Else
                            Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = 0
                            Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = 0
                            Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = 0
                        End If
        
                        Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
                        Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
                        Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
                        
                        
                        'Si existe el backup lo cargamos
                        If Npclist(MapData(Map, X, Y).NpcIndex).Flags.BackUp = 1 Then
                                'cargamos el nuevo del backup
                                Call CargarNpcBackUp(MapData(Map, X, Y).NpcIndex, Npclist(MapData(Map, X, Y).NpcIndex).Numero)
                                
                        End If
                        
                        Call MakeNPCChar(ToNone, 0, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
                    End If

                    'Get and make Object
                    Get #2, , MapData(Map, X, Y).OBJInfo.ObjIndex
                    Get #2, , MapData(Map, X, Y).OBJInfo.Amount
        
                    'Space holder for future expansion (Objects, ect.
                    Get #2, , DummyInt
                    Get #2, , DummyInt
            Next X
        Next Y
        Close #1
        Close #2
          MapInfo(Map).Name = GetVar(c$, "Mapa" & Map, "Name")
          MapInfo(Map).Music = GetVar(c$, "Mapa" & Map, "MusicNum")
          MapInfo(Map).StartPos.Map = val(ReadField(1, GetVar(c$, "Mapa" & Map, "StartPos"), 45))
          MapInfo(Map).StartPos.X = val(ReadField(2, GetVar(c$, "Mapa" & Map, "StartPos"), 45))
          MapInfo(Map).StartPos.Y = val(ReadField(3, GetVar(c$, "Mapa" & Map, "StartPos"), 45))
          If val(GetVar(c$, "Mapa" & Map, "Pk")) = 0 Then
                MapInfo(Map).Pk = True
          Else
                MapInfo(Map).Pk = False
          End If
          MapInfo(Map).Restringir = GetVar(c$, "Mapa" & Map, "Restringir")
          MapInfo(Map).BackUp = val(GetVar(c$, "Mapa" & Map, "BackUp"))
          MapInfo(Map).Terreno = GetVar(c$, "Mapa" & Map, "Terreno")
          MapInfo(Map).Zona = GetVar(c$, "Mapa" & Map, "Zona")
          
          frmCargando.cargar.Value = frmCargando.cargar.Value + 1
          
          DoEvents
Next Map

FrmStat.Visible = False

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas.")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

  

End Sub

Sub CargarBackUp_Nuevo()

'Call LogTarea("Sub CargarBackUp")

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."

Dim Map As Integer
Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim TempInt As Integer
Dim SaveAs As String
Dim npcfile As String
Dim Porc As Long
Dim FileNamE As String
Dim c$

Dim archmap As String, archinf As String

On Error GoTo man

 
NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumMaps
frmCargando.cargar.Value = 0

MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")

ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
ReDim MapInfo(1 To NumMaps) As MapInfo
  
For Map = 1 To NumMaps
    
    FileNamE = App.Path & "\WorldBackUp\Map" & Map & ".map"
    
    If FileExist(FileNamE, vbNormal) Then
        archmap = App.Path & "\WorldBackUp\Map" & Map & ".map"
        archinf = App.Path & "\WorldBackUp\Map" & Map & ".inf"
        c$ = App.Path & "\WorldBackUp\Map" & Map & ".dat"
    Else
        archmap = App.Path & MapPath & "Mapa" & Map & ".map"
        archinf = App.Path & MapPath & "Mapa" & Map & ".inf"
        c$ = App.Path & MapPath & "Mapa" & Map & ".dat"
    End If
    
        Call CargarUnMapa(Map, archmap, archinf)
        
          frmCargando.cargar.Value = frmCargando.cargar.Value + 1
          
          DoEvents
Next Map

FrmStat.Visible = False

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas.")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

  

End Sub

Sub LoadMapData()


'Call LogTarea("Sub LoadMapData")

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas."

Dim Map As Integer
Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim TempInt As Integer
Dim npcfile As String

On Error GoTo man

NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))

frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumMaps
frmCargando.cargar.Value = 0

MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")

ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
ReDim MapInfo(1 To NumMaps) As MapInfo
  
For Map = 1 To NumMaps
    DoEvents
    
    
    Open App.Path & MapPath & "Mapa" & Map & ".map" For Binary As #1
    Seek #1, 1
    
    'inf
    Open App.Path & MapPath & "Mapa" & Map & ".inf" For Binary As #2
    Seek #2, 1
    
     'map Header
    Get #1, , MapInfo(Map).MapVersion
    Get #1, , MiCabecera
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt

    'inf Header
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
    Get #2, , TempInt
        
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            '.dat file
            Get #1, , MapData(Map, X, Y).Blocked
            
            For LoopC = 1 To 4
                Get #1, , MapData(Map, X, Y).Graphic(LoopC)
            Next LoopC
            
            Get #1, , MapData(Map, X, Y).trigger
            Get #1, , TempInt
            
                                
            '.inf file
            Get #2, , MapData(Map, X, Y).TileExit.Map
            Get #2, , MapData(Map, X, Y).TileExit.X
            Get #2, , MapData(Map, X, Y).TileExit.Y
            
            'Get and make NPC
            Get #2, , MapData(Map, X, Y).NpcIndex
            If MapData(Map, X, Y).NpcIndex > 0 Then
                
                If MapData(Map, X, Y).NpcIndex > 499 Then
                        npcfile = DatPath & "NPCs-HOSTILES.dat"
                Else
                        npcfile = DatPath & "NPCs.dat"
                End If
                
                'Si el npc debe hacer respawn en la pos
                'original la guardamos
                If val(GetVar(npcfile, "NPC" & MapData(Map, X, Y).NpcIndex, "PosOrig")) = 1 Then
                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                Else
                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                End If
                
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
                
                Call MakeNPCChar(ToNone, 0, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
            End If

            'Get and make Object
            Get #2, , MapData(Map, X, Y).OBJInfo.ObjIndex
            Get #2, , MapData(Map, X, Y).OBJInfo.Amount

            'Space holder for future expansion (Objects, ect.
            Get #2, , DummyInt
            Get #2, , DummyInt
        
        Next X
    Next Y

   
    Close #1
    Close #2

  
    MapInfo(Map).Name = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Name")
    MapInfo(Map).Music = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "MusicNum")
    MapInfo(Map).StartPos.Map = val(ReadField(1, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.X = val(ReadField(2, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.Y = val(ReadField(3, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
    
    If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Pk")) = 0 Then
        MapInfo(Map).Pk = True
    Else
        MapInfo(Map).Pk = False
    End If
    
    
    MapInfo(Map).Terreno = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Terreno")

    MapInfo(Map).Zona = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Zona")
    
    MapInfo(Map).Restringir = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Restringir")
    
    MapInfo(Map).BackUp = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "BACKUP"))

frmCargando.cargar.Value = frmCargando.cargar.Value + 1

Next Map


Exit Sub

man:
    MsgBox ("Error durante la carga de mapas.")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

    
End Sub

Sub LoadMapData_Nuevo()


'Call LogTarea("Sub LoadMapData")

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas."

Dim Map As Integer
Dim LoopC As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim TempInt As Integer
Dim npcfile As String

Dim archmap As String, archinf As String

On Error GoTo man

NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))

frmCargando.cargar.Min = 0
frmCargando.cargar.max = NumMaps
frmCargando.cargar.Value = 0

MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")

ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
ReDim MapInfo(1 To NumMaps) As MapInfo
  
For Map = 1 To NumMaps
    DoEvents
    
    archmap = App.Path & MapPath & "Mapa" & Map & ".map"
    archinf = App.Path & MapPath & "Mapa" & Map & ".inf"

    Call CargarUnMapa(Map, archmap, archinf)
    
    frmCargando.cargar.Value = frmCargando.cargar.Value + 1
Next Map


Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

    
End Sub


Sub LoadSini()

'Call LogTarea("Sub LoadSini")

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."

BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

'Misc
puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
'Lee la version correcta del cliente
ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")

FastMapLoad = val(GetVar(IniPath & "Server.ini", "INIT", "FastMapLoad"))

FastObjLoad = val(GetVar(IniPath & "Server.ini", "INIT", "FastObjLoad"))

Rain = val(GetVar(IniPath & "Server.ini", "INIT", "Lluvia"))

AntiSH = val(GetVar(IniPath & "Server.ini", "INIT", "AntiSH"))

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

IntervaloUserPuedeProyectiles = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaProyectil"))
FrmInterv.txtProyectiles.Text = IntervaloUserPuedeProyectiles

frmMain.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval

frmMain.npcataca.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval

IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar

IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar

frmMain.tLluvia.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval

frmMain.CmdExec.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTimerExec"))
FrmInterv.txtCmdExec.Text = frmMain.CmdExec.Interval

FrmInterv.TextPociones.Text = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUsarItem"))

MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))
If MinutosWs < 60 Then MinutosWs = 180

'Ressurect pos
ResPos.Map = val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.X = val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.Y = val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
  
recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
  
'Max users
MaxUsers = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
Call SetKernelMaxUsers(MaxUsers)
ReDim GameInputMapArray(1 To MaxUsers) As Integer
ReDim GameOutputIndexArray(1 To MaxUsers) As Integer
ReDim UserList(1 To MaxUsers) As User
ReDim SocketHandle(1 To MaxUsers) As Long

Nix.Map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
Nix.X = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
Nix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")

'Ullathorpe.Map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
'Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
'Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")

Banderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")

'Lindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
'Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
'Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")

'[KEVIN]
NuevaEsperanza.Map = GetVar(DatPath & "Ciudades.dat", "NuevaEsperanza", "Mapa")
NuevaEsperanza.X = GetVar(DatPath & "Ciudades.dat", "NuevaEsperanza", "X")
NuevaEsperanza.Y = GetVar(DatPath & "Ciudades.dat", "NuevaEsperanza", "Y")

Dahl.Map = GetVar(DatPath & "Ciudades.dat", "Dahl", "Mapa")
Dahl.X = GetVar(DatPath & "Ciudades.dat", "Dahl", "X")
Dahl.Y = GetVar(DatPath & "Ciudades.dat", "Dahl", "Y")

Terminus.Map = GetVar(DatPath & "Ciudades.dat", "Terminus", "Mapa")
Terminus.X = GetVar(DatPath & "Ciudades.dat", "Terminus", "X")
Terminus.Y = GetVar(DatPath & "Ciudades.dat", "Terminus", "Y")

Nexon.Map = GetVar(DatPath & "Ciudades.dat", "Nexon", "Mapa")
Nexon.X = GetVar(DatPath & "Ciudades.dat", "Nexon", "X")
Nexon.Y = GetVar(DatPath & "Ciudades.dat", "Nexon", "Y")
'[/KEVIN]

End Sub

Sub WriteVar(file As String, Main As String, Var As String, Value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, Value, file
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, ByVal UserFile As String, Optional ByVal NewChar As Boolean)
On Error GoTo errhandler


If FileExist(UserFile, vbNormal) Then
       If UserList(UserIndex).Flags.Muerto = 1 Then UserList(UserIndex).Char.Head = val(GetVar(UserFile, "INIT", "Head"))
       Kill UserFile
End If

Dim LoopC As Integer

Dim LoopD As Integer


Call WriteVar(UserFile, "FLAGS", "Muerto", val(UserList(UserIndex).Flags.Muerto))
Call WriteVar(UserFile, "FLAGS", "Escondido", val(UserList(UserIndex).Flags.Escondido))
Call WriteVar(UserFile, "FLAGS", "Hambre", val(UserList(UserIndex).Flags.Hambre))
Call WriteVar(UserFile, "FLAGS", "Sed", val(UserList(UserIndex).Flags.Sed))
Call WriteVar(UserFile, "FLAGS", "Desnudo", val(UserList(UserIndex).Flags.Desnudo))
Call WriteVar(UserFile, "FLAGS", "Ban", val(UserList(UserIndex).Flags.Ban))
Call WriteVar(UserFile, "FLAGS", "Navegando", val(UserList(UserIndex).Flags.Navegando))

Call WriteVar(UserFile, "FLAGS", "Envenenado", val(UserList(UserIndex).Flags.Envenenado))
Call WriteVar(UserFile, "FLAGS", "Paralizado", val(UserList(UserIndex).Flags.Paralizado))

Call WriteVar(UserFile, "COUNTERS", "Pena", val(UserList(UserIndex).Counters.Pena))

Call WriteVar(UserFile, "FACCIONES", "EjercitoReal", val(UserList(UserIndex).Faccion.ArmadaReal))
Call WriteVar(UserFile, "FACCIONES", "EjercitoCaos", val(UserList(UserIndex).Faccion.FuerzasCaos))
Call WriteVar(UserFile, "FACCIONES", "CiudMatados", val(UserList(UserIndex).Faccion.CiudadanosMatados))
Call WriteVar(UserFile, "FACCIONES", "CrimMatados", val(UserList(UserIndex).Faccion.CriminalesMatados))
Call WriteVar(UserFile, "FACCIONES", "rArCaos", val(UserList(UserIndex).Faccion.RecibioArmaduraCaos))
Call WriteVar(UserFile, "FACCIONES", "rArReal", val(UserList(UserIndex).Faccion.RecibioArmaduraReal))
Call WriteVar(UserFile, "FACCIONES", "rExCaos", val(UserList(UserIndex).Faccion.RecibioExpInicialCaos))
Call WriteVar(UserFile, "FACCIONES", "rExReal", val(UserList(UserIndex).Faccion.RecibioExpInicialReal))
Call WriteVar(UserFile, "FACCIONES", "recCaos", val(UserList(UserIndex).Faccion.RecompensasCaos))
Call WriteVar(UserFile, "FACCIONES", "recReal", val(UserList(UserIndex).Faccion.RecompensasReal))

'[KEVIN]
If NewChar = True Then
    Call WriteVar(UserFile, "QUEST", "NroQuest", 0)
    Call WriteVar(UserFile, "QUEST", "EnQuest", 0)
    Call WriteVar(UserFile, "QUEST", "RealizoQuest", 0)
    Call WriteVar(UserFile, "QUEST", "Recompensa", 0)
Else
    Call WriteVar(UserFile, "QUEST", "NroQuest", val(UserList(UserIndex).Quest.Quest))
    Call WriteVar(UserFile, "QUEST", "EnQuest", val(UserList(UserIndex).Quest.EnQuest))
    Call WriteVar(UserFile, "QUEST", "RealizoQuest", val(UserList(UserIndex).Quest.RealizoQuest))
    Call WriteVar(UserFile, "QUEST", "Recompensa", val(UserList(UserIndex).Quest.Recompensa))
End If
'[/KEVIN]

Call WriteVar(UserFile, "GUILD", "EsGuildLeader", val(UserList(UserIndex).GuildInfo.EsGuildLeader))
Call WriteVar(UserFile, "GUILD", "Echadas", val(UserList(UserIndex).GuildInfo.Echadas))
Call WriteVar(UserFile, "GUILD", "Solicitudes", val(UserList(UserIndex).GuildInfo.Solicitudes))
Call WriteVar(UserFile, "GUILD", "SolicitudesRechazadas", val(UserList(UserIndex).GuildInfo.SolicitudesRechazadas))
Call WriteVar(UserFile, "GUILD", "VecesFueGuildLeader", val(UserList(UserIndex).GuildInfo.VecesFueGuildLeader))
Call WriteVar(UserFile, "GUILD", "YaVoto", val(UserList(UserIndex).GuildInfo.YaVoto))
Call WriteVar(UserFile, "GUILD", "FundoClan", val(UserList(UserIndex).GuildInfo.FundoClan))

Call WriteVar(UserFile, "GUILD", "GuildName", UserList(UserIndex).GuildInfo.GuildName)
Call WriteVar(UserFile, "GUILD", "ClanFundado", UserList(UserIndex).GuildInfo.ClanFundado)
Call WriteVar(UserFile, "GUILD", "ClanesParticipo", Str(UserList(UserIndex).GuildInfo.ClanesParticipo))
Call WriteVar(UserFile, "GUILD", "GuildPts", Str(UserList(UserIndex).GuildInfo.GuildPoints))

'¿Fueron modificados los atributos del usuario?
'If Not UserList(UserIndex).Flags.TomoPocion Then
    'For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        'Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributos(LoopC)))
    'Next
'Else
    'For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        'UserList(UserIndex).Stats.UserAtributos(LoopC) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)
        'Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributos(LoopC)))
    'Next
'End If

'[KEVIN] ********************************************************************
If UserList(UserIndex).Flags.TomoPocion Then
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        UserList(UserIndex).Stats.UserAtributos(LoopC) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributos(LoopC)))
    Next
ElseIf UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)

        Select Case ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).TipoAnillo
                    
            Case 1
                UserList(UserIndex).Stats.UserAtributos(2) = UserList(UserIndex).Stats.UserAtributosBackUP(2)
                Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributos(LoopC)))
                            
            Case 2
                UserList(UserIndex).Stats.UserAtributos(1) = UserList(UserIndex).Stats.UserAtributosBackUP(1)
                Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributos(LoopC)))
                            
            Case 3
                UserList(UserIndex).Stats.UserAtributos(4) = UserList(UserIndex).Stats.UserAtributosBackUP(4)
                Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributos(LoopC)))
                        
        End Select
        
    Next
ElseIf UserList(UserIndex).Invent.Anillo2EqpObjIndex > 0 Then
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)

        Select Case ObjData(UserList(UserIndex).Invent.Anillo2EqpObjIndex).TipoAnillo
                    
            Case 1
                UserList(UserIndex).Stats.UserAtributos(2) = UserList(UserIndex).Stats.UserAtributosBackUP(2)
                Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributos(LoopC)))
                            
            Case 2
                UserList(UserIndex).Stats.UserAtributos(1) = UserList(UserIndex).Stats.UserAtributosBackUP(1)
                Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributos(LoopC)))
                            
            Case 3
                UserList(UserIndex).Stats.UserAtributos(4) = UserList(UserIndex).Stats.UserAtributosBackUP(4)
                Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributos(LoopC)))
                        
        End Select
        
    Next
Else
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, val(UserList(UserIndex).Stats.UserAtributos(LoopC)))
    Next
End If
'[/KEVIN]******************************************************************

For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserSkills)
    Call WriteVar(UserFile, "SKILLS", "SK" & LoopC, val(UserList(UserIndex).Stats.UserSkills(LoopC)))
Next


Call WriteVar(UserFile, "CONTACTO", "Email", UserList(UserIndex).Email)

Call WriteVar(UserFile, "INIT", "Genero", UserList(UserIndex).Genero)
Call WriteVar(UserFile, "INIT", "Raza", UserList(UserIndex).Raza)
Call WriteVar(UserFile, "INIT", "Hogar", UserList(UserIndex).Hogar)
Call WriteVar(UserFile, "INIT", "Clase", UserList(UserIndex).Clase)
Call WriteVar(UserFile, "INIT", "Password", UserList(UserIndex).Password)
Call WriteVar(UserFile, "INIT", "Desc", UserList(UserIndex).Desc)

Call WriteVar(UserFile, "INIT", "Heading", Str(UserList(UserIndex).Char.Heading))

Call WriteVar(UserFile, "INIT", "Head", Str(UserList(UserIndex).OrigChar.Head))

If UserList(UserIndex).Flags.Muerto = 0 Then
    Call WriteVar(UserFile, "INIT", "Body", Str(UserList(UserIndex).Char.Body))
End If

Call WriteVar(UserFile, "INIT", "Arma", Str(UserList(UserIndex).Char.WeaponAnim))
Call WriteVar(UserFile, "INIT", "Escudo", Str(UserList(UserIndex).Char.ShieldAnim))
Call WriteVar(UserFile, "INIT", "Casco", Str(UserList(UserIndex).Char.CascoAnim))

Call WriteVar(UserFile, "INIT", "LastIP", UserList(UserIndex).ip)
Call WriteVar(UserFile, "INIT", "Position", UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.Y)

'[KEVIN]
Call WriteVar(UserFile, "STATS", "Matrimonio", UserList(UserIndex).Stats.Matrimonio)
'[/KEVIN]

Call WriteVar(UserFile, "STATS", "GLD", Str(UserList(UserIndex).Stats.GLD))
Call WriteVar(UserFile, "STATS", "BANCO", Str(UserList(UserIndex).Stats.Banco))

Call WriteVar(UserFile, "STATS", "MET", Str(UserList(UserIndex).Stats.MET))
Call WriteVar(UserFile, "STATS", "MaxHP", Str(UserList(UserIndex).Stats.MaxHP))
Call WriteVar(UserFile, "STATS", "MinHP", Str(UserList(UserIndex).Stats.MinHP))

Call WriteVar(UserFile, "STATS", "FIT", Str(UserList(UserIndex).Stats.FIT))
Call WriteVar(UserFile, "STATS", "MaxSTA", Str(UserList(UserIndex).Stats.MaxSta))
Call WriteVar(UserFile, "STATS", "MinSTA", Str(UserList(UserIndex).Stats.MinSta))

Call WriteVar(UserFile, "STATS", "MaxMAN", Str(UserList(UserIndex).Stats.MaxMAN))
Call WriteVar(UserFile, "STATS", "MinMAN", Str(UserList(UserIndex).Stats.MinMAN))

Call WriteVar(UserFile, "STATS", "MaxHIT", Str(UserList(UserIndex).Stats.MaxHIT))
Call WriteVar(UserFile, "STATS", "MinHIT", Str(UserList(UserIndex).Stats.MinHIT))

Call WriteVar(UserFile, "STATS", "MaxAGU", Str(UserList(UserIndex).Stats.MaxAGU))
Call WriteVar(UserFile, "STATS", "MinAGU", Str(UserList(UserIndex).Stats.MinAGU))

Call WriteVar(UserFile, "STATS", "MaxHAM", Str(UserList(UserIndex).Stats.MaxHam))
Call WriteVar(UserFile, "STATS", "MinHAM", Str(UserList(UserIndex).Stats.MinHam))

Call WriteVar(UserFile, "STATS", "SkillPtsLibres", Str(UserList(UserIndex).Stats.SkillPts))
  
Call WriteVar(UserFile, "STATS", "EXP", Str(UserList(UserIndex).Stats.Exp))
Call WriteVar(UserFile, "STATS", "ELV", Str(UserList(UserIndex).Stats.ELV))
Call WriteVar(UserFile, "STATS", "ELU", Str(UserList(UserIndex).Stats.ELU))
Call WriteVar(UserFile, "MUERTES", "UserMuertes", val(UserList(UserIndex).Stats.UsuariosMatados))
Call WriteVar(UserFile, "MUERTES", "CrimMuertes", val(UserList(UserIndex).Stats.CriminalesMatados))
Call WriteVar(UserFile, "MUERTES", "NpcsMuertes", val(UserList(UserIndex).Stats.NPCsMuertos))
  
'Save Inv
Call WriteVar(UserFile, "Inventory", "CantidadItems", val(UserList(UserIndex).Invent.NroItems))

'[KEVIN]----------------------------------------------------------------------------
'*******************************************************************************************
If NewChar = True Then
    Call WriteVar(UserFile, "BancoInventory", "CantidadItems", 0)
    
    For LoopD = 1 To MAX_BANCOINVENTORY_SLOTS
        Call WriteVar(UserFile, "BancoInventory", "Obj" & LoopD, 0 & "-" & 0)
    Next LoopD
Else
    Call WriteVar(UserFile, "BancoInventory", "CantidadItems", val(UserList(UserIndex).BancoInvent.NroItems))
    
    For LoopD = 1 To MAX_BANCOINVENTORY_SLOTS
        Call WriteVar(UserFile, "BancoInventory", "Obj" & LoopD, UserList(UserIndex).BancoInvent.Object(LoopD).ObjIndex & "-" & UserList(UserIndex).BancoInvent.Object(LoopD).Amount)
    Next LoopD
End If
'*******************************************************************************************
'[/KEVIN]---------------------------------------------------------------------------------------

For LoopC = 1 To MAX_INVENTORY_SLOTS
    Call WriteVar(UserFile, "Inventory", "Obj" & LoopC, UserList(UserIndex).Invent.Object(LoopC).ObjIndex & "-" & UserList(UserIndex).Invent.Object(LoopC).Amount & "-" & UserList(UserIndex).Invent.Object(LoopC).Equipped)
Next LoopC

Call WriteVar(UserFile, "Inventory", "WeaponEqpSlot", Str(UserList(UserIndex).Invent.WeaponEqpSlot))
Call WriteVar(UserFile, "Inventory", "ArmourEqpSlot", Str(UserList(UserIndex).Invent.ArmourEqpSlot))
Call WriteVar(UserFile, "Inventory", "CascoEqpSlot", Str(UserList(UserIndex).Invent.CascoEqpSlot))
Call WriteVar(UserFile, "Inventory", "EscudoEqpSlot", Str(UserList(UserIndex).Invent.EscudoEqpSlot))
Call WriteVar(UserFile, "Inventory", "BarcoSlot", Str(UserList(UserIndex).Invent.BarcoSlot))
Call WriteVar(UserFile, "Inventory", "MunicionSlot", Str(UserList(UserIndex).Invent.MunicionEqpSlot))
'[KEVIN]
Call WriteVar(UserFile, "Inventory", "HerramientaSlot", Str(UserList(UserIndex).Invent.HerramientaEqpSlot))
Call WriteVar(UserFile, "Inventory", "AnilloSlot", Str(UserList(UserIndex).Invent.AnilloEqpSlot))
Call WriteVar(UserFile, "Inventory", "Anillo2Slot", Str(UserList(UserIndex).Invent.Anillo2EqpSlot))
'[/KEVIN]


'Reputacion
Call WriteVar(UserFile, "REP", "Asesino", val(UserList(UserIndex).Reputacion.AsesinoRep))
Call WriteVar(UserFile, "REP", "Bandido", val(UserList(UserIndex).Reputacion.BandidoRep))
Call WriteVar(UserFile, "REP", "Burguesia", val(UserList(UserIndex).Reputacion.BurguesRep))
Call WriteVar(UserFile, "REP", "Ladrones", val(UserList(UserIndex).Reputacion.LadronesRep))
Call WriteVar(UserFile, "REP", "Nobles", val(UserList(UserIndex).Reputacion.NobleRep))
Call WriteVar(UserFile, "REP", "Plebe", val(UserList(UserIndex).Reputacion.PlebeRep))

Dim l As Long
l = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
l = l / 6
Call WriteVar(UserFile, "REP", "Promedio", val(l))

Dim cad As String

For LoopC = 1 To MAXUSERHECHIZOS
    cad = UserList(UserIndex).Stats.UserHechizos(LoopC)
    Call WriteVar(UserFile, "HECHIZOS", "H" & LoopC, cad)
Next


For LoopC = 1 To MAXMASCOTAS
    ' Mascota valida?
    If UserList(UserIndex).MascotasIndex(LoopC) > 0 Then
        ' Nos aseguramos que la criatura no fue invocada
        If Npclist(UserList(UserIndex).MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
            cad = UserList(UserIndex).MascotasType(LoopC)
        Else 'Si fue invocada no la guardamos
            cad = "0"
            UserList(UserIndex).NroMacotas = UserList(UserIndex).NroMacotas - 1
        End If
        Call WriteVar(UserFile, "MASCOTAS", "MAS" & LoopC, cad)
    End If

Next

Call WriteVar(UserFile, "MASCOTAS", "NroMascotas", Str(UserList(UserIndex).NroMacotas))

Exit Sub

errhandler:
Call LogError("Error en SaveUser")

End Sub

Function Criminal(ByVal UserIndex As Integer) As Boolean

Dim l As Long
l = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
l = l / 6
Criminal = (l < 0)

End Function




Sub BackUPnPc(NpcIndex As Integer)

'Call LogTarea("Sub BackUPnPc NpcIndex:" & NpcIndex)

Dim NpcNumero As Integer
Dim npcfile As String
Dim LoopC As Integer


NpcNumero = Npclist(NpcIndex).Numero

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


'Stats
Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.Def))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.UsuariosMatados))




'Flags
Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).Flags.Respawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(Npclist(NpcIndex).Flags.BackUp))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).Flags.Domable))

'Inventario
Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
   For LoopC = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex & "-" & Npclist(NpcIndex).Invent.Object(LoopC).Amount)
   Next
End If


End Sub



Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)

'Call LogTarea("Sub CargarNpcBackUp NpcIndex:" & NpcIndex & " NpcNumber:" & NpcNumber)

'Status
If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"


Dim npcfile As String

If NpcNumber > 499 Then
        npcfile = DatPath & "bkNPCs-HOSTILES.dat"
Else
        npcfile = DatPath & "bkNPCs.dat"
End If

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.Heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))


Npclist(NpcIndex).GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))

Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.Def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
Npclist(NpcIndex).Stats.ImpactRate = val(GetVar(npcfile, "NPC" & NpcNumber, "ImpactRate"))


Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
       
    Next LoopC
Else
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = 0
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = 0
    Next LoopC
End If

Npclist(NpcIndex).Inflacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Inflacion"))


Npclist(NpcIndex).Flags.NPCActive = True
Npclist(NpcIndex).Flags.UseAINow = False
Npclist(NpcIndex).Flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
Npclist(NpcIndex).Flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).Flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
Npclist(NpcIndex).Flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

End Sub


Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, UserList(BannedIndex).Name
Close #mifile

End Sub

Sub LogBanServer(ByVal BannedIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "BannedBy", "Server")
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserList(BannedIndex).Name, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, UserList(BannedIndex).Name
Close #mifile

End Sub

'[KEVIN]
Sub CargarJerarquias()

'1era jerarquia
ArmaduraImperial1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "ArmaduraImperial1"))
ArmaduraImperial2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "ArmaduraImperial2"))
ArmaduraImperial3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "ArmaduraImperial3"))
TunicaMagoImperial = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "TunicaMagoImperial"))
TunicaMagoImperialEnanos = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "TunicaMagoImperialEnanos"))

ArmaduraCaos1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "ArmaduraCaos1"))
ArmaduraCaos2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "ArmaduraCaos2"))
ArmaduraCaos3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "ArmaduraCaos3"))
TunicaMagoCaos = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "TunicaMagoCaos"))
TunicaMagoCaosEnanos = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "TunicaMagoCaosEnanos"))

ArmaduraImpMujer1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "ArmaduraImpMujer1"))
ArmaduraImpMujer2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "ArmaduraImpMujer2"))
ArmaduraImpMujer3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "ArmaduraImpMujer3"))
TunicaMagoImpMujer = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "TunicaMagoImpMujer"))
TunicaMagoImpEnanosMujer = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "TunicaMagoImpEnanosMujer"))

ArmaduraCaosMujer1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "ArmaduraCaosMujer1"))
ArmaduraCaosMujer2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "ArmaduraCaosMujer2"))
ArmaduraCaosMujer3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "ArmaduraCaosMujer3"))
TunicaMagoCaosMujer = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "TunicaMagoCaosMujer"))
TunicaMagoCaosEnanoMujer = val(GetVar(DatPath & "Armadas.dat", "JerarquiaI", "TunicaMagoCaosEnanoMujer"))
'-------------

'2da jerarquía
JIIArmaduraImperial1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "ArmaduraImperial1"))
JIIArmaduraImperial2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "ArmaduraImperial2"))
JIIArmaduraImperial3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "ArmaduraImperial3"))
JIITunicaMagoImperial = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "TunicaMagoImperial"))
JIITunicaMagoImperialEnanos = val(GetVar(DatPath & "Server.ini", "JerarquiaII", "TunicaMagoImperialEnanos"))

JIIArmaduraCaos1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "ArmaduraCaos1"))
JIIArmaduraCaos2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "ArmaduraCaos2"))
JIIArmaduraCaos3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "ArmaduraCaos3"))
JIITunicaMagoCaos = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "TunicaMagoCaos"))
JIITunicaMagoCaosEnanos = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "TunicaMagoCaosEnanos"))

JIIArmaduraImpMujer1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "ArmaduraImpMujer1"))
JIIArmaduraImpMujer2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "ArmaduraImpMujer2"))
JIIArmaduraImpMujer3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "ArmaduraImpMujer3"))
JIITunicaMagoImpMujer = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "TunicaMagoImpMujer"))
JIITunicaMagoImpEnanosMujer = val(GetVar(DatPath & "Server.ini", "JerarquiaII", "TunicaMagoImpEnanosMujer"))

JIIArmaduraCaosMujer1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "ArmaduraCaosMujer1"))
JIIArmaduraCaosMujer2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "ArmaduraCaosMujer2"))
JIIArmaduraCaosMujer3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "ArmaduraCaosMujer3"))
JIITunicaMagoCaosMujer = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "TunicaMagoCaosMujer"))
JIITunicaMagoCaosEnanoMujer = val(GetVar(DatPath & "Armadas.dat", "JerarquiaII", "TunicaMagoCaosEnanoMujer"))
'-----------

'3era Jerarquia
JIIIArmaduraImperial1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "ArmaduraImperial1"))
JIIIArmaduraImperial2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "ArmaduraImperial2"))
JIIIArmaduraImperial3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "ArmaduraImperial3"))
JIIITunicaMagoImperial = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "TunicaMagoImperial"))
JIIITunicaMagoImperialEnanos = val(GetVar(DatPath & "Server.ini", "JerarquiaIII", "TunicaMagoImperialEnanos"))

JIIIArmaduraCaos1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "ArmaduraCaos1"))
JIIIArmaduraCaos2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "ArmaduraCaos2"))
JIIIArmaduraCaos3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "ArmaduraCaos3"))
JIIITunicaMagoCaos = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "TunicaMagoCaos"))
JIIITunicaMagoCaosEnanos = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "TunicaMagoCaosEnanos"))

JIIIArmaduraImpMujer1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "ArmaduraImpMujer1"))
JIIIArmaduraImpMujer2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "ArmaduraImpMujer2"))
JIIIArmaduraImpMujer3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "ArmaduraImpMujer3"))
JIIITunicaMagoImpMujer = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "TunicaMagoImpMujer"))
JIIITunicaMagoImpEnanosMujer = val(GetVar(DatPath & "Server.ini", "JerarquiaIII", "TunicaMagoImpEnanosMujer"))

JIIIArmaduraCaosMujer1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "ArmaduraCaosMujer1"))
JIIIArmaduraCaosMujer2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "ArmaduraCaosMujer2"))
JIIIArmaduraCaosMujer3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "ArmaduraCaosMujer3"))
JIIITunicaMagoCaosMujer = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "TunicaMagoCaosMujer"))
JIIITunicaMagoCaosEnanoMujer = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIII", "TunicaMagoCaosEnanoMujer"))
'-----------

'4ta Jerarquía
JIVArmaduraImperial1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "ArmaduraImperial1"))
JIVArmaduraImperial2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "ArmaduraImperial2"))
JIVArmaduraImperial3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "ArmaduraImperial3"))
JIVTunicaMagoImperial = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "TunicaMagoImperial"))
JIVTunicaMagoImperialEnanos = val(GetVar(DatPath & "Server.ini", "JerarquiaIV", "TunicaMagoImperialEnanos"))

JIVArmaduraCaos1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "ArmaduraCaos1"))
JIVArmaduraCaos2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "ArmaduraCaos2"))
JIVArmaduraCaos3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "ArmaduraCaos3"))
JIVTunicaMagoCaos = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "TunicaMagoCaos"))
JIVTunicaMagoCaosEnanos = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "TunicaMagoCaosEnanos"))

JIVArmaduraImpMujer1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "ArmaduraImpMujer1"))
JIVArmaduraImpMujer2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "ArmaduraImpMujer2"))
JIVArmaduraImpMujer3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "ArmaduraImpMujer3"))
JIVTunicaMagoImpMujer = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "TunicaMagoImpMujer"))
JIVTunicaMagoImpEnanosMujer = val(GetVar(DatPath & "Server.ini", "JerarquiaIV", "TunicaMagoImpEnanosMujer"))

JIVArmaduraCaosMujer1 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "ArmaduraCaosMujer1"))
JIVArmaduraCaosMujer2 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "ArmaduraCaosMujer2"))
JIVArmaduraCaosMujer3 = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "ArmaduraCaosMujer3"))
JIVTunicaMagoCaosMujer = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "TunicaMagoCaosMujer"))
JIVTunicaMagoCaosEnanoMujer = val(GetVar(DatPath & "Armadas.dat", "JerarquiaIV", "TunicaMagoCaosEnanoMujer"))
'-------------

End Sub
'[/KEVIN]

Public Sub CargarUnMapa(Map As Integer, archmap As String, archinf As String)
Dim Dm As Long
Dim TM As TileMap, TI As TileInf
Dim X As Integer, Y As Integer
Dim LoopC As Integer
Dim npcfile As String

    Dm = MAPCargaMapa(archmap, archinf)
    If Dm = 0 Then
        Debug.Print "kk " & Map
    End If
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            Call MAPLeeMapa(Dm, TM, TI)
            
            '.dat file
            'Get #1, , MapData(Map, X, Y).Blocked
            MapData(Map, X, Y).Blocked = TM.bloqueado
            
'            For LoopC = 1 To 4
'                'Get #1, , MapData(Map, X, Y).Graphic(LoopC)
'                MapData(Map, X, Y).Graphic(LoopC) = TM.grafs(LoopC)
'                If TM.grafs(LoopC) <> 0 Then
'                    TM.grafs(LoopC) = TM.grafs(LoopC)
'                End If
'            Next LoopC
            MapData(Map, X, Y).Graphic(1) = TM.grafs1
            MapData(Map, X, Y).Graphic(2) = TM.grafs2
            MapData(Map, X, Y).Graphic(3) = TM.grafs3
            MapData(Map, X, Y).Graphic(4) = TM.grafs4
            
            'Get #1, , MapData(Map, X, Y).trigger
            'Get #1, , TempInt
            MapData(Map, X, Y).trigger = TM.trigger
            If TM.trigger <> 0 Then
                TM.trigger = TM.trigger
            End If
            
            '.inf file
            'Get #2, , MapData(Map, X, Y).TileExit.Map
            'Get #2, , MapData(Map, X, Y).TileExit.X
            'Get #2, , MapData(Map, X, Y).TileExit.Y
            
            MapData(Map, X, Y).TileExit.Map = TI.dest_mapa
            MapData(Map, X, Y).TileExit.X = TI.dest_x
            MapData(Map, X, Y).TileExit.Y = TI.dest_y
            
            'Get and make NPC
            'Get #2, , MapData(Map, X, Y).NpcIndex
            MapData(Map, X, Y).NpcIndex = TI.npc
            
            If MapData(Map, X, Y).NpcIndex > 0 Then
                
                If MapData(Map, X, Y).NpcIndex > 499 Then
                        npcfile = DatPath & "NPCs-HOSTILES.dat"
                Else
                        npcfile = DatPath & "NPCs.dat"
                End If
                
                'Si el npc debe hacer respawn en la pos
                'original la guardamos
                If val(GetVar(npcfile, "NPC" & MapData(Map, X, Y).NpcIndex, "PosOrig")) = 1 Then
                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                    Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                Else
                    MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                End If
                
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Map = Map
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.X = X
                Npclist(MapData(Map, X, Y).NpcIndex).Pos.Y = Y
                
                Call MakeNPCChar(ToNone, 0, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
            End If

            'Get and make Object
            'Get #2, , MapData(Map, X, Y).OBJInfo.ObjIndex
            'Get #2, , MapData(Map, X, Y).OBJInfo.Amount
            MapData(Map, X, Y).OBJInfo.ObjIndex = TI.obj_ind
            MapData(Map, X, Y).OBJInfo.Amount = TI.obj_cant
            
            'Space holder for future expansion (Objects, ect.
            'Get #2, , DummyInt
            'Get #2, , DummyInt
        
        Next X
    Next Y

    Call MAPCierraMapa(Dm)
    ''Close #1
    ''Close #2

  
    MapInfo(Map).Name = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Name")
    MapInfo(Map).Music = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "MusicNum")
    MapInfo(Map).StartPos.Map = val(ReadField(1, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.X = val(ReadField(2, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
    MapInfo(Map).StartPos.Y = val(ReadField(3, GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "StartPos"), 45))
    
    If val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Pk")) = 0 Then
        MapInfo(Map).Pk = True
    Else
        MapInfo(Map).Pk = False
    End If
    
    
    MapInfo(Map).Terreno = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Terreno")

    MapInfo(Map).Zona = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Zona")
    
    MapInfo(Map).Restringir = GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "Restringir")
    
    MapInfo(Map).BackUp = val(GetVar(App.Path & MapPath & "Mapa" & Map & ".dat", "Mapa" & Map, "BACKUP"))

End Sub


