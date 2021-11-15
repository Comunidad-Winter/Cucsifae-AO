VERSION 5.00
Begin VB.Form frmServidor 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Servidor"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command21 
      Caption         =   "Respawn Recursos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   21
      Top             =   4080
      Width           =   3555
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Actualizar INI"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   20
      Top             =   3840
      Width           =   3555
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Actualizar Quests"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   19
      Top             =   3600
      Width           =   3555
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Actualizar Spawn List"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   18
      Top             =   3360
      Width           =   3555
   End
   Begin VB.TextBox txtMacros 
      Height          =   285
      Left            =   3240
      TabIndex        =   17
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Cambiar Intervalos de los Macros a:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5520
      Width           =   2895
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Actualizar Dat de Trabajadores"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   15
      Top             =   3120
      Width           =   3555
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Echar a Todos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   14
      Top             =   2880
      Width           =   3555
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Unban All"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   13
      Top             =   2640
      Width           =   3555
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Debug listening socket"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   12
      Top             =   2400
      Width           =   3555
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Debug Npcs"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   11
      Top             =   2160
      Width           =   3555
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reiniciar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   10
      Top             =   720
      Width           =   3555
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ReSpawn Guardias en posiciones originales"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   9
      Top             =   480
      Width           =   3555
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Stats de los slots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   3555
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Trafico"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   3555
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Reload Lista Nombres Prohibidos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   3555
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Actualizar hechizos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   3555
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Configurar intervalos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   3555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar objetos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   3555
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hacer un Backup del mundo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   315
      TabIndex        =   2
      Top             =   4680
      Width           =   3645
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cargar BackUp del mundo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   315
      TabIndex        =   1
      Top             =   5040
      Width           =   3645
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   5880
      Width           =   1065
   End
   Begin VB.Shape Shape3 
      Height          =   4305
      Left            =   120
      Top             =   120
      Width           =   3975
   End
   Begin VB.Shape Shape2 
      Height          =   855
      Left            =   120
      Top             =   4560
      Width           =   3975
   End
End
Attribute VB_Name = "frmServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Command1_Click()

On Error GoTo eh
    Me.MousePointer = 11
    FrmStat.Titu.Caption = "Procesando objetos..."
    FrmStat.Show
    Call LoadOBJData
    Me.MousePointer = 0
    MsgBox "OBJETOS ACTUALIZADOS!!"
Exit Sub

eh:
Call LogError("Error en Actualizar Objetos")


End Sub

Private Sub Command10_Click()
frmTrafic.Show
End Sub

Private Sub Command11_Click()
frmConID.Show
End Sub

Private Sub Command12_Click()
frmDebugNpc.Show
End Sub

Private Sub Command13_Click()
frmDebugSocket.Visible = True
End Sub

Private Sub Command14_Click()

Dim LoopC As Integer
For LoopC = 1 To LastUser
CloseSocket (LoopC)
Next LoopC

End Sub

Private Sub Command15_Click()
On Error Resume Next

Dim Fn As String
Dim cad$
Dim N As Integer, k As Integer

Fn = App.Path & "\logs\GenteBanned.log"

If FileExist(Fn, vbNormal) Then
    N = FreeFile
    Open Fn For Input Shared As #N
    Do While Not EOF(N)
        k = k + 1
        Input #N, cad$
        Call UnBan(cad$)
        
    Loop
    Close #N
    MsgBox "Se han habilitado " & k & " personajes."
    Kill Fn
End If




End Sub

Private Sub Command16_Click()
Call LoadArmasHerreria
Call LoadArmadurasHerreria
Call LoadObjCarpintero
Call LoadObjDruida
Call LoadObjSastre
End Sub

Private Sub Command17_Click()
If Not Numeric(txtMacros.Text) Then
MsgBox ("Error")
Exit Sub
End If

Dim Interval As Integer
Interval = txtMacros.Text

If Interval > 5000 Then
MsgBox ("Debes poner un número menor a 5000!")
Exit Sub
End If

Call SendData(ToAll, 0, 0, "CMR" & txtMacros.Text)
End Sub

Private Sub Command18_Click()
Call CargarSpawnList
End Sub

Private Sub Command19_Click()
Call CargarQuests
End Sub

Private Sub Command2_Click()
frmServidor.Visible = False
End Sub

Private Sub Command20_Click()
HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
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
End Sub

Private Sub Command21_Click()
Dim RespawnMateriales As Integer

Call SendData(ToAll, 0, 0, "||%%%%RECURSOS RESPAWN...%%%%" & FONTTYPE_INFO)

For RespawnMateriales = 1 To NumObjDatas
        If ObjData(RespawnMateriales).ObjType = 22 Then
            If ObjData(RespawnMateriales).Materiales <= 0 Then _
            ObjData(RespawnMateriales).Materiales = 50000
        End If
    
        If ObjData(RespawnMateriales).ObjType = 4 Then
            If ObjData(RespawnMateriales).Materiales <= 0 Then _
            ObjData(RespawnMateriales).Materiales = 50000
            If ObjData(RespawnMateriales).Materiales2 <= 0 Then _
            ObjData(RespawnMateriales).Materiales2 = 25000
        End If
    Next RespawnMateriales
    
Call SendData(ToAll, 0, 0, "||%%%%RECURSOS RESPAWN DONE%%%%" & FONTTYPE_INFO)
    
End Sub

Private Sub Command22_Click()
FastMapLoad = val(GetVar(IniPath & "Server.ini", "INIT", "FastMapLoad"))

FastObjLoad = val(GetVar(IniPath & "Server.ini", "INIT", "FastObjLoad"))

Rain = val(GetVar(IniPath & "Server.ini", "INIT", "Lluvia"))

AntiSH = val(GetVar(IniPath & "Server.ini", "INIT", "AntiSH"))
End Sub

Private Sub Command3_Click()
If MsgBox("¡¡Atencion!! Si reinicia el servidor puede provocar la perdida de datos de los usarios. ¿Desea reiniciar el servidor de todas maneras?", vbYesNo) = vbYes Then
    Me.Visible = False
    Call Restart
End If
End Sub

Private Sub Command4_Click()
On Error GoTo eh
    Me.MousePointer = 11
    FrmStat.Titu.Caption = "Procesando mapas..."
    FrmStat.Show
    Call DoBackUp
    Me.MousePointer = 0
    MsgBox "WORLDSAVE OK!!"
Exit Sub
eh:
Call LogError("Error en WORLDSAVE")
End Sub

Private Sub Command5_Click()

'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

FrmStat.Show

If FileExist(App.Path & "\logs\errores.log", vbNormal) Then Kill App.Path & "\logs\errores.log"
If FileExist(App.Path & "\logs\connect.log", vbNormal) Then Kill App.Path & "\logs\Connect.log"
If FileExist(App.Path & "\logs\HackAttemps.log", vbNormal) Then Kill App.Path & "\logs\HackAttemps.log"
If FileExist(App.Path & "\logs\Asesinatos.log", vbNormal) Then Kill App.Path & "\logs\Asesinatos.log"
If FileExist(App.Path & "\logs\Resurrecciones.log", vbNormal) Then Kill App.Path & "\logs\Resurrecciones.log"
If FileExist(App.Path & "\logs\Teleports.Log", vbNormal) Then Kill App.Path & "\logs\Teleports.Log"

  
frmMain.Socket1.Cleanup
frmMain.Socket2(0).Cleanup
  
Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next
  

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call CargarBackUp
Call LoadOBJData


frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = puerto
frmMain.Socket1.Listen

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

End Sub

Private Sub Command6_Click()
Call ReSpawnOrigPosNpcs
End Sub

Private Sub Command7_Click()
FrmInterv.Show
End Sub

Private Sub Command8_Click()
Call CargarHechizos
End Sub

Private Sub Command9_Click()
Call CargarForbidenWords
End Sub

Private Sub Form_Deactivate()
frmServidor.Visible = False
End Sub

