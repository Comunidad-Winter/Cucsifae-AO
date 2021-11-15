VERSION 5.00
Begin VB.Form frmGmGodTool 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herramientas de GM"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command54 
      Caption         =   "Gmsg"
      Height          =   195
      Left            =   3240
      TabIndex        =   71
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command53 
      Caption         =   "Ver Skills de"
      Height          =   195
      Left            =   2520
      TabIndex        =   70
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command52 
      Caption         =   "Modificar Skills"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2520
      TabIndex        =   69
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command44 
      Caption         =   "Ip del Nick"
      Height          =   195
      Left            =   120
      TabIndex        =   68
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command50 
      Caption         =   "Inventario PJ"
      Height          =   195
      Left            =   2520
      TabIndex        =   67
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command51 
      Caption         =   "Informacion PJ"
      Height          =   195
      Left            =   2520
      TabIndex        =   66
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command49 
      Caption         =   "Encarcelar"
      Height          =   195
      Left            =   120
      TabIndex        =   65
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Frame fraOtros 
      BackColor       =   &H00000000&
      Caption         =   "Otros"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   57
      Top             =   7800
      Width           =   4455
      Begin VB.CommandButton Command56 
         Caption         =   "RFlag Torneo"
         Height          =   195
         Left            =   3120
         TabIndex        =   73
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command55 
         Caption         =   "Verinsc"
         Height          =   195
         Left            =   3120
         TabIndex        =   72
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command48 
         Caption         =   "Show SOS"
         Height          =   195
         Left            =   1680
         TabIndex        =   64
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command47 
         Caption         =   "TelepLoc"
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command46 
         Caption         =   "Invisible"
         Height          =   195
         Left            =   1680
         TabIndex        =   62
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command45 
         Caption         =   "Reset Npc Inv"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command43 
         Caption         =   "Destruir (OBJ)"
         Height          =   195
         Left            =   3120
         TabIndex        =   60
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command42 
         Caption         =   "Bloquear"
         Height          =   195
         Left            =   1680
         TabIndex        =   59
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command41 
         Caption         =   "MassKill"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Echar de la Armada"
      Height          =   195
      Left            =   2520
      TabIndex        =   55
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Carisma"
      Height          =   195
      Left            =   2520
      TabIndex        =   54
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Desbanear a TODOS"
      Enabled         =   0   'False
      Height          =   195
      Left            =   1560
      TabIndex        =   53
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtNPC 
      Height          =   285
      Left            =   2400
      TabIndex        =   52
      Top             =   7440
      Width           =   615
   End
   Begin VB.CommandButton Command38 
      Caption         =   "Sumonear NPC"
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Criminales Matados"
      Height          =   195
      Left            =   2520
      TabIndex        =   50
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Ciudadanos Matados"
      Height          =   195
      Left            =   120
      TabIndex        =   49
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Defensa"
      Height          =   195
      Left            =   2520
      TabIndex        =   48
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Max HIT"
      Height          =   195
      Left            =   120
      TabIndex        =   47
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Max Mana"
      Height          =   195
      Left            =   2520
      TabIndex        =   46
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Max Stamina"
      Height          =   195
      Left            =   120
      TabIndex        =   45
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Max HP"
      Height          =   195
      Left            =   120
      TabIndex        =   44
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Inteligencia"
      Height          =   195
      Left            =   120
      TabIndex        =   43
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Agilidad"
      Height          =   195
      Left            =   2520
      TabIndex        =   42
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Constitución"
      Height          =   195
      Left            =   120
      TabIndex        =   41
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Fuerza"
      Height          =   195
      Left            =   2520
      TabIndex        =   40
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox txtCant 
      Height          =   285
      Left            =   3960
      TabIndex        =   39
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox txtObj 
      Height          =   285
      Left            =   2400
      TabIndex        =   38
      Top             =   7080
      Width           =   615
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Crear Objeto"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Echar a TODOS"
      Height          =   195
      Left            =   1560
      TabIndex        =   35
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Hacer criminal a"
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Hacer ciudadano a"
      Height          =   195
      Left            =   2520
      TabIndex        =   33
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Frame fraServer 
      BackColor       =   &H00000000&
      Caption         =   "Server"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   6240
      Width           =   4455
      Begin VB.CommandButton Command40 
         Caption         =   "Limpiar"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Width           =   4215
      End
      Begin VB.CommandButton Command22 
         Caption         =   "World Save"
         Height          =   195
         Left            =   2880
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Apagar"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Resetear"
         Height          =   195
         Left            =   1440
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtMod 
      Height          =   285
      Left            =   1080
      TabIndex        =   28
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Cabeza"
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Cuerpo"
      Height          =   195
      Left            =   2520
      TabIndex        =   26
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Experiencia"
      Height          =   195
      Left            =   2520
      TabIndex        =   25
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Nivel"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Oro"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton Command12 
      Caption         =   "IpBan a"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2520
      TabIndex        =   21
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Curar a"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Transportar al"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtYN 
      Height          =   285
      Left            =   4080
      TabIndex        =   18
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtXN 
      Height          =   285
      Left            =   3240
      TabIndex        =   16
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtMN 
      Height          =   285
      Left            =   2400
      TabIndex        =   14
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Revivir a"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Rmsg"
      Height          =   195
      Left            =   1680
      TabIndex        =   11
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Smsg"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox txtMsg 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   5640
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Sumonear a"
      Height          =   195
      Left            =   2520
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Donde"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ir a"
      Height          =   195
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Echar a"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Desbanear a"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Banear a"
      Height          =   195
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Cantidad:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   37
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Modificar:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Y:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "X:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Mapa:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Mensaje:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Nick:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmGmGodTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
SendData ("/ban porgmtool@" & txtNick)
End Sub

Private Sub Command10_Click()
SendData ("/rajar " & txtNick)
End Sub

Private Sub Command11_Click()
SendData ("/revivir " & txtNick)
End Sub

Private Sub Command12_Click()
'SendData ("/ipban " & stxtNickBuffer)
End Sub

Private Sub Command13_Click()
SendData ("/telep " & txtNick & " " & txtMN & " " & txtXN & " " & txtYN)
End Sub

Private Sub Command14_Click()
SendData ("/curar " & stxtNick)
End Sub

Private Sub Command15_Click()
SendData ("/mod " & txtNick & " oro " & txtMod)
End Sub

Private Sub Command16_Click()
SendData ("/mod " & txtNick & " level " & txtMod)
End Sub

Private Sub Command17_Click()
SendData ("/mod " & txtNick & " exp " & txtMod)
End Sub

Private Sub Command18_Click()
SendData ("/mod " & txtNick & " body " & txtMod)
End Sub

Private Sub Command19_Click()
SendData ("/mod " & txtNick & " head " & txtMod)
End Sub

Private Sub Command2_Click()
SendData ("/unban " & txtNick)
End Sub

Private Sub Command20_Click()

If MsgBox("¿Estás Seguro que querés apagar el server?", vbYesNo + vbCritical + vbDefaultBotton2) = vbNo Then
Cancel = True
Else
SendData ("/apagar")
End If

End Sub

Private Sub Command21_Click()

If MsgBox("¿Estás Seguro que querés apagar el server?", vbYesNo + vbCritical + vbDefaultBotton2) = vbNo Then
Cancel = True
Else
SendData ("/resetear")
End If

End Sub

Private Sub Command22_Click()
SendData ("/dobackup")
End Sub

Private Sub Command23_Click()
SendData ("/perdonc " & stxtNickBuffer)
End Sub

Private Sub Command24_Click()
SendData ("/conden " & txtNick)
End Sub

Private Sub Command25_Click()
SendData ("/bootall")
End Sub

Private Sub Command26_Click()

If txtCant.Text = "" Then
MsgBox ("Debes escribir una cantidad.")
Exit Sub
End If

SendData ("/crearobj " & txtObj & " " & txtCant)

End Sub

Private Sub Command27_Click()
SendData ("/mod " & txtNick & " fuerza " & txtMod)
End Sub

Private Sub Command28_Click()
SendData ("/mod " & txtNick & " const " & txtMod)
End Sub

Private Sub Command29_Click()
SendData ("/mod " & txtNick & " agil " & txtMod)
End Sub

Private Sub Command3_Click()
SendData ("/echar " & txtNick)
End Sub

Private Sub Command30_Click()
SendData ("/mod " & txtNick & " intel " & txtMod)
End Sub

Private Sub Command31_Click()
SendData ("/mod " & txtNick & " maxhp " & txtMod)
End Sub

Private Sub Command32_Click()
SendData ("/mod " & txtNick & " maxsta " & txtMod)
End Sub

Private Sub Command33_Click()
SendData ("/mod " & txtNick & " maxman " & txtMod)
End Sub

Private Sub Command34_Click()
SendData ("/mod " & txtNick & " maxhit " & txtMod)
End Sub

Private Sub Command35_Click()
SendData ("/mod " & txtNick & " def " & txtMod)
End Sub

Private Sub Command36_Click()
SendData ("/mod " & txtNick & " ciu " & txtMod)
End Sub

Private Sub Command37_Click()
SendData ("/mod " & txtNick & " cri " & txtMod)
End Sub

Private Sub Command38_Click()
SendData ("/acc " & txtNPC)
End Sub

Private Sub Command4_Click()
SendData ("/ira " & txtNick)
End Sub

Private Sub Command40_Click()
SendData ("/limpiar")
End Sub

Private Sub Command41_Click()
SendData ("/masskill")
End Sub

Private Sub Command42_Click()
SendData ("/bloq")
End Sub

Private Sub Command43_Click()
SendData ("/dest")
End Sub

Private Sub Command44_Click()
SendData ("/ipnick " & txtNick)
End Sub

Private Sub Command45_Click()
SendData ("/resetinv")
End Sub

Private Sub Command46_Click()
SendData ("/invisible")
End Sub

Private Sub Command47_Click()
SendData ("/teleploc")
End Sub

Private Sub Command48_Click()
SendData ("/show sos")
End Sub

Private Sub Command49_Click()
SendData ("/carcel " & txtNick & " 15")
End Sub

Private Sub Command5_Click()
SendData ("/donde " & txtNick)
End Sub

Private Sub Command50_Click()
SendData ("/inv " & txtNick)
End Sub

Private Sub Command51_Click()
SendData ("/info " & txtNick)
End Sub

Private Sub Command52_Click()
frmModSkills.Show
End Sub

Private Sub Command53_Click()
SendData ("$USKILL " & txtNick)
End Sub

Private Sub Command54_Click()
SendData ("/gmsg " & txtMsg)
End Sub

Private Sub Command55_Click()
SendData ("/verinsc")
End Sub

Private Sub Command56_Click()
SendData ("/rflag")
End Sub

Private Sub Command6_Click()
SendData ("/sum " & txtNick)
End Sub

Private Sub Command7_Click()
SendData ("/smsg " & txtMsg)
End Sub


Private Sub Command8_Click()
SendData ("/mod " & txtNick & " caris " & txtMod)
End Sub

Private Sub Command9_Click()
SendData ("/rmsg " & txtMsg)
End Sub

Private Sub txtCant_Change()
stxtCantBuffer = txtCant.Text
End Sub

Private Sub txtMod_Change()
stxtModBuffer = txtMod.Text
End Sub

Private Sub txtMsg_Change()
stxtMensajeBuffer = txtMsg.Text
End Sub

Private Sub txtNick_Change()
stxtNickBuffer = txtNick.Text
End Sub

Private Sub txtNPC_Change()
stxtNPCBuffer = txtNPC.Text
End Sub

Private Sub txtObj_Change()
stxtObjBuffer = txtObj.Text
End Sub
