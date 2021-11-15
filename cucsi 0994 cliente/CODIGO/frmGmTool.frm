VERSION 5.00
Begin VB.Form frmGmTool 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herramientas de GM"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command53 
      Caption         =   "Ver Skills de"
      Height          =   195
      Left            =   2520
      TabIndex        =   35
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command44 
      Caption         =   "Ip del Nick"
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command50 
      Caption         =   "Inventario PJ"
      Height          =   195
      Left            =   2520
      TabIndex        =   33
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command51 
      Caption         =   "Informacion PJ"
      Height          =   195
      Left            =   2520
      TabIndex        =   32
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command49 
      Caption         =   "Encarcelar"
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Frame fraOtros 
      BackColor       =   &H00000000&
      Caption         =   "Otros"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   4080
      Width           =   4455
      Begin VB.CommandButton Command48 
         Caption         =   "Show SOS"
         Height          =   195
         Left            =   1680
         TabIndex        =   30
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command47 
         Caption         =   "TelepLoc"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command46 
         Caption         =   "Invisible"
         Height          =   195
         Left            =   3120
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command45 
         Caption         =   "Reset Npc Inv"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command42 
         Caption         =   "Bloquear"
         Height          =   195
         Left            =   1680
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraServer 
      BackColor       =   &H00000000&
      Caption         =   "Server"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   4455
      Begin VB.CommandButton Command40 
         Caption         =   "Limpiar"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   4215
      End
      Begin VB.CommandButton Command22 
         Caption         =   "World Save"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   4215
      End
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
      Left            =   2520
      TabIndex        =   11
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Smsg"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txtMsg 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   2640
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
      Enabled         =   0   'False
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
      Top             =   2640
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
Attribute VB_Name = "frmGmTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
SendData ("/ban " & txtNick)
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
'SendData ("/ciudadano " & stxtNickBuffer)
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
SendData ("/mod " & txtNickB & " ciu " & txtMod)
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
