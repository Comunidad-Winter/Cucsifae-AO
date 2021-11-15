VERSION 5.00
Begin VB.Form Main 
   Caption         =   "INI"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Leer GameINI"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Escribir GameINI"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()

Dim GameIni As tGameIni
GameIni.Fx = 0
GameIni.Musica = 0
GameIni.Name = "morgolock"
GameIni.Password = "mipassword"
GameIni.tip = 1
GameIni.Puerto = 7666
GameIni.DirGraficos = "Graficos"
GameIni.DirMusica = "Midi"
GameIni.DirSonidos = "Wav"
GameIni.DirMapas = "Mapas"
GameIni.NumeroDeBMPs = 12300
GameIni.NumeroMapas = 300

Call IniciarCabecera(MiCabecera)
Call EscribirGameIni(GameIni)

End Sub

Private Sub Command3_Click()
Dim gm As tGameIni
gm = LeerGameIni()
MsgBox ("Fx:" & gm.Fx & "  --- " & " Musica:" & gm.Musica _
& " ---- Nombre:" & gm.Name & " ----- pass:" & gm.Password _
& " ---- tip:" & gm.tip & " ---- puerto:" & gm.Puerto & " --- " & _
gm.DirGraficos & "  " & gm.DirMapas & "  " & gm.DirMusica & "  " & gm.DirSonidos)
End Sub

