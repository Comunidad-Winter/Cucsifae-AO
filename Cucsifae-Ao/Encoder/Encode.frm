VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Encoder"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Encode Mapas"
      Height          =   615
      Left            =   1065
      TabIndex        =   5
      Top             =   3765
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Encode Cascos"
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   3075
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Encode Tips"
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Encode FXs"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Encode Cuerpos"
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encode Cabezas"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
On Error GoTo errhandler
Call CargarHeads
Dim n As Integer, i As Integer
n = FreeFile
Open App.Path & "\init\Cabezas.ind" For Binary As #n
'Escribimos la cabecera
Put #n, , MiCabecera
'Guardamos las cabezas
Put #n, , Numheads

For i = 1 To Numheads
    Put #n, , HeadData(i)
Next i
Close #n
Call MsgBox("Listo, encode ok!!")

Exit Sub
errhandler:
Call MsgBox("Error en cabeza " & i)
End Sub

Private Sub Command2_Click()
On Error GoTo errhandler
Call CargarBodys
Dim n As Integer, i As Integer
n = FreeFile
Open App.Path & "\init\Personajes.ind" For Binary As #n
'Escribimos la cabecera
Put #n, , MiCabecera
'Guardamos las cabezas
Put #n, , NumCuerpos

For i = 1 To NumCuerpos
    Put #n, , CuerpoData(i)
Next i

Close #n
Call MsgBox("Listo, encode ok!!")

Exit Sub
errhandler:
Call MsgBox("Error en cuerpo " & i)

End Sub

Private Sub Command3_Click()
On Error GoTo errhandler
Call CargarFxs
Dim n As Integer, i As Integer
n = FreeFile
Open App.Path & "\init\Fxs.ind" For Binary As #n
'Escribimos la cabecera
Put #n, , MiCabecera
'Guardamos las cabezas
Put #n, , NumFxs

For i = 1 To NumFxs
    Put #n, , FxData(i)
Next i
Close #n
Call MsgBox("Listo, encode ok!!")

Exit Sub
errhandler:
Call MsgBox("Error en fx " & i)

End Sub

Private Sub Command4_Click()
On Error GoTo errhandler
Call CargarTips

Dim n As Integer, i As Integer
n = FreeFile

Open App.Path & "\init\Tips.ayu" For Binary As #n
'Escribimos la cabecera
Put #n, , MiCabecera
'Guardamos las cabezas
Put #n, , NumTips

For i = 1 To NumTips
    Put #n, , Tips(i)
Next i

Close #n
Call MsgBox("Listo, encode ok!!")

Exit Sub
errhandler:
Call MsgBox("Error en tip " & i)

End Sub

Private Sub Command5_Click()
On Error GoTo errhandler
Call CargarCascos

Dim n As Integer, i As Integer
n = FreeFile
Open App.Path & "\init\Cascos.ind" For Binary As #n
'Escribimos la cabecera
Put #n, , MiCabecera
'Guardamos las cabezas
Put #n, , NumCascos

For i = 1 To NumCascos
    Put #n, , CascoSData(i)
Next i
Close #n
Call MsgBox("Listo, encode ok!!")

Exit Sub
errhandler:
Call MsgBox("Error en casco " & i)

End Sub

Private Sub Command6_Click()


On Error GoTo errhandler
Call CargarMapas

Dim n As Integer, i As Integer
n = FreeFile
Open App.Path & "\init\FK.ind" For Binary As #n
'Escribimos la cabecera
Put #n, , MiCabecera
'Guardamos las cabezas
Put #n, , NumMapas

For i = 1 To NumMapas
    Put #n, , Mapas(i)
Next i
Close #n

Call MsgBox("Listo, encode ok!!")

Exit Sub

errhandler:
Call MsgBox("Error en casco " & i)
End Sub

Private Sub Form_Load()
Call IniciarCabecera(MiCabecera)
End Sub

Private Sub CargarBodys()
Dim loopc As Integer


NumCuerpos = Val(GetVar(App.Path & "\encode\Body.dat", "INIT", "NumBodies"))

ReDim CuerpoData(0 To NumCuerpos + 1) As tIndiceCuerpo

For loopc = 1 To NumCuerpos
    CuerpoData(loopc).Body(1) = Val(GetVar(App.Path & "\encode\body.dat", "Body" & loopc, "WALK1"))
    CuerpoData(loopc).Body(2) = Val(GetVar(App.Path & "\encode\body.dat", "Body" & loopc, "WALK2"))
    CuerpoData(loopc).Body(3) = Val(GetVar(App.Path & "\encode\body.dat", "Body" & loopc, "WALK3"))
    CuerpoData(loopc).Body(4) = Val(GetVar(App.Path & "\encode\body.dat", "body" & loopc, "WALK4"))
    CuerpoData(loopc).HeadOffsetX = Val(GetVar(App.Path & "\encode\body.dat", "body" & loopc, "HeadOffsetX"))
    CuerpoData(loopc).HeadOffsetY = Val(GetVar(App.Path & "\encode\body.dat", "body" & loopc, "HeadOffsety"))
Next loopc


End Sub


Private Sub CargarHeads()
Dim loopc As Integer


Numheads = Val(GetVar(App.Path & "\encode\Head.dat", "INIT", "NumHeads"))

ReDim HeadData(0 To Numheads + 1) As tIndiceCabeza

For loopc = 1 To Numheads
    HeadData(loopc).Head(1) = Val(GetVar(App.Path & "\encode\Head.dat", "Head" & loopc, "Head1"))
    HeadData(loopc).Head(2) = Val(GetVar(App.Path & "\encode\Head.dat", "Head" & loopc, "Head2"))
    HeadData(loopc).Head(3) = Val(GetVar(App.Path & "\encode\Head.dat", "Head" & loopc, "Head3"))
    HeadData(loopc).Head(4) = Val(GetVar(App.Path & "\encode\Head.dat", "Head" & loopc, "Head4"))
Next loopc


End Sub

Private Sub CargarCascos()
Dim loopc As Integer

NumCascos = Val(GetVar(App.Path & "\encode\Cascos.dat", "INIT", "NumCascos"))

ReDim CascoSData(0 To NumCascos + 1) As tIndiceCabeza

For loopc = 1 To NumCascos
    CascoSData(loopc).Head(1) = Val(GetVar(App.Path & "\encode\Cascos.dat", "Casco" & loopc, "Head1"))
    CascoSData(loopc).Head(2) = Val(GetVar(App.Path & "\encode\Cascos.dat", "Casco" & loopc, "Head2"))
    CascoSData(loopc).Head(3) = Val(GetVar(App.Path & "\encode\Cascos.dat", "Casco" & loopc, "Head3"))
    CascoSData(loopc).Head(4) = Val(GetVar(App.Path & "\encode\Cascos.dat", "Casco" & loopc, "Head4"))
Next loopc


End Sub

Private Sub CargarMapas()
Dim loopc As Integer

NumMapas = Val(GetVar(App.Path & "\encode\mapas.dat", "INIT", "NumMaps"))

ReDim Mapas(0 To NumMapas + 1) As Byte

For loopc = 1 To NumMapas
    Mapas(loopc) = Val(GetVar(App.Path & "\encode\mapas.dat", "Map" & loopc, "Lluvia"))
Next loopc

End Sub

Private Sub CargarFxs()
Dim loopc As Integer
NumFxs = Val(GetVar(App.Path & "\encode\fx.dat", "INIT", "NumFxs"))

ReDim FxData(0 To NumFxs + 1) As tIndiceFx

For loopc = 1 To NumFxs
    FxData(loopc).Animacion = Val(GetVar(App.Path & "\encode\fx.dat", "Fx" & loopc, "Animacion"))
    FxData(loopc).OffsetX = Val(GetVar(App.Path & "\encode\fx.dat", "Fx" & loopc, "OffsetX"))
    FxData(loopc).OffsetY = Val(GetVar(App.Path & "\encode\fx.dat", "Fx" & loopc, "OffsetY"))
Next loopc

End Sub

Private Sub CargarTips()
Dim loopc As Integer
NumTips = Val(GetVar(App.Path & "\encode\tips.dat", "INIT", "Tips"))

ReDim Tips(0 To NumTips + 1) As String * 255

For loopc = 1 To NumTips
    Tips(loopc) = GetVar(App.Path & "\encode\tips.dat", "Tip" & loopc, "Tip")
Next loopc

End Sub

