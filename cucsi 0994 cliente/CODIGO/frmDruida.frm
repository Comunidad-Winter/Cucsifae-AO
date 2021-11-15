VERSION 5.00
Begin VB.Form frmDruida 
   Caption         =   "Druida"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCrear 
      Caption         =   "Crear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "1"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ListBox lstPociones 
      Height          =   2205
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Atención!!!   Nunca poner una cantidad mayor a 10000 porque sino saltará un error y se cerrará el juego!!"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   855
   End
End
Attribute VB_Name = "frmDruida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCrear_Click()
On Error Resume Next
Dim stxtCantBuffer As String
stxtCantBuffer = txtCantidad.Text

Call SendData("DCI" & ObjDruida(lstPociones.ListIndex) & " " & stxtCantBuffer)

Unload Me
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

