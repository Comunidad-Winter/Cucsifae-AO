VERSION 5.00
Begin VB.Form frmSastre 
   Caption         =   "Sastre"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   4935
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Text            =   "1"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ListBox lstRopas 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4665
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Realizar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      MouseIcon       =   "frmSastre.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3240
      Width           =   1710
   End
   Begin VB.CommandButton Command4 
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
      Height          =   435
      Left            =   240
      MouseIcon       =   "frmSastre.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3240
      Width           =   1710
   End
   Begin VB.Label Label2 
      Caption         =   "Atención!!!   Nunca poner una cantidad mayor a 10000 porque sino saltará un error y se cerrará el juego!!"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   2880
      Width           =   735
   End
End
Attribute VB_Name = "frmSastre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
On Error Resume Next
Dim stxtCantBuffer As String
stxtCantBuffer = txtCantidad.Text

Call SendData("SCR" & ObjSastre(lstRopas.ListIndex) & " " & stxtCantBuffer)

Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

