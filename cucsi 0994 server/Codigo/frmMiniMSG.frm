VERSION 5.00
Begin VB.Form frmMiniMSG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensaje a"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar BroadCast"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMiniMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SendData(ToIndex, TmpInx, 0, "!!Servidor: " & Text1.Text & ENDC)
End Sub

Private Sub Form_Load()
Me.Caption = "Mensaje a " & UserList(TmpInx).Name & " !!"
End Sub
