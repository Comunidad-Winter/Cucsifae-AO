VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAntiRoboPjs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Clase para evitar el robo de Personajes por medio
'del uso de fuerza bruta en Argentum-Online.
'Por Alejandro Santos.
'
'REVISIONES:
' 25-5-2003:Alejo: Creación de la clase

Private Type tARPJ
    Nombre As String
    N As Long
    Veces As Long
End Type

Private m_lastindex As Long
Private m_tiempoprohibido As Long
Private m_maxelems As Long
Private m_lista() As tARPJ

Public Function PuedeLogear(Nombre As String) As Boolean
Dim I As Long
Dim sUNombre As String

sUNombre = UCase(Nombre)

For I = 1 To m_lastindex
    If sUNombre = m_lista(I).Nombre Then
        If m_lista(I).N > 0 Then
            m_lista(I).Veces = m_lista(I).Veces + 1
            PuedeLogear = False
            Exit Function
        End If
    End If
Next I
PuedeLogear = True

End Function

Public Function Agrega(Nombre As String) As Boolean
Dim I As Long, F As Long

F = IIf(m_lastindex < m_maxelems, m_lastindex + 1, m_lastindex)

For I = 1 To F
    If m_lista(I).N <= 0 Then
        m_lista(I).N = m_tiempoprohibido
        m_lista(I).Nombre = UCase(Nombre)
        m_lista(I).Veces = 0
        
        If I > m_lastindex Then m_lastindex = I
        
        Agrega = True
        Exit Function
    End If
Next I
Agrega = False

End Function

Public Sub PasarSegundo()
Dim I As Long

For I = 1 To m_lastindex
    If m_lista(I).N > 0 Then
        If m_lista(I).Veces > 0 Then
            m_lista(I).N = m_lista(I).N + Int(m_lista(I).Veces / m_tiempoprohibido)
            m_lista(I).Veces = (m_lista(I).Veces Mod m_tiempoprohibido) - 1
            If m_lista(I).Veces < 0 Then m_lista(I).Veces = 0
        Else
            m_lista(I).N = m_lista(I).N - 1
        End If
    End If
    
    If m_lista(I).N <= 0 Then
        m_lista(I).N = 0
        m_lista(I).Nombre = ""
        m_lista(I).Veces = 0
        
        'If I = m_lastindex Then
        Do While m_lista(m_lastindex).N = 0 And m_lastindex > 1
            m_lastindex = m_lastindex - 1
        Loop
        'End If
    End If
Next I

End Sub

Private Sub Class_Initialize()
Dim I As Long

m_maxelems = 400
m_lastindex = 1
m_tiempoprohibido = 3
ReDim m_lista(1 To m_maxelems)

For I = 1 To m_maxelems
    m_lista(I).N = 0
    m_lista(I).Nombre = ""
    m_lista(I).Veces = 0
Next I

End Sub

Public Property Get TiempoProhibido() As Long
TiempoProhibido = m_tiempoprohibido
End Property

Public Property Let TiempoProhibido(ByVal lNewValue As Long)
m_tiempoprohibido = lNewValue
End Property

Public Property Get MaxElems() As Long
MaxElems = m_maxelems
End Property

Public Property Let MaxElems(ByVal lNewValue As Long)
m_maxelems = lNewValue
ReDim Preserve m_lista(1 To m_maxelems)
End Property
