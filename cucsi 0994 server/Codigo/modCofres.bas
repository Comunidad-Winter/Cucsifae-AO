Attribute VB_Name = "modCofres"
'MODULO PROGRAMADO POR NEB
'Kevin Birmingham
'kbneb@hotmail.com

Sub IniciarDepCofre(ByVal UserIndex As Integer)
On Error GoTo errhandler

Dim ObjInd As Integer
ObjInd = UserList(UserIndex).Flags.TargetObj

'Hacemos un Update del inventario del usuario
Call UpdateCofreInv(True, UserIndex, ObjInd, 0)
'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
SendData ToIndex, UserIndex, 0, "INITCOFRE"
UserList(UserIndex).Flags.Comerciando = True

errhandler:

End Sub

Sub UpdateCofreInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal CofreIndex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Byte
Dim NCofre As Integer
NCofre = ObjData(CofreIndex).CofreNro

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If Cofres(NCofre).Object(Slot).ObjIndex > 0 Then
        Call SendCofreObj(UserIndex, NCofre, Slot, Cofres(NCofre).Object(Slot))
    Else
        Call SendCofreObj(UserIndex, NCofre, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_COFREINVENTORY_SLOTS

        'Actualiza el inventario
        If Cofres(NCofre).Object(LoopC).ObjIndex > 0 Then
            Call SendCofreObj(UserIndex, NCofre, LoopC, Cofres(NCofre).Object(LoopC))
        Else
            
            Call SendCofreObj(UserIndex, NCofre, LoopC, NullObj)
            
        End If

    Next LoopC

End If

End Sub

Sub SendCofreObj(UserIndex As Integer, NroCofre As Integer, Slot As Byte, Object As UserOBJ)


Cofres(NroCofre).Object(Slot) = Object

If Object.ObjIndex > 0 Then

    Call SendData(ToIndex, UserIndex, 0, "SCO" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).Name & "," & Object.Amount & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
    & ObjData(Object.ObjIndex).ObjType & "," _
    & ObjData(Object.ObjIndex).MaxHIT & "," _
    & ObjData(Object.ObjIndex).MinHIT & "," _
    & ObjData(Object.ObjIndex).MaxDef)

Else

    Call SendData(ToIndex, UserIndex, 0, "SCO" & Slot & "," & "0" & "," & "(None)" & "," & "0" & "," & "0")

End If


End Sub

Sub UserCRetiraItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer)
On Error GoTo errhandler

Dim ObjInd As Integer
Dim NCofre As Integer
ObjInd = UserList(UserIndex).Flags.TargetObj
NCofre = ObjData(ObjInd).CofreNro


If Cantidad < 1 Then Exit Sub

       If Cofres(NCofre).Object(i).Amount > 0 Then
            If Cantidad > Cofres(NCofre).Object(i).Amount Then Cantidad = Cofres(NCofre).Object(i).Amount
            'Agregamos el obj que compro al inventario
            Call UserReciveCObj(UserIndex, NCofre, CInt(i), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el banco
            Call UpdateCofreInv(True, UserIndex, ObjInd, 0)
            'Actualizamos la ventana de comercio
            Call UpdateVentanaCofres(i, 0, UserIndex)
       End If



errhandler:

End Sub

Sub UserReciveCObj(ByVal UserIndex As Integer, ByVal NroCofre As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim Slot As Integer
Dim obji As Integer


If Cofres(NroCofre).Object(ObjIndex).Amount <= 0 Then Exit Sub

obji = Cofres(NroCofre).Object(ObjIndex).ObjIndex


'¿Ya tiene un objeto de este tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji And _
   UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
    
    Slot = Slot + 1
    If Slot > MAX_INVENTORY_SLOTS Then
        Exit Do
    End If
Loop

'Sino se fija por un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(ToIndex, UserIndex, 0, "||No podés tener mas objetos." & FONTTYPE_INFO)
                Exit Sub
            End If
        Loop
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If



'Mete el obj en el slot
If UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
    'Menor que MAX_INV_OBJS
    UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji
    UserList(UserIndex).Invent.Object(Slot).Amount = UserList(UserIndex).Invent.Object(Slot).Amount + Cantidad
    
    Call QuitarCofreInvItem(NroCofre, CByte(ObjIndex), Cantidad)
Else
    Call SendData(ToIndex, UserIndex, 0, "||No podés tener mas objetos." & FONTTYPE_INFO)
End If


End Sub

Sub QuitarCofreInvItem(ByVal NroCofre As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)



Dim ObjIndex As Integer
ObjIndex = Cofres(NroCofre).Object(Slot).ObjIndex

    'Quita un Obj

       Cofres(NroCofre).Object(Slot).Amount = Cofres(NroCofre).Object(Slot).Amount - Cantidad
        
        If Cofres(NroCofre).Object(Slot).Amount <= 0 Then
            Cofres(NroCofre).NroItems = Cofres(NroCofre).NroItems - 1
            Cofres(NroCofre).Object(Slot).ObjIndex = 0
            Cofres(NroCofre).Object(Slot).Amount = 0
        End If

    
    
End Sub

Sub UpdateVentanaCofres(ByVal Slot As Integer, ByVal NpcInv As Byte, ByVal UserIndex As Integer)
 
 
 Call SendData(ToIndex, UserIndex, 0, "COFREOK" & Slot & "," & NpcInv)
 
End Sub

Sub UserDepositaCItem(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

On Error GoTo errhandler

Dim ObjInd As Integer
Dim NCofre As Integer
ObjInd = UserList(UserIndex).Flags.TargetObj
NCofre = ObjData(ObjInd).CofreNro

If UserList(UserIndex).Invent.Object(Item).Amount > 0 And UserList(UserIndex).Invent.Object(Item).Equipped = 0 Then
            
            If Cantidad > 0 And Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).Amount
            'Agregamos el obj que compro al inventario
            Call UserDejaCObj(UserIndex, NCofre, CInt(Item), Cantidad)
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            'Actualizamos el inventario del banco
            Call UpdateCofreInv(True, UserIndex, ObjInd, 0)
            'Actualizamos la ventana del banco
            
            Call UpdateVentanaCofres(Item, 1, UserIndex)
            
End If

errhandler:

End Sub

Sub UserDejaCObj(ByVal UserIndex As Integer, ByVal NroCofre As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)

Dim Slot As Integer
Dim obji As Integer

If Cantidad < 1 Then Exit Sub

obji = UserList(UserIndex).Invent.Object(ObjIndex).ObjIndex

'¿Ya tiene un objeto de este tipo?
Slot = 1
Do Until Cofres(NroCofre).Object(Slot).ObjIndex = obji And _
         Cofres(NroCofre).Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
        
            If Slot > MAX_COFREINVENTORY_SLOTS Then
                Exit Do
            End If
Loop

'Sino se fija por un slot vacio antes del slot devuelto
If Slot > MAX_COFREINVENTORY_SLOTS Then
        Slot = 1
        Do Until Cofres(NroCofre).Object(Slot).ObjIndex = 0
            Slot = Slot + 1

            If Slot > MAX_COFREINVENTORY_SLOTS Then
                Call SendData(ToIndex, UserIndex, 0, "||No hay mas espacio en el cofre!!" & FONTTYPE_INFO)
                Exit Sub
                Exit Do
            End If
        Loop
        If Slot <= MAX_COFREINVENTORY_SLOTS Then Cofres(NroCofre).NroItems = Cofres(NroCofre).NroItems + 1
        
        
End If

If Slot <= MAX_COFREINVENTORY_SLOTS Then 'Slot valido
    'Mete el obj en el slot
    If Cofres(NroCofre).Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        
        'Menor que MAX_INV_OBJS
        Cofres(NroCofre).Object(Slot).ObjIndex = obji
        Cofres(NroCofre).Object(Slot).Amount = Cofres(NroCofre).Object(Slot).Amount + Cantidad
        
        Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)

    Else
        Call SendData(ToIndex, UserIndex, 0, "||El banco no puede cargar tantos objetos." & FONTTYPE_INFO)
    End If

Else
    Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
End If

End Sub

Sub SaveCofres()

On Error GoTo errhandler

Dim Cofre As Integer

For Cofre = 1 To NumeroCofres

Call SCofre(Cofre)
    
Next

Exit Sub

errhandler:
Call LogError("Error Grabando Cofres")


End Sub

Sub SCofre(ByVal Cofre As Integer)

Dim LoopC As Integer

Call WriteVar(DatPath & "Cofres.dat", "Cofre" & Cofre, "NroItems", val(Cofres(Cofre).NroItems))

For LoopC = 1 To MAX_COFREINVENTORY_SLOTS
    Call WriteVar(DatPath & "Cofres.dat", "Cofre" & Cofre, "Obj" & LoopC, Cofres(Cofre).Object(LoopC).ObjIndex & "-" & Cofres(Cofre).Object(LoopC).Amount)
Next LoopC

End Sub
