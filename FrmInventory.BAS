VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmInventory 
   Caption         =   "X-EL QUEST"
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11295
   OleObjectBlob   =   "FrmInventory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Chris Brogly
'Description: Simple RPG in Excel
'Purpose: I dno?
'See LICENSE file (Apache 2.0)

Option Explicit

Public Sub ShowStuff()
    Call ShowArmor
    Call ShowWeapon
    Call ShowPotions
End Sub

Private Sub ShowArmor()
    ImgArmor.Picture = LoadPicture(ActiveWorkbook.path & "\inventory\" & Game.strWepName & ".gif")
End Sub

Private Sub ShowWeapon()
    ImgWeapon.Picture = LoadPicture(ActiveWorkbook.path & "\inventory\" & Game.strArmorName & ".gif")
End Sub

Private Sub ShowPotions()

    ImgPot1.Picture = Nothing
    ImgPot2.Picture = Nothing
    ImgPot3.Picture = Nothing
    ImgPot4.Picture = Nothing
    ImgPot5.Picture = Nothing

    If Game.intPotionCount = 1 Then
        ImgPot1.Picture = LoadPicture(ActiveWorkbook.path & "\inventory\" & "potion" & ".gif")
    ElseIf Game.intPotionCount = 2 Then
        ImgPot2.Picture = LoadPicture(ActiveWorkbook.path & "\inventory\" & "potion" & ".gif")
    ElseIf Game.intPotionCount = 3 Then
        ImgPot3.Picture = LoadPicture(ActiveWorkbook.path & "\inventory\" & "potion" & ".gif")
    ElseIf Game.intPotionCount = 4 Then
        ImgPot4.Picture = LoadPicture(ActiveWorkbook.path & "\inventory\" & "potion" & ".gif")
    ElseIf Game.intPotionCount = 5 Then
        ImgPot5.Picture = LoadPicture(ActiveWorkbook.path & "\inventory\" & "potion" & ".gif")
    End If
End Sub

Private Sub ImgPot1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not ImgPot1.Picture Is Nothing Then
        Game.intPotionCount = Game.intPotionCount - 1
        Game.intPlayerHP = Game.intPlayerLevel * Game.HP_PER_LEVEL
        ImgPot1.Picture = Nothing
    End If
End Sub

Private Sub ImgPot2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not ImgPot2.Picture Is Nothing Then
        Game.intPotionCount = Game.intPotionCount - 1
        Game.intPlayerHP = Game.intPlayerLevel * Game.HP_PER_LEVEL
        ImgPot2.Picture = Nothing
    End If
End Sub

Private Sub ImgPot3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not ImgPot3.Picture Is Nothing Then
        Game.intPotionCount = Game.intPotionCount - 1
        Game.intPlayerHP = Game.intPlayerLevel * Game.HP_PER_LEVEL
        ImgPot3.Picture = Nothing
    End If
End Sub

Private Sub ImgPot4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not ImgPot4.Picture Is Nothing Then
        Game.intPotionCount = Game.intPotionCount - 1
        Game.intPlayerHP = Game.intPlayerLevel * Game.HP_PER_LEVEL
        ImgPot4.Picture = Nothing
    End If
End Sub

Private Sub ImgPot5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not ImgPot5.Picture Is Nothing Then
        Game.intPotionCount = Game.intPotionCount - 1
        Game.intPlayerHP = Game.intPlayerLevel * Game.HP_PER_LEVEL
        ImgPot5.Picture = Nothing
    End If
End Sub

Private Sub UserForm_Click()

End Sub
