VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCharacterCreator 
   Caption         =   "X-EL QUEST"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360
   OleObjectBlob   =   "FrmCharacterCreator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmCharacterCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public idxFace As Integer
Public idxHair As Integer
Public idxAccessory As Integer

Const TOTAL_FACES = 4
Const TOTAL_HAIR = 5
Const TOTAL_ACCS = 3

Private Sub cmdFaceNext_Click()
    idxFace = (idxFace + 1) Mod (TOTAL_FACES + 1)
    If idxFace = 0 Then idxFace = 1
    ImgFace.Picture = LoadPicture(ActiveWorkbook.path & "\character\face" & idxFace & ".gif")
End Sub

Private Sub cmdFacePrev_Click()
    idxFace = idxFace - 1
    If idxFace <= 0 Then idxFace = TOTAL_FACES
    ImgFace.Picture = LoadPicture(ActiveWorkbook.path & "\character\face" & idxFace & ".gif")
End Sub

Private Sub cmdGlassesNext_Click()
    idxAccessory = (idxAccessory + 1) Mod (TOTAL_ACCS + 1)
    If idxAccessory = 0 Then
        ImgAccessory.Picture = Nothing
    Else
        ImgAccessory.Picture = LoadPicture(ActiveWorkbook.path & "\character\accs" & idxAccessory & ".gif")
    End If
End Sub

Private Sub cmdGlassesPrev_Click()
    idxAccessory = idxAccessory - 1
    If idxAccessory < 0 Then
        idxAccessory = TOTAL_ACCS
        ImgAccessory.Picture = LoadPicture(ActiveWorkbook.path & "\character\accs" & idxAccessory & ".gif")
    Else
        If idxAccessory = 0 Then
            ImgAccessory.Picture = Nothing
        Else
            ImgAccessory.Picture = LoadPicture(ActiveWorkbook.path & "\character\accs" & idxAccessory & ".gif")
        End If
    End If
End Sub

Private Sub cmdHairNext_Click()
    idxHair = (idxHair + 1) Mod (TOTAL_HAIR + 1)
    If idxHair = 0 Then
        ImgHair.Picture = Nothing
    Else
        ImgHair.Picture = LoadPicture(ActiveWorkbook.path & "\character\hair" & idxHair & ".gif")
    End If
End Sub

Private Sub cmdHairPrev_Click()
    idxHair = idxHair - 1
    If idxHair < 0 Then
        idxHair = TOTAL_HAIR
        ImgHair.Picture = LoadPicture(ActiveWorkbook.path & "\character\hair" & idxHair & ".gif")
    Else
        If idxHair = 0 Then
            ImgHair.Picture = Nothing
        Else
            ImgHair.Picture = LoadPicture(ActiveWorkbook.path & "\character\hair" & idxHair & ".gif")
        End If
    End If
End Sub

Private Sub cmdStart_Click()
    
    SetupCharacterPic
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    idxFace = 1
    idxHair = 0
    idxAccessory = 0
End Sub

Sub SetupCharacterPic()
    GameSheet.ImgFace.Picture = LoadPicture(ActiveWorkbook.path & "\character\face" & idxFace & ".gif")

    If idxHair = 0 Then
       GameSheet.ImgHair.Picture = Nothing
    Else
       GameSheet.ImgHair.Picture = LoadPicture(ActiveWorkbook.path & "\character\hair" & idxHair & ".gif")
    End If
        
    If idxAccessory = 0 Then
        GameSheet.ImgAccessory.Picture = Nothing
    Else
        GameSheet.ImgAccessory.Picture = LoadPicture(ActiveWorkbook.path & "\character\accs" & idxAccessory & ".gif")
    End If
End Sub
