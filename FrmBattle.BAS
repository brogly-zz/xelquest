VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmBattle 
   Caption         =   "BATTLE LOG"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9750
   OleObjectBlob   =   "FrmBattle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmBattle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private turn As Integer

Private Sub cmdAttemptToFlee_Click()
    
End Sub

Private Sub cmdContinue_Click()
    Me.Hide
End Sub

Public Sub StartUp()
    turn = 1
    TxtBattleLog.Text = ""
End Sub

Public Sub WriteLine(ln As String)
    TxtBattleLog.Text = TxtBattleLog.Text & vbCrLf & "Turn #" & turn & ": " & ln & vbCrLf
    TxtBattleLog.SelStart = Len(TxtBattleLog.Text)
End Sub

Public Sub WriteHealthBar(hp As Integer, maxHp As Integer, name As String)
    Dim percent As Integer
    percent = Int((hp / maxHp) * 100)
    
    Dim healthBar As String
    
    Dim i As Integer
    For i = 0 To percent Step 10
        healthBar = healthBar & "="
    Next i
    
    For i = percent To 100 Step 10
        healthBar = healthBar & " "
    Next i

    TxtBattleLog.Text = TxtBattleLog.Text & vbCrLf & name & " HP" & ": " & vbCrLf & healthBar & " (" & hp & " / " & maxHp & ")" & vbCrLf
    TxtBattleLog.SelStart = Len(TxtBattleLog.Text)
End Sub

Public Sub NextTurn()
    turn = turn + 1
End Sub

Private Sub UserForm_Terminate()
    Me.Hide
End Sub
