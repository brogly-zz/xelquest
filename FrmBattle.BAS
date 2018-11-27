VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmBattle 
   Caption         =   "BATTLE LOG"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9630
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
Public fled As Boolean

Private Sub cmdAttemptToFlee_Click()
    Dim chance As Integer
    
    If chance < 15 Then
        'flee
        MsgBox "You successfully fled!"
        fled = True
    Else
        MsgBox "You tried to flee and didnt XD"
    End If
    
End Sub

Private Sub cmdContinue_Click()
    Me.Hide
End Sub

Public Sub StartUp()
    fled = False
    turn = 0
    TxtBattleLog.Text = ""
End Sub

Public Sub WriteLine(ln As String)
    TxtBattleLog.Text = TxtBattleLog.Text & vbCrLf & "Turn #" & turn & ": " & ln & vbCrLf
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
End Sub

Public Sub NextTurn()
    turn = turn + 1
    TxtBattleLog.Text = ""
End Sub

Private Sub UserForm_Terminate()
    Me.Hide
End Sub
