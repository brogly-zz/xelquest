Attribute VB_Name = "Game"
'Author: Chris Brogly
'Description: Simple RPG in Excel
'Purpose: I dno?
'See LICENSE file (Apache 2.0)

Option Explicit

'player attributes
Public strPlayerName As String
Public intPlayerHP As Integer
Public intPlayerLevel As Integer
Public intPlayerXP As Integer
Public intPlayerAtk As Integer
Public intPlayerDef As Integer
Public strWepName As String
Public intWepStr As Integer
Public intPotionCount As Integer

'player position on the game map
'added X/Y since these are sort of like pixels
Public intCurColX As Integer
Public intCurRowY As Integer

'names of all of the enemy graphics
Public arrEnemies As Variant

'useful constants
Const COUNT_ENEMIES = 7
Const MAP_WIDTH = 50
Const MAP_HEIGHT = 50

Const NAME_BOSS = "boss"
Const NAME_GIBS = "gibs"

Const XP_PER_LEVEL = 100
Const HP_PER_LEVEL = 10

Const PERCENT_OF_BATTLE = 40

Sub StartGame_Click()
    Call Main
End Sub

'main entry point of the model
Sub Main()

    'get different values from Rnd each time
    Randomize

    'put all the names of enemies in an array for random encounters
    arrEnemies = Array("ai", "bsod", "document", "error", "paperclip", "pie", "user friendly interface")
    
    'abstract the implementation of key features by organizing them into subs
    Call Reset
    Call SetupNewPlayer
    Call SetupGameScreen
    Call FrmCharacterCreator.Show
    Call CreateGameWorld
    Call UpdateLabels
    Call RunGame
    
    GameSheet.cmdInventory.Visible = True

End Sub

'reset the graphics/player values
Sub Reset()
    GameSheet.LblHeroName = ""
    GameSheet.LblHP = ""
    GameSheet.LblSTR = ""
    GameSheet.LblDEF = ""
    GameSheet.LblLevel = ""
    GameSheet.LblXP = ""
    GameSheet.ImgFace.Picture = Nothing
    GameSheet.ImgAccessory.Picture = Nothing
    GameSheet.ImgHair.Picture = Nothing
    
    GameSheet.cmdDown.Visible = False
    GameSheet.cmdUp.Visible = False
    GameSheet.cmdLeft.Visible = False
    GameSheet.cmdRight.Visible = False
    GameSheet.cmdInventory.Visible = False
    
    strPlayerName = ""
    intPlayerHP = 0
    intPlayerLevel = 0
    intPlayerAtk = 0
    intPlayerDef = 0
    intPlayerXP = 0
    strWepName = "slap"
    intWepStr = 1
    
    intCurColX = 1
    intCurRowY = 1
    
    DrawScene "title"
End Sub

'this is a comment about this sub
Sub SetupNewPlayer()
    strPlayerName = InputBox("Enter your name traveller (otherwise the game might crash):", "Enter your name Traveller")
    intPlayerHP = 10
    intPlayerLevel = 1
    intPlayerAtk = 3
    intPlayerDef = 2
    intCurColX = 1
    intCurRowY = 1
    strWepName = "slap"
    intPotionCount = 3
    MsgBox "Welcome to X-EL QUEST, " & strPlayerName, vbInformation, "You have started X-EL QUEST"
End Sub

Sub UpdateLabels()
    GameSheet.LblHeroName = strPlayerName
    GameSheet.LblLevel = intPlayerLevel
    GameSheet.LblHP = intPlayerHP
    GameSheet.LblSTR = intPlayerAtk
    GameSheet.LblDEF = intPlayerDef
    GameSheet.LblXP = intPlayerXP & " / " & XP_PER_LEVEL * intPlayerLevel
End Sub

Sub SetupGameScreen()
    GameSheet.Activate
    'ActiveSheet.Range("A1:BM64") 'Game Screen will be 64X64px
    ActiveSheet.Columns("A:BM").ColumnWidth = 1
    ActiveSheet.Rows("1:64").RowHeight = 8
End Sub

Sub RunGame()
    Call CheckAdjecentTiles
    Call DrawGameMap
    DrawScene strWepName, "pickups"
End Sub

Sub DrawGameMap()
    Dim tileName As String
    tileName = GameMap.Cells(intCurRowY, intCurColX).Value
    DrawScene tileName
End Sub

Sub DrawGibs()
    DrawScene NAME_GIBS, "enemies"
End Sub

Sub LevelUp()
    intPlayerXP = intPlayerXP + Int((XP_PER_LEVEL * intPlayerLevel) * Rnd + 1)
    
    If intPlayerXP >= (intPlayerLevel * XP_PER_LEVEL) Then
        MsgBox "You have levellled OOP!", vbInformation, "YAY"
        intPlayerLevel = intPlayerLevel + 1
        intPlayerAtk = intPlayerAtk * intPlayerLevel
        intPlayerDef = intPlayerDef * intPlayerLevel
        intPlayerHP = HP_PER_LEVEL * intPlayerLevel
    End If
End Sub

'movement
Sub GoUp()
    intCurRowY = intCurRowY - 1
    Call DrawGameMap
    Call CheckAdjecentTiles
    Call RandomEncounter
    Call UpdateLabels
    DrawScene strWepName, "pickups"
End Sub

Sub GoDown()
    intCurRowY = intCurRowY + 1
    Call DrawGameMap
    Call CheckAdjecentTiles
    Call RandomEncounter
    Call UpdateLabels
    DrawScene strWepName, "pickups"
End Sub

Sub GoLeft()
    intCurColX = intCurColX - 1
    Call DrawGameMap
    Call CheckAdjecentTiles
    Call RandomEncounter
    Call UpdateLabels
    DrawScene strWepName, "pickups"
End Sub

Sub GoRight()
    intCurColX = intCurColX + 1
    Call DrawGameMap
    Call CheckAdjecentTiles
    Call RandomEncounter
    Call UpdateLabels
    DrawScene strWepName, "pickups"
End Sub

Sub RandomEncounter()
    Dim chance As Integer
    chance = Int((100) * Rnd + 1)
    
    If chance < PERCENT_OF_BATTLE Then
        
        'figure out what this random encounter will be
        Dim whichEnemy As Integer
        whichEnemy = Int((COUNT_ENEMIES) * Rnd)
        Dim strEnemyName As String
        strEnemyName = arrEnemies(whichEnemy)
    
        MsgBox "You have encountered an enemy " & UCase(strEnemyName) & "!", vbInformation, "BATTLE"
        
        'draw the enemy graphics on top of everything else
        DrawScene strEnemyName, "enemies"
        
        'draw weapon on top of enemy
        DrawScene strWepName, "pickups"
        
        'enemy stats
        Dim enemyHP As Integer
        Dim enemyMaxHP As Integer
        Dim enemyATK As Integer
        Dim enemyDEF As Integer
        
        enemyHP = Int((HP_PER_LEVEL * intPlayerLevel) * Rnd + 1)
        enemyMaxHP = enemyHP
        enemyATK = Int((intPlayerAtk) * Rnd + 1)
        enemyDEF = Int((intPlayerDef) * Rnd + 1)
        
        'temp variables for damage values
        Dim dmgFromPlayer As Integer
        Dim dmgFromEnemy As Integer
        Dim enemyDefended As Integer
        Dim playerDefended As Integer
        
        Call FrmBattle.StartUp
        FrmBattle.WriteHealthBar intPlayerHP, HP_PER_LEVEL * intPlayerLevel, strPlayerName
        
        'main battle loop, keep going until someone is TOAST
        Do Until enemyHP <= 0 Or intPlayerHP <= 0
            Call FrmBattle.NextTurn
        
            'the player's attack
            dmgFromPlayer = Int((intPlayerAtk) * Rnd + 1)
            enemyDefended = Int((dmgFromPlayer) * Rnd)
            enemyHP = (enemyHP - dmgFromPlayer) + enemyDefended
            FrmBattle.WriteLine "You did " & dmgFromPlayer & " damage with Weapon " & UCase(strWepName) & "!"
            
            If enemyDefended > 0 Then
                FrmBattle.WriteLine "Enemy" & UCase(strEnemyName) & " managed to block " & enemyDefended & " damage though!"
            End If
            
            FrmBattle.WriteHealthBar enemyHP, enemyMaxHP, strEnemyName
            
            If enemyHP <= 0 Then Exit Do
            
            'the enemy attack
            dmgFromEnemy = Int((enemyATK) * Rnd + 1)
            playerDefended = Int((dmgFromEnemy) * Rnd)
            intPlayerHP = (intPlayerHP - dmgFromEnemy) + playerDefended
            FrmBattle.WriteLine "Enemy " & UCase(strEnemyName) & " did " & dmgFromEnemy & " damage to you!"
            
            If playerDefended > 0 Then
                FrmBattle.WriteLine "You managed to block " & playerDefended & " damage though!"
            End If
            
            FrmBattle.WriteHealthBar intPlayerHP, HP_PER_LEVEL * intPlayerLevel, strPlayerName
            
            GameSheet.LblHP = intPlayerHP
            
            Call FrmBattle.Show
        Loop
        
        'check who won
        If enemyHP <= 0 Then
            MsgBox UCase(strEnemyName) & " is DED", vbInformation, "YAY"
            MsgBox "You WINNED!", vbInformation, "BATTLE WINNED!"
            Call DrawGameMap
            Call DrawGibs
            DrawScene strWepName, "pickups"
            Call LevelUp
        Else
            MsgBox "U r DED. Plz play again. Or Not. I dno", vbInformation, "OH NO"
            Call Reset
            End
        End If
        
        'loot that you could pickup
        Call Loot
    End If
End Sub

Sub Loot()
    'TODO
End Sub

Sub CheckAdjecentTiles()
    GameSheet.cmdUp.Visible = False
    GameSheet.cmdDown.Visible = False
    GameSheet.cmdLeft.Visible = False
    GameSheet.cmdRight.Visible = False
    
    If intCurRowY - 1 > 0 Then
        If GameMap.Cells(intCurRowY - 1, intCurColX).Value <> "" Then
            GameSheet.cmdUp.Visible = True
        End If
    End If

    If GameMap.Cells(intCurRowY + 1, intCurColX).Value <> "" Then
        GameSheet.cmdDown.Visible = True
    End If

    If intCurColX - 1 > 0 Then
        If GameMap.Cells(intCurRowY, intCurColX - 1).Value <> "" Then
            GameSheet.cmdLeft.Visible = True
        End If
    End If

    If GameMap.Cells(intCurRowY, intCurColX + 1).Value <> "" Then
        GameSheet.cmdRight.Visible = True
    End If

End Sub

Sub DrawScene(tileName As String, Optional folder As String = "map")
    Dim i As Integer
    Dim j As Integer
    Dim dataLine As String
    
    Dim tilePath As String
    tilePath = ActiveWorkbook.path & "/" & folder & "/" & tileName & ".csv"
    
    Open tilePath For Input As #1
    Dim ColX As String
    Dim RowY As String
    Dim ColNum As Integer
    Dim ASCIINum As Integer
    
    ASCIINum = 65
    RowY = 1
    
    'go through each line in the tile file, which represents a row
    Do Until EOF(1)
        'start at A
        ColNum = 0
        ColX = Chr(ASCIINum)
        
        'read the current line of the file (dataLine)
        Line Input #1, dataLine
        
        'split this line up to get the "pixel" ARGB values
        Dim lineArray() As String
        lineArray = Split(dataLine, ",")
        
        For i = 0 To UBound(lineArray) - 1
            'iterate through this pixel value to break it up into A, R, G, and B
            Dim pixel As String
            pixel = lineArray(i)
            
            'need RGB colors
            Dim A As String
            Dim R As String
            Dim G As String
            Dim B As String
            
            A = ""
            R = ""
            G = ""
            B = ""
            
            'read pixel data
            For j = 1 To Len(pixel)
                If Len(A) <> 3 Then
                    A = A + Mid(pixel, j, 1)
                ElseIf Len(R) <> 3 Then
                    R = R + Mid(pixel, j, 1)
                ElseIf Len(G) <> 3 Then
                    G = G + Mid(pixel, j, 1)
                ElseIf Len(B) <> 3 Then
                    B = B + Mid(pixel, j, 1)
                End If
            Next j
            
            'if the pixel is opaque, then draw it, skip otherwise
            If A <> "000" Then
                'setup pixel color
                Dim pos As String
                pos = Trim(ColX) & Trim(Str(RowY))
                ActiveSheet.Range(pos).Interior.Color = RGB(Int(R), Int(G), Int(B))
            End If
            
            'after drawing this pixel go to the next column
            ColNum = ColNum + 1
            
            'get the address of a cell (things get weird after column Z)
            Dim strColPos() As String
            strColPos = Split(Cells(RowY, ColNum).Address, "$")
            ColX = strColPos(0) & strColPos(1)
        Next i
        
        RowY = RowY + 1
    Loop
    Close #1
End Sub

Sub CreateGameWorld()
    Dim i As Integer
    Dim j As Integer
    For i = 1 To MAP_WIDTH
        For j = 1 To MAP_HEIGHT
            GameMap.Cells(j, i).Value = "forest" & Int(Rnd() * 20 + 1)
        Next j
    Next i
End Sub

Sub CharacterCreatorTester()
    FrmCharacterCreator.Show
End Sub
