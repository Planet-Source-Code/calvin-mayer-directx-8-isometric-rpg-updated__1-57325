VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DirectXText As New clsText
Public map As New clsIsometricMap
Dim MapX As Integer, MapY As Integer
Dim PlayerSpeedX As Integer, PlayerSpeedY As Integer

Private Sub Form_Keydown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: bRunning = False
        Case vbKeyUp
            PlayerSpeedY = -5
        Case vbKeyDown
            PlayerSpeedY = 5
        Case vbKeyLeft
            PlayerSpeedX = -5
        Case vbKeyRight
            PlayerSpeedX = 5
        Case vbKey1
            map.DrawingMap = Not map.DrawingMap
        Case vbKey2
            map.DrawingBuildings = Not map.DrawingBuildings
        Case vbKey3
            map.DrawingPlayer = Not map.DrawingPlayer
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: bRunning = False
        Case vbKeyUp
            PlayerSpeedY = 0
        Case vbKeyDown
            PlayerSpeedY = 0
        Case vbKeyLeft
            PlayerSpeedX = 0
        Case vbKeyRight
            PlayerSpeedX = 0
    End Select
End Sub

Private Sub Form_Load()
    
    Game_Initializing
    
    Do While bRunning
        Game_MainLoop
    Loop
    
    Game_CleanUp
    
End Sub

Sub Render()
    ClearDevice 0
    BeginScene
    
    'drawPosX = posX + (X * (TILE_WIDTH / 2)) - (Y * (TILE_WIDTH / 2))
    'drawPosY = posY + (Y * (TILE_HEIGHT / 2)) + (X * (TILE_HEIGHT / 2))
    
    'Me.Caption = "PlayerPos.X = " & map.GetPlayerX & " PlayerPos.Y = " & map.GetPlayerY
    map.MovePlayer PlayerSpeedX, PlayerSpeedY
    
    map.DrawMap , , True

    DirectXText.Render FPS_Current, 0, 0
    EndScene
    
    UpdateFps
    PresentDevice
End Sub

Private Sub Form_Terminate()
    bRunning = False
    Game_CleanUp
End Sub

Sub Game_Initializing()
    Me.Show
    
    bRunning = InitDirectX(1024, 768, Me, True, 32, False)
    
    map.LoadMapFile App.Path & "\MapFiles\town.map"
    map.LoadBuildings App.Path & "\MapFiles\theBuildings.map"
    map.InitPlayer App.Path & "\Images\Characters\Player.bmp", 50, 108, 1, 0, 0
    map.CreateNPCs App.Path & "\MapFiles\theNPCs.map"
    
    
    MapX = ScreenWidth / 2
    MapY = ScreenHeight / 2 - map.GetMapHeight * map.GetTileHeight / 2
    
    DirectXText.Create
End Sub

Sub Game_MainLoop()
    DoEvents
    Render
End Sub


Sub Game_CleanUp()
    map.DestroyTextures
    CleanUpDirectX
    Unload Me
    End
End Sub
