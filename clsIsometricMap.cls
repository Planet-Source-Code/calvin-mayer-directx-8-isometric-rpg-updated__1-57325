VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIsometricMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright 2004 Calvin Mayer
Option Explicit

Private Const NUM_TILES = 16
Private Const TILE_GRASS = 0
Private Const TILE_DIRT = 1
Private Const TILE_WATER = 2
Private Const TILE_COBBLESTONE = 3
Private Const TILE_WOOD = 4
Private Const TILE_SNOW = 5
Private Const TILE_SNOWTOPGRASSBOTTOM = 6
Private Const TILE_SNOWBOTTOMGRASSTOP = 7
Private Const TILE_SNOWLEFTGRASSRIGHT = 8
Private Const TILE_SNOWRIGHTGRASSLEFT = 9
Private Const TILE_SNOWTOPDIRTBOTTOM = 10
Private Const TILE_SNOWBOTTOMDIRTTOP = 11
Private Const TILE_SNOWLEFTDIRTRIGHT = 12
Private Const TILE_SNOWRIGHTDIRTLEFT = 13
Private Const TILE_REDDIRT = 14
Private Const TILE_DEADGRASS = 15

Private Const NUM_WALLS = 4
Private Const WALL_BRICK1 = 0
Private Const WALL_BRICK2 = 1
Private Const WALL_BRICK3 = 2
Private Const WALL_BRICK4 = 3

Private Const TILE_WIDTH = 50 'width of the tile image
Private Const TILE_HEIGHT = 30 'height of the tile image

Private HalfTileWidth As Single
Private HalfTileHeight As Single

Private Const WALL_HEIGHT = 100

Private Const SIDE_NONE = -1
Private Const SIDE_TOP = 0
Private Const SIDE_BOTTOM = 1
Private Const SIDE_LEFT = 2
Private Const SIDE_RIGHT = 3
Private Const SIDE_TOPLEFT = 4
Private Const SIDE_TOPRIGHT = 5
Private Const SIDE_BOTTOMLEFT = 6
Private Const SIDE_BOTTOMRIGHT = 7


Private Type tMapTile
    Walkable As Boolean
    TileID As Integer
End Type

Private Type tPlayerOBJ
    Texture As New clsTexture
    Pos As Point
    PosOnTile As Point
    Speed As Single
    Width As Single
    Height As Single
End Type

Private Type tBuildingOBJ
    Name As String 'name of the house
    houseWidth As Integer
    houseHeight As Integer
    Door As Point 'the position of the door
    roofSlopesHoriz As Boolean 'holds the way the roof is pointing
    Pos As Point 'position the topleft corner will occupy
End Type

Private m_MapWidth As Long 'number of tiles horizontally
Private m_MapHeight As Long 'number of tiles vertically
Private m_MapTextures(0 To NUM_TILES - 1) As New clsTexture 'the different tiles in the map
Private m_WallTextures(0 To NUM_WALLS - 1) As New clsTexture
Private m_MapPos As Point

Private m_NumNPCs As Integer

Private m_NumBuildings As Integer

Dim m_Map() As tMapTile
Private m_Player As tPlayerOBJ
Private m_Buildings() As tBuildingOBJ
Private m_NPCs As New clsNPC

Public DrawingBuildings As Boolean
Public DrawingPlayer As Boolean
Public DrawingNPCs As Boolean
Public DrawingMap As Boolean

Public Sub InitPlayer(filepath As String, Width As Integer, Height As Integer, Speed As Single, PosX As Long, PosY As Long, Optional PosOnTileX As Long = -1, Optional PosOnTileY As Long = -1)
    HalfTileWidth = TILE_WIDTH / 2
    HalfTileHeight = TILE_HEIGHT / 2
    
    DrawingPlayer = True
    
    If PosOnTileX = -1 Then
        PosOnTileX = HalfTileWidth
    End If
    
    If PosOnTileY = -1 Then
        PosOnTileY = HalfTileHeight
    End If
    
    With m_Player
        .Texture.Init filepath, , , RGB(255, 0, 255)
        .Width = Width
        .Height = Height
        .Pos.x = PosX
        .Pos.y = PosY
        .Speed = Speed
        .PosOnTile.x = PosOnTileX
        .PosOnTile.y = PosOnTileY
    End With
End Sub

Public Sub MovePlayer(ByVal AmmountX As Long, ByVal AmmountY As Long)
    Dim lineY As Double
    
    With m_Player
        Select Case CheckTileBoundaries(AmmountX, AmmountY)
            Case SIDE_TOP
                If CheckWalkability(.Pos.x - 1, .Pos.y - 1) Then
                    If CheckForBuildings(SIDE_TOP) = False Then
                        .PosOnTile.x = .PosOnTile.x + AmmountX
                        .PosOnTile.y = .PosOnTile.y + AmmountY
                        .Pos.x = .Pos.x - 1
                        .Pos.y = .Pos.y - 1
                        .PosOnTile.y = TILE_HEIGHT - 1
                    End If
                End If
            Case SIDE_BOTTOM
                If CheckWalkability(.Pos.x + 1, .Pos.y + 1) Then
                    If CheckForBuildings(SIDE_BOTTOM) = False Then
                        .PosOnTile.x = .PosOnTile.x + AmmountX
                        .PosOnTile.y = .PosOnTile.y + AmmountY
                        .Pos.x = .Pos.x + 1
                        .Pos.y = .Pos.y + 1
                        .PosOnTile.y = 0
                    End If
                End If
            Case SIDE_LEFT
                If CheckWalkability(.Pos.x - 1, .Pos.y + 1) Then
                    If CheckForBuildings(SIDE_LEFT) = False Then
                        .PosOnTile.x = .PosOnTile.x + AmmountX
                        .PosOnTile.y = .PosOnTile.y + AmmountY
                        .Pos.x = .Pos.x - 1
                        .Pos.y = .Pos.y + 1
                        .PosOnTile.x = TILE_WIDTH - 1
                    End If
                End If
            Case SIDE_RIGHT
                If CheckWalkability(.Pos.x + 1, .Pos.y - 1) Then
                    If CheckForBuildings(SIDE_RIGHT) = False Then
                        .PosOnTile.x = .PosOnTile.x + AmmountX
                        .PosOnTile.y = .PosOnTile.y + AmmountY
                        .Pos.x = .Pos.x + 1
                        .Pos.y = .Pos.y - 1
                        .PosOnTile.x = 0
                    End If
                End If
            Case SIDE_TOPLEFT
                If CheckWalkability(.Pos.x - 1, .Pos.y) Then
                    If CheckForBuildings(SIDE_TOPLEFT) = False Then
                        .PosOnTile.x = .PosOnTile.x + AmmountX
                        .PosOnTile.y = .PosOnTile.y + AmmountY
                        .Pos.x = .Pos.x - 1
                        .PosOnTile.x = .PosOnTile.x + HalfTileWidth
                        .PosOnTile.y = .PosOnTile.y + HalfTileHeight
                    End If
                End If
            Case SIDE_TOPRIGHT
                If CheckWalkability(.Pos.x, .Pos.y - 1) Then
                    If CheckForBuildings(SIDE_TOPRIGHT) = False Then
                        .PosOnTile.x = .PosOnTile.x + AmmountX
                        .PosOnTile.y = .PosOnTile.y + AmmountY
                        .Pos.y = .Pos.y - 1
                        .PosOnTile.x = .PosOnTile.x - HalfTileWidth
                        .PosOnTile.y = .PosOnTile.y + HalfTileHeight
                    End If
                End If
            Case SIDE_BOTTOMLEFT
                If CheckWalkability(.Pos.x, .Pos.y + 1) Then
                    If CheckForBuildings(SIDE_BOTTOMLEFT) = False Then
                        .PosOnTile.x = .PosOnTile.x + AmmountX
                        .PosOnTile.y = .PosOnTile.y + AmmountY
                        .Pos.y = .Pos.y + 1
                        .PosOnTile.x = .PosOnTile.x + HalfTileWidth
                        .PosOnTile.y = .PosOnTile.y - HalfTileHeight
                    End If
                End If
            Case SIDE_BOTTOMRIGHT
                If CheckWalkability(.Pos.x + 1, .Pos.y) Then
                    If CheckForBuildings(SIDE_BOTTOMRIGHT) = False Then
                        .PosOnTile.x = .PosOnTile.x + AmmountX
                        .PosOnTile.y = .PosOnTile.y + AmmountY
                        .Pos.x = .Pos.x + 1
                        .PosOnTile.x = .PosOnTile.x - HalfTileWidth
                        .PosOnTile.y = .PosOnTile.y - HalfTileHeight
                    End If
                End If
            Case SIDE_NONE
                .PosOnTile.x = .PosOnTile.x + AmmountX
                .PosOnTile.y = .PosOnTile.y + AmmountY
        End Select
    End With
    
    '.Pos.Y = .Pos.Y - 1
    '.PosOnTile.X = .PosOnTile.X - HalfTileWidth
    '.PosOnTile.Y = lineY + HalfTileHeight
End Sub

Public Sub LoadMapFile(filepath As String)
    Dim x As Integer, y As Integer
    Dim FileNum As Integer
    
    DrawingMap = True
    
    FileNum = FreeFile
    
    Open filepath For Input As FileNum
        Input #FileNum, m_MapWidth
        Input #FileNum, m_MapHeight
        
        ReDim m_Map(0 To m_MapWidth - 1, 0 To m_MapHeight - 1)
        
        For y = 0 To m_MapHeight - 1
            For x = 0 To m_MapWidth - 1
                If EOF(FileNum) = False Then
                    Input #FileNum, m_Map(x, y).TileID
                Else
                    'm_Map(X, Y).TileID = Rnd() * 4
                    m_Map(x, y).TileID = TILE_GRASS
                End If
            Next
        Next
    Close
    
    m_MapTextures(TILE_GRASS).Init App.Path & "\Images\Tiles\Grass.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_DIRT).Init App.Path & "\Images\Tiles\Dirt.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_WATER).Init App.Path & "\Images\Tiles\Water1.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_COBBLESTONE).Init App.Path & "\Images\Tiles\Cobblestone1.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_WOOD).Init App.Path & "\Images\Tiles\Wood1.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_SNOW).Init App.Path & "\Images\Tiles\Snow.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_SNOWTOPGRASSBOTTOM).Init App.Path & "\Images\Tiles\snowTopGrassBottom.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_SNOWBOTTOMGRASSTOP).Init App.Path & "\Images\Tiles\snowBottomGrassTop.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_SNOWLEFTGRASSRIGHT).Init App.Path & "\Images\Tiles\snowLeftGrassRight.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_SNOWRIGHTGRASSLEFT).Init App.Path & "\Images\Tiles\snowRightGrassLeft.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_SNOWTOPDIRTBOTTOM).Init App.Path & "\Images\Tiles\snowTopDirtBottom.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_SNOWBOTTOMDIRTTOP).Init App.Path & "\Images\Tiles\snowBottomDirtTop.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_SNOWLEFTDIRTRIGHT).Init App.Path & "\Images\Tiles\snowLeftDirtRight.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_SNOWRIGHTDIRTLEFT).Init App.Path & "\Images\Tiles\snowRightDirtLeft.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_REDDIRT).Init App.Path & "\Images\Tiles\RedDirt.bmp", , , RGB(255, 0, 255)
    m_MapTextures(TILE_DEADGRASS).Init App.Path & "\Images\Tiles\DeadGrass.bmp", , , RGB(255, 0, 255)
    
    SetWalkable
    
    HalfTileWidth = TILE_WIDTH / 2
    HalfTileHeight = TILE_HEIGHT / 2
    
End Sub

Public Sub LoadBuildings(filepath As String)
    Dim I As Integer, j As Integer
    Dim x As Integer, y As Integer
    Dim curSpotInStr As Integer
    Dim FileNum As Integer
    
    DrawingBuildings = True
    
    FileNum = FreeFile
    
    Open filepath For Input As #FileNum
        Input #1, m_NumBuildings
        
        ReDim m_Buildings(m_NumBuildings - 1)
        For I = 0 To m_NumBuildings - 1
            With m_Buildings(I)
                Input #FileNum, .Name
                Input #FileNum, .Pos.x
                Input #FileNum, .Pos.y
                Input #FileNum, .houseWidth
                Input #FileNum, .houseHeight
                Input #FileNum, .Door.x
                Input #FileNum, .Door.y
                Input #FileNum, .roofSlopesHoriz
            End With
        Next
    Close #FileNum
    
    m_WallTextures(0).Init App.Path & "\Images\Tiles\Brick1.bmp", , , RGB(255, 0, 255)
    m_WallTextures(1).Init App.Path & "\Images\Tiles\Brick2.bmp", , , RGB(255, 0, 255)
    m_WallTextures(2).Init App.Path & "\Images\Tiles\Brick3.bmp", , , RGB(255, 0, 255)
    m_WallTextures(3).Init App.Path & "\Images\Tiles\Brick4.bmp", , , RGB(255, 0, 255)
End Sub

Public Sub CreateNPCs(filepath As String)
    Dim I As Integer
    
    'm_NumNPCs = numNPCs
    
    'ReDim m_NPCs(0 To m_NumNPCs - 1)
    
    'For I = 0 To numNPCs - 1
    m_NPCs.Setup filepath
    'Next
End Sub

Public Sub DrawMap(Optional ByVal PosX As Long = 0, Optional ByVal PosY As Long = 0, Optional HeroIsCentered As Boolean = True)
    Dim x As Long, y As Long
    Dim DrawPosX As Long, DrawPosY As Long
    Dim StartX As Long, EndX As Long, StartY As Long, EndY As Long
    Dim MapX As Long, MapY As Long
    Dim MinX As Long, MinY As Long
    Dim ClipX As Long, ClipY As Long
    Dim I As Integer
    
    If HeroIsCentered = False Then
        m_MapPos.x = PosX
        m_MapPos.y = PosY
    Else
        m_MapPos.x = ScreenWidth / 2 - (m_Player.Pos.x * HalfTileWidth - m_Player.Pos.y * HalfTileWidth) - m_Player.PosOnTile.x
        m_MapPos.y = ScreenHeight / 2 - (m_Player.Pos.y * HalfTileHeight + m_Player.Pos.x * HalfTileHeight) - m_Player.PosOnTile.y
    End If
    DrawPosX = m_MapPos.x 'start drawing the map at the given position
    DrawPosY = m_MapPos.y
    
    StartX = 0
    StartY = 0
    EndX = m_MapWidth - 1
    EndY = m_MapHeight - 1
    MapX = m_MapPos.x
    MapY = m_MapPos.y
    MinX = -TILE_WIDTH
    MinY = -TILE_HEIGHT
    
    For y = StartY To EndY
        For x = StartX To EndX
            DrawPosX = MapX + x * HalfTileWidth - y * HalfTileWidth
            DrawPosY = MapY + y * HalfTileHeight + x * HalfTileHeight
            
            If DrawingMap Then
                If DrawPosX > MinX And DrawPosX < ScreenWidth Then
                    If DrawPosY > MinY And DrawPosY < ScreenHeight Then
                        m_MapTextures(m_Map(x, y).TileID).Render DrawPosX, DrawPosY, TILE_WIDTH, TILE_HEIGHT, 0, 0, TILE_WIDTH, TILE_HEIGHT
                    End If
                End If
            End If
            
            If DrawingBuildings Then
                For I = 0 To m_NumBuildings - 1
                    With m_Buildings(I)
                        If x = .Pos.x And y >= .Pos.y And y <= .Pos.y + .houseHeight - 1 Then
                            m_WallTextures(2).Render DrawPosX, DrawPosY - WALL_HEIGHT + TILE_HEIGHT, TILE_WIDTH, WALL_HEIGHT, 0, 0, TILE_WIDTH, WALL_HEIGHT
                        End If
                        If y = .Pos.y And x >= .Pos.x And x <= .Pos.x + .houseWidth - 1 Then
                            m_WallTextures(3).Render DrawPosX, DrawPosY - WALL_HEIGHT + TILE_HEIGHT, TILE_WIDTH, WALL_HEIGHT, 0, 0, TILE_WIDTH, WALL_HEIGHT
                        End If
                    End With
                Next
            End If
                    
            'If DrawingNPCs Then
                'For I = 0 To m_NumNPCs - 1
                    If m_NPCs.PosX = x And m_NPCs.PosY = y Then
                        m_NPCs.Render (DrawPosX - (m_NPCs.m_Width / 2)) + m_NPCs.PosOnTileX, DrawPosY - m_NPCs.m_Height + m_NPCs.PosOnTileY
                    End If
                'Next
            'End If
            
            If DrawingPlayer Then
                If m_Player.Pos.x = x And m_Player.Pos.y = y Then
                    m_Player.Texture.Render (DrawPosX - HalfTileWidth) + m_Player.PosOnTile.x, (DrawPosY - m_Player.Height + m_Player.PosOnTile.y), m_Player.Width, m_Player.Height, 0, 0
                End If
            End If
            
            If DrawingBuildings Then
                For I = 0 To m_NumBuildings - 1
                    With m_Buildings(I)
                        If x = .Pos.x + .houseWidth - 1 And y >= .Pos.y And y <= .Pos.y + .houseHeight - 1 Then
                            m_WallTextures(1).Render DrawPosX, DrawPosY - WALL_HEIGHT + TILE_HEIGHT, TILE_WIDTH, WALL_HEIGHT, 0, 0, TILE_WIDTH, WALL_HEIGHT
                        End If
                        If y = .Pos.y + .houseHeight - 1 And x >= .Pos.x And x <= .Pos.x + .houseWidth - 1 Then
                            m_WallTextures(0).Render DrawPosX, DrawPosY - WALL_HEIGHT + TILE_HEIGHT, TILE_WIDTH, WALL_HEIGHT, 0, 0, TILE_WIDTH, WALL_HEIGHT
                        End If
                    End With
                Next
            End If
            
        Next
    Next
End Sub

Public Function GetTileWidth() As Long
    GetTileWidth = TILE_WIDTH
End Function

Public Function GetTileHeight() As Long
    GetTileHeight = TILE_HEIGHT
End Function

Public Function GetPlayerX() As Long
    GetPlayerX = m_Player.Pos.x
End Function

Public Function GetPlayerY() As Long
    GetPlayerY = m_Player.Pos.y
End Function

Public Function GetMapWidth() As Long
    GetMapWidth = m_MapWidth
End Function

Public Function GetMapHeight() As Long
    GetMapHeight = m_MapHeight
End Function

Private Sub SetWalkable()
    Dim x As Integer, y As Integer
    
    For y = 0 To m_MapHeight - 1
        For x = 0 To m_MapWidth - 1
            m_Map(x, y).Walkable = True
            Select Case m_Map(x, y).TileID
                Case TILE_WATER: m_Map(x, y).Walkable = False
            End Select
        Next
    Next
End Sub

Public Function CheckTileBoundaries(AmntX As Long, AmntY As Long) As Single
    'INFO: TopLeft lines equation is Y = -0.583X + 14
    '      TopRight lines equation is Y = 0.583X + -14.567
    '      BottomLeft lines equation is Y = 0.583X + 16.008
    '      BottomRight lines equation is Y = -0.583X + 44.567
    Dim lineY As Double
    
    
    CheckTileBoundaries = -1
    
    With m_Player
        'check if player is beyond topleft
        lineY = (-0.583) * (.PosOnTile.x + AmntX) + 14
        If .PosOnTile.y + AmntY < lineY And .PosOnTile.x + AmntX < HalfTileWidth Then
            CheckTileBoundaries = SIDE_TOPLEFT
        End If
        
        'check if player is beyond topright
        lineY = 0.583 * (.PosOnTile.x + AmntX) + -14.567
        If .PosOnTile.y + AmntY < lineY And .PosOnTile.x + AmntX > HalfTileWidth + 1 Then
            CheckTileBoundaries = SIDE_TOPRIGHT
        End If
        
        'check if player is beyond bottomleft
        lineY = 0.583 * (.PosOnTile.x + AmntX) + 16.008
        If .PosOnTile.y + AmntY > lineY And .PosOnTile.x + AmntX < HalfTileWidth Then
            CheckTileBoundaries = SIDE_BOTTOMLEFT
        End If
        
        'check if player is beyond bottomright
        lineY = (-0.583) * (.PosOnTile.x + AmntX) + 44.567
        If .PosOnTile.y + AmntY > lineY And .PosOnTile.x + AmntX > HalfTileWidth + 1 Then
            CheckTileBoundaries = SIDE_BOTTOMRIGHT
        End If
        
        'check if player is beyond top of tile
        If .PosOnTile.y + AmntY < 0 And .PosOnTile.x + AmntX >= HalfTileWidth And .PosOnTile.x + AmntX <= HalfTileWidth + 1 Then
            CheckTileBoundaries = SIDE_TOP
        End If
        
        'check if player is beyond bottom of tile
        If .PosOnTile.y + AmntY > TILE_HEIGHT - 1 And .PosOnTile.x + AmntX >= HalfTileWidth And .PosOnTile.x + AmntX <= HalfTileWidth + 1 Then
            CheckTileBoundaries = SIDE_BOTTOM
        End If
        
        'check if player is beyond left of tile
        If .PosOnTile.x + AmntX < 0 And .PosOnTile.y + AmntY >= HalfTileHeight And .PosOnTile.y + AmntY <= HalfTileHeight + 1 Then
            CheckTileBoundaries = SIDE_LEFT
        End If
        
        'check if player is beyond right of tile
        If .PosOnTile.x + AmntX > TILE_WIDTH - 1 And .PosOnTile.y + AmntY >= HalfTileHeight And .PosOnTile.y + AmntY <= HalfTileHeight + 1 Then
            CheckTileBoundaries = SIDE_RIGHT
        End If
        
    End With
End Function

Private Function CheckForBuildings(Direction As Integer) As Boolean
    Dim I As Integer
    
    CheckForBuildings = False
        
    For I = 0 To m_NumBuildings - 1
    
        'check topright (player's pos.y -1)
        If Direction = SIDE_TOPRIGHT Then
            'check if player's x is inside building x
            If m_Player.Pos.x >= m_Buildings(I).Pos.x And m_Player.Pos.x <= m_Buildings(I).Pos.x + m_Buildings(I).houseWidth - 1 Then
                'check if player's y is right below building.
                If m_Player.Pos.y = m_Buildings(I).Pos.y + m_Buildings(I).houseHeight Then
                    'the player is right beneath a building - return true
                    CheckForBuildings = True
                End If
            End If
        End If
        
        'check topleft (player's pos.x -1)
        If Direction = SIDE_TOPLEFT Then
            'check if player's y is inside building y
            If m_Player.Pos.y >= m_Buildings(I).Pos.y And m_Player.Pos.y <= m_Buildings(I).Pos.y + m_Buildings(I).houseHeight - 1 Then
                'check if player's x is right next to the building.
                If m_Player.Pos.x = m_Buildings(I).Pos.x + m_Buildings(I).houseWidth Then
                    'the player is right beside a building - return true
                    CheckForBuildings = True
                End If
            End If
        End If
        
        'check bottomright (player's pos.x +1)
        If Direction = SIDE_BOTTOMRIGHT Then
            'check if player's y is inside building y
            If m_Player.Pos.y >= m_Buildings(I).Pos.y And m_Player.Pos.y <= m_Buildings(I).Pos.y + m_Buildings(I).houseHeight - 1 Then
                'check if player's x is right next to the building.
                If m_Player.Pos.x = m_Buildings(I).Pos.x - 1 Then
                    'the player is right beside a building - return true
                    CheckForBuildings = True
                End If
            End If
        End If
        
        'check bottomleft (player's pos.y +1)
        If Direction = SIDE_BOTTOMLEFT Then
            'check if player's x is inside building x
            If m_Player.Pos.x >= m_Buildings(I).Pos.x And m_Player.Pos.x <= m_Buildings(I).Pos.x + m_Buildings(I).houseWidth - 1 Then
                'check if player's y is right above building.
                If m_Player.Pos.y = m_Buildings(I).Pos.y - 1 Then
                    'the player is right above a building - return true
                    CheckForBuildings = True
                End If
            End If
        End If

    Next
End Function


Public Sub DestroyTextures()
    Dim I As Integer
    
    For I = 0 To NUM_TILES - 1
        m_MapTextures(I).CleanUp
    Next
End Sub

Private Function CheckWalkability(TileX As Long, TileY As Long) As Boolean
    CheckWalkability = False
    
    If TileX >= 0 And TileX <= m_MapWidth - 1 Then
        If TileY >= 0 And TileY <= m_MapHeight - 1 Then
            If m_Map(TileX, TileY).Walkable = True Then
                CheckWalkability = True
            End If
        End If
    End If
End Function

