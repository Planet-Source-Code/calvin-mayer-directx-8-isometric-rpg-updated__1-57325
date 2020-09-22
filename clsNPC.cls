VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TILE_WIDTH = 50
Private Const TILE_HEIGHT = 30

Private Const HalfTileWidth = TILE_WIDTH / 2
Private Const HalfTileHeight = TILE_HEIGHT / 2

Private m_Name As String 'name of NPC
Private m_Pos As Point 'Tile the NPC is on
Private m_PosOnTile As Point 'position on the tile the NPC is on
Private m_Texture As New clsTexture 'NPC's texture
Private m_TexturePath As String ' location of teh pic u want to use from the app's directory
Private m_Speed As Single 'number of pixels it moves
Public m_Width As Integer 'width of NPC's pic
Public m_Height As Integer 'height of npc's pic
Private m_NpcZone As RECT 'the square in which the npc will remain so they dont wander off
Private m_Text As String 'the text the NPC will say
Private bShopkeeper As Boolean 'is the npc a shopkeeper?
Private store As New clsStore

Public Property Get PosX()
    PosX = m_Pos.x
End Property

Public Property Get PosY()
    PosY = m_Pos.y
End Property

Public Property Get PosOnTileX()
    PosOnTileX = m_PosOnTile.x
End Property

Public Property Get PosOnTileY()
    PosOnTileY = m_PosOnTile.y
End Property

Public Sub Setup(filepath As String)
    Dim I As Integer
    Dim FileNum As Integer
    Dim itemName As String
    
    Dim texturePath As String
    texturePath = App.Path & "\Images\Characters\NPC2.bmp"
    
    FileNum = FreeFile
    
    m_TexturePath = texturePath
    
    m_Texture.Init texturePath, , , RGB(255, 0, 255)
    
    Open filepath For Input As #FileNum
        Input #FileNum, m_Name
        Input #FileNum, m_Pos.x
        Input #FileNum, m_Pos.y
        Input #FileNum, m_PosOnTile.x
        Input #FileNum, m_PosOnTile.y
        Input #FileNum, m_Speed
        Input #FileNum, m_Width
        Input #FileNum, m_Height
        Input #FileNum, m_NpcZone.Top
        Input #FileNum, m_NpcZone.bottom
        Input #FileNum, m_NpcZone.Left
        Input #FileNum, m_NpcZone.Right
        Input #FileNum, m_Text
        Input #FileNum, bShopkeeper

    Close #FileNum
    
    If m_PosOnTile.x = -1 Then
        m_PosOnTile.x = HalfTileWidth
    End If
    
    If m_PosOnTile.y = -1 Then
        m_PosOnTile.y = HalfTileHeight
    End If
    
End Sub

Public Sub Render(x As Single, y As Single)
    m_Texture.Render x, y, m_Width, m_Height, 0, 0, m_Width, m_Height
End Sub



