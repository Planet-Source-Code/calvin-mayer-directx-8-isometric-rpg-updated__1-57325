Attribute VB_Name = "modStuff"
Option Explicit

Public Const NUM_TILES = 16
Public Const TILE_GRASS = 0
Public Const TILE_DIRT = 1
Public Const TILE_WATER = 2
Public Const TILE_COBBLESTONE = 3
Public Const TILE_WOOD = 4
Public Const TILE_SNOW = 5
Public Const TILE_SNOWTOPGRASSBOTTOM = 6
Public Const TILE_SNOWBOTTOMGRASSTOP = 7
Public Const TILE_SNOWLEFTGRASSRIGHT = 8
Public Const TILE_SNOWRIGHTGRASSLEFT = 9
Public Const TILE_SNOWTOPDIRTBOTTOM = 10
Public Const TILE_SNOWBOTTOMDIRTTOP = 11
Public Const TILE_SNOWLEFTDIRTRIGHT = 12
Public Const TILE_SNOWRIGHTDIRTLEFT = 13
Public Const TILE_REDDIRT = 14
Public Const TILE_DEADGRASS = 15

Public Const NUM_WALLS = 4
Public Const WALL_BRICK1 = 0
Public Const WALL_BRICK2 = 1
Public Const WALL_BRICK3 = 2
Public Const WALL_BRICK4 = 3

Public Const TILE_WIDTH = 20 'width of the tile image
Public Const TILE_HEIGHT = 20 'height of the tile image

Public HalfTileWidth As Single
Public HalfTileHeight As Single

Type tileOBJ
    Pic As StdPicture
    hdc As Long
End Type

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

