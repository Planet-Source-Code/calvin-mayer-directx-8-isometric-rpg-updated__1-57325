VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Map"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   3720
      TabIndex        =   8
      Text            =   "10"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   3720
      TabIndex        =   7
      Text            =   "10"
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Map"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame fraBase 
      Caption         =   "Base Tile"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton optBaseTile 
         Caption         =   "Dead Grass"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optBaseTile 
         Caption         =   "Red Dirt"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optBaseTile 
         Caption         =   "Snow"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optBaseTile 
         Caption         =   "Water"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton optBaseTile 
         Caption         =   "Dirt"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optBaseTile 
         Caption         =   "Grass"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Label lblHeight 
      Caption         =   "Map Height:"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblWidth 
      Caption         =   "Map Width:"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreate_Click()
    Dim baseTile As Integer
    
    frmMain.MapWidth = txtWidth.Text
    frmMain.MapHeight = txtHeight.Text
    
    frmMain.RedimMap
    
    If optBaseTile(0).Value Then
        baseTile = TILE_GRASS
    ElseIf optBaseTile(1).Value Then
        baseTile = TILE_DIRT
    ElseIf optBaseTile(2).Value Then
        baseTile = TILE_WATER
    ElseIf optBaseTile(3).Value Then
        baseTile = TILE_SNOW
    ElseIf optBaseTile(4).Value Then
        baseTile = TILE_REDDIRT
    ElseIf optBaseTile(5).Value Then
        baseTile = TILE_DEADGRASS
    Else
        baseTile = TILE_GRASS
    End If
    
    frmMain.SetTiles baseTile
    
    frmNew.Hide
    
    Unload frmNew
    
End Sub

