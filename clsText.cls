VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMainFont As D3DXFont
Private mMainFontDesc As IFont
Private mTextRect As RECT
Private mVertices(3) As TLVERTEX
Private mColor As Long

Sub Create(Optional font As String = "Arial", Optional fntSize As Long = 12, Optional bold As Boolean = False, Optional Color As Long = &HFFCCCCFF)
    On Error GoTo Handler
    Dim fnt As New StdFont
    
    With fnt
        .Name = font
        .Size = fntSize
        .bold = bold
    End With
    

    mColor = Color
    Set mMainFontDesc = fnt
    Set mMainFont = D3DX.CreateFont(D3DDevice, mMainFontDesc.hFont)

    
    Exit Sub
Handler:
    MsgBox "Text creation failed...", , "cText Error"
End Sub

Sub Render(ByVal Text As String, X As Long, Y As Long, Optional Width As Long, Optional Height As Long)
    If X > ScreenWidth Or Y > ScreenHeight Then Exit Sub
    
    If Width = 0 Then Width = ScreenWidth - X
    If Height = 0 Then Height = ScreenHeight - Y
    
    With mTextRect
        .Top = Y
        .Left = X
        .bottom = Y + Height
        .Right = X + Width
    End With
    
    D3DX.DrawText mMainFont, mColor, Text, mTextRect, DT_TOP Or DT_LEFT
End Sub
