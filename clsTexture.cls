VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright 2004 Calvin Mayer
Option Explicit

Private Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

Private m_MaxUnusedTime As Long
Private m_TimeSinceRender As Long
Private m_TransColor As Long
Private m_FilePath As String
Private m_Width As Long
Private m_Height As Long
Private m_Texture As Direct3DTexture8
Private m_Vertices(0 To 3) As TLVERTEX

Sub Init(filePath As String, Optional Width As Long = 0, Optional Height As Long = 0, Optional transColor As Long = 0)
    m_FilePath = filePath
    m_Width = Width
    m_Height = Height
    m_TransColor = transColor
    m_TimeSinceRender = GetTickCount()
End Sub

Private Sub LoadTexture()

    Set m_Texture = D3DX.CreateTextureFromFileEx(D3DDevice, m_FilePath, m_Width, _
    m_Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, _
    D3DX_FILTER_POINT, D3DX_FILTER_POINT, m_TransColor, ByVal 0, ByVal 0)
    
    If m_Texture Is Nothing Then MsgBox "Error loading texture.", , "cTexture Error"
    
End Sub

Sub Render(ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal SrcX As Single, ByVal SrcY As Single, Optional ByVal SrcWidth As Long = -1, Optional ByVal SrcHeight As Single = -1, Optional angle As Single = 0, Optional alphaValue As Integer = 255)
    Dim CosQ As Single
    Dim SinQ As Single
    Dim TempX As Single, TempY As Single
    Dim XCenter As Single, YCenter As Single
    
    If m_Texture Is Nothing Then LoadTexture
    
    If SrcWidth = -1 Then SrcWidth = Width
    If SrcHeight = -1 Then SrcHeight = Height
        
    XCenter = X + Width / 2
    YCenter = Y + Height / 2
        
    SinQ = Sin(angle * Deg2Rad)
    CosQ = Cos(angle * Deg2Rad)

    m_Vertices(0) = CreateTLVertex(X, Y, 1, 1, D3DColorARGB(alphaValue, 255, 255, 255), 0, SrcX / Width, SrcY / Height)
    m_Vertices(1) = CreateTLVertex(X + Width, Y, 1, 1, D3DColorARGB(alphaValue, 255, 255, 255), 0, (SrcX + SrcWidth) / Width, SrcY / Height)
    m_Vertices(2) = CreateTLVertex(X, Y + Height, 1, 1, D3DColorARGB(alphaValue, 255, 255, 255), 0, SrcX / Width, (SrcY + SrcHeight) / Height)
    m_Vertices(3) = CreateTLVertex(X + Width, Y + Height, 1, 1, D3DColorARGB(alphaValue, 255, 255, 255), 0, (SrcX + SrcWidth) / Width, (SrcY + SrcHeight) / Height)
     
    Dim I As Integer
    For I = 0 To 3
        With m_Vertices(I)
            TempX = ((.X - XCenter) * CosQ) - ((.Y - YCenter) * SinQ)
            TempY = ((.X - XCenter) * SinQ) + ((.Y - YCenter) * CosQ)
            .X = TempX + XCenter
            .Y = TempY + YCenter
        End With
    Next
          
    D3DDevice.SetTexture 0, m_Texture
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, m_Vertices(0), Len(m_Vertices(0))


    m_TimeSinceRender = GetTickCount()
End Sub

Private Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, Color As Long, specular As Long, tu As Single, tv As Single) As TLVERTEX
    With CreateTLVertex
        .X = X
        .Y = Y
        .Z = Z
        .rhw = rhw
        .Color = Color
        .specular = specular
        .tu = tu
        .tv = tv
    End With
End Function

Private Sub Class_Initialize()
    m_MaxUnusedTime = 2000
End Sub

Public Function GetWidth() As Long
    GetWidth = m_Width
End Function

Public Function GetHeight() As Long
    GetHeight = m_Height
End Function

Sub CheckTime()
    If GetTickCount() - m_TimeSinceRender > m_MaxUnusedTime Then
        Set m_Texture = Nothing
    End If
End Sub

Function debugThing() As String
    If m_Texture Is Nothing Then
        debugThing = "cTexture is not loaded!"
    Else
        debugThing = "cTexture is loaded!"
    End If
    
End Function
Sub CleanUp()
    Set m_Texture = Nothing
End Sub


