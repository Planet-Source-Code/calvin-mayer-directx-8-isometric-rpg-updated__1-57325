Attribute VB_Name = "modDirectXEngine"
'Copyright 2004 Calvin Mayer

Option Explicit

Public DX As DirectX8 'root object
Public DI As DirectInput8
Public D3D As Direct3D8 'direct3d interface

Public D3DX As D3DX8 'helper library

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal Length As Long)

Public D3DDevice As Direct3DDevice8 'represents the hardware doing the rendering

'//Important, only set one of the following 2 to be true.
Private Const UsePollingMethod As Boolean = True
Private Const UseEventMethod As Boolean = Not UsePollingMethod

'//These next two objects are used to access our device (keyboard)
Private DIDevice As DirectInputDevice8
Private DIState As DIKEYBOARDSTATE
Private KeyState(0 To 255) As Boolean 'so we can detect if the key has gone up or down!
Private Const BufferSize As Long = 10 'how many events the buffer holds.
'This can be 1 if using event based, but 10-20 if polling based...

'Sleep() - stops our polling loop going too fast ;)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type Point
    x As Long
    y As Long
End Type

Public Type tRGBA
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Public Type TLVERTEX
    x As Single
    y As Single
    Z As Single
    rhw As Single
    Color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

Public Const Pi = 3.141592653589
Public Const Deg2Rad = Pi / 180#
Public Const Rad2Deg = 180# / Pi

Public bRunning As Boolean 'controls whether the program is running or not

Public ScreenWidth As Integer
Public ScreenHeight As Integer

'Flexible-Vertex-Format description for a 2d vertex
Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Private Type TextInfo
    MainFont As D3DXFont
'    MainFontDesc As IFont
'    TextRect As RECT
'    vertChar(3) As TLVERTEX
'   Text As String
'    Color As Long
End Type

Public FPS_LastCheck As Long
Public FPS_Count As Long
Public FPS_Current As Long

Dim fpsText As TextInfo

Sub UpdateFps()
    If FPS_LastCheck = 0 Then FPS_LastCheck = GetTickCount - 1000
    If GetTickCount - FPS_LastCheck >= 1000 Then
        FPS_Current = FPS_Count
        FPS_Count = 0
        FPS_LastCheck = GetTickCount
    Else
        FPS_Count = FPS_Count + 1
    End If
End Sub

Function InitDirectX(Width As Long, Height As Long, TheForm As Form, Optional Windowed As Boolean = False, Optional BitsPerPixel As Long = 32, Optional AntiAlias As Boolean) As Boolean
    On Error GoTo Handler
    
    Dim DispMode As D3DDISPLAYMODE 'describes our display mode
    Dim D3DWindow As D3DPRESENT_PARAMETERS 'Describes our ViewPort
    
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    If InitDirectInput = False Then GoTo Handler
    
    'Fullscreen mode stuff here
    Select Case BitsPerPixel
        Case 32
            DispMode.Format = D3DFMT_A8R8G8B8
        Case 24
            DispMode.Format = D3DFMT_R8G8B8
        Case 16
            DispMode.Format = D3DFMT_R5G6B5
        Case 4
            DispMode.Format = D3DFMT_D16
    End Select
    
    If Windowed = False Then
        DispMode.Width = Width
        DispMode.Height = Height
    End If
    
    ScreenWidth = Width
    ScreenHeight = Height
    
    'SetScaleDimentions theForm, width, height
    
    With D3DWindow
        .SwapEffect = D3DSWAPEFFECT_DISCARD
        .BackBufferCount = 1
        .BackBufferFormat = DispMode.Format
        .BackBufferHeight = Height
        .BackBufferWidth = Width
        .hDeviceWindow = TheForm.hWnd
        If AntiAlias Then .MultiSampleType = D3DMULTISAMPLE_4_SAMPLES
        If Windowed Then .Windowed = 1
    End With
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    'create the device
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, TheForm.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
    
    With D3DDevice
        'Set the vertex shader to use our vertex format
        .SetVertexShader FVF
        
        'transformed and lit vertices dont need lighting so we disable it:
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_ALPHABLENDENABLE, 1
        
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        

        If AntiAlias Then
            .SetRenderState D3DRS_MULTISAMPLE_ANTIALIAS, True
            .SetRenderState D3DRS_EDGEANTIALIAS, True
        End If
    End With
    'Set Char.ourTexture = D3DX.CreateTextureFromFileEx(D3DDevice, _
    App.Path & "\ExampleTexture.bmp", 0, 0, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
    D3DPOOL_DEFAULT, D3DX_FILTER_POINT, D3DX_FILTER_POINT, vbWhite, ByVal 0, ByVal 0)
    
    'fpsText = CreateText("FPS: ", "Terminal", , True)
    
    InitDirectX = True
    Exit Function
    
Handler:
    MsgBox "An error occurred while initializing DirectX!", , "ERROR"
    Debug.Print "Error Number Returned: " & Err.Number
    InitDirectX = False
End Function

Sub CleanUpDirectX()
    On Error Resume Next
    
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set DX = Nothing
    DIDevice.Unacquire
    Set DIDevice = Nothing
    Set DI = Nothing
End Sub

Sub ClearDevice(Optional Color As Long = vbBlack)
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, Color, 1#, 0
End Sub

Sub PresentDevice()
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Function InitDirectInput() As Boolean
    On Error GoTo Handler
    '//0. Any variables
     Dim I As Long
     Dim DevProp As DIPROPLONG
     Dim DevInfo As DirectInputDeviceInstance8
     Dim pBuffer(0 To BufferSize) As DIDEVICEOBJECTDATA
                                                             
    '//1. Check options.
     If UsePollingMethod And UseEventMethod Then
          MsgBox "Error, UsePollingMethod and UseEventMethod are both set to true!"
          GoTo Handler
     End If

     'If UsePollingMethod Then txtOutput.Text = "Using Polling Method" & vbCrLf
     'If UseEventMethod Then txtOutput.Text = "Using Event Based Method" & vbCrLf
    
     Set DI = DX.DirectInputCreate
     Set DIDevice = DI.CreateDevice("GUID_SysKeyboard") 'the string is important, not just a random string...
                                                             
     DIDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
     DIDevice.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
                                                             
     'set up the buffer...
     DevProp.lHow = DIPH_DEVICE
     DevProp.lData = BufferSize
     DIDevice.SetProperty DIPROP_BUFFERSIZE, DevProp
                                                             
     DIDevice.Acquire 'let DirectX know that we want to use the device now.
     
     InitDirectInput = True
     Exit Function
Handler:
    MsgBox "Direct Input Initialization failed...", , "DirectX Error"
     InitDirectInput = False
End Function

Public Sub BeginScene()
    D3DDevice.BeginScene
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
End Sub

Public Sub EndScene()
    D3DDevice.EndScene
End Sub
