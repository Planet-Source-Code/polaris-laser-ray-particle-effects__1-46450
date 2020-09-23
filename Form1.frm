VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "NemoX Tut 14 - Adding great particle Effects"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'========================================
' -NemoX Tut 14 - Adding great particle Effects
'========================================


'============================================
'
' YOU HAVE TO SETUP THE LATEST NEMOX ENGINE
' TO MAKE THIS TURORIAL RUN
'
' GET IT AT
' http://perso.wanadoo.fr/malakoff/NemoXsetup1.074.exe
'
'===========Prerequisities

Option Explicit


'The Main One is The NemoX renderer
Dim Nemo As NemoX



'we're gonna use a Nemo class for rendering a mesh or polygons
'we use cNemo_Mesh class

Dim WithEvents Mesh As cNemo_Mesh
Attribute Mesh.VB_VarHelpID = -1

'we will need a quick acces to very important and useful functions
Dim Tool As cNemo_Tools

'========NEW OBJECTS=======

Dim SHOCK1 As cNemo_ParticleEngine

Dim SHOCK2 As cNemo_ParticleEngine

Dim FX_2 As cNemo_ParticleEngine


Dim FIRE As cNemo_ParticleEngine

Dim SNOW As cNemo_ParticleEngine





'New in NemoX 1.74
'this will handle camera moves
'the new camera class support Mouse rotation
' and Freelook (6DOF style)
Dim KAMERA As cNemo_Camera2

'this will handle for keyboard input state
Dim KEY As cNemo_Input

'some constant for player speed
'Check at Sub GetKey()
Private Const RotationSpeed = 1 / 500

Private Const MoveSpeed = 2

'Entry point for our project
Private Sub Form_Load()

    Me.Show
    Me.Refresh

    'call the initializer sub
    Call InitEngine

  
    'make geometry
    Call BuilGeometry

    'call the main game Loop
    Call gameLoop

End Sub

'we build our mesh here

Sub BuilGeometry()

  Dim I As Long

    'adding a plane surface for a simple floor$

    'first off very important we pass a texture to the meshbuilder
    Mesh.Add_Texture (App.Path + "\ground.jpg")          '0

    'Add flooor surface here
    Mesh.Add_WallFloor Tool.Vector(-5000, -1, -5000), Tool.Vector(5000, -1, 5000), 10, 10, 0

    'Just add some details at scene

    'Feel free to add more geometry details
    For I = 1 To 5
        Mesh.Add_Cilynder Tool.Vector(450 - I * 450 - 50, 10, -815 + 1500), 50, 490, 8, 0
    
        Mesh.Add_Cilynder Tool.Vector(450 - I * 450 - 50, 10, -1035 + 1500), 50, 490, 8, 0

    Next I

    '========IMPORTANT========
    'then we build our mesh
    Mesh.BuilDMesh
    
    
    
    'Prepare our particle system here
    
    SHOCK1.InitParticles Tool.Vector(0, 15, 10), App.Path + "\particle.BMP", 80, NEMO_EXPLOSION_SHOCKWAVE_FX
    SHOCK1.Set_ShockWave SHOCHWAVE_X
    SHOCK1.Set_ParticleSize 50, 50
    SHOCK1.Set_ParticleLifeTime 5000
    
    SHOCK2.InitParticles Tool.Vector(0, 15, 10), App.Path + "\particle.BMP", 100, NEMO_EXPLOSION_SHOCKWAVE_FX
    SHOCK2.Set_ParticleSize 50, 50
    SHOCK2.Set_ParticleLifeTime 5000
    SHOCK2.Set_ShockWave SHOCHWAVE_Z

    

    FX_2.InitParticles Tool.Vector(0, 15, 10), App.Path + "\particle.BMP", 100, NEMO_FONTAIN_FX
    FX_2.Set_ParticleLifeTime 5000
    FX_2.Set_ParticleSize 50, 50

    
    
    SNOW.InitParticles Tool.Vector(0, 15, 0), App.Path + "\snow.bmp", 200, NEMO_SNOW_FX
    SNOW.Set_Snow 800, 800, 800, 1
    SNOW.Set_ParticleSize 10, 10
    
    
    
    'prepare Fire Effect
    
    FIRE.InitParticles Tool.Vector(0, 15, 10), App.Path + "\FIRE.bmp", 80, NEMO_FIRE_FX
    'customize fire Blending effect here from NEMO_FIRE_FX1 to NEMO_FIRE_FX21
    FIRE.Set_Fire 60, , 0.5, NEMO_FIRE_FX14
    FIRE.Set_ParticleSize 10, 10
    FIRE.Set_ParticleDrawingMode True
    
    
   
    

End Sub



'we will used that sub for the engine initialization

Sub InitEngine()

  'first thing allocate memory for the main Object

    Set Nemo = New NemoX

    Set Tool = New cNemo_Tools
    
  

    'allocate memory for our meshbuilder
    Set Mesh = New cNemo_Mesh

    '====New code======
    
    
      Set SHOCK2 = New cNemo_ParticleEngine
    Set FX_2 = New cNemo_ParticleEngine
    
    Set FIRE = New cNemo_ParticleEngine

    Set SNOW = New cNemo_ParticleEngine

    Set SHOCK1 = New cNemo_ParticleEngine
    
    
     ' NEW CAMERA CLASS FASTER then cNemo_Camera and support
     ' Quaternion rotation and 6DOF
    Set KAMERA = New cNemo_Camera2
    
    

    '.......MEMORY ALLOCATION....
    Set KEY = New cNemo_Input

    'we use this method
    'now we allow the user to choose options
    '32/16 bit backbuffer

    'for this demo Windowed mode is recommanded
    If Not (Nemo.INIT_ShowDeviceDLG(Form1.hWnd)) Then
        End 'terminate here if error
    End If


    'Nemo.Initialize Me.hWnd
    


   

    

    'set the back clearcolor
    Nemo.BackBuffer_ClearCOLOR = RGB(80, 80, 80) 'Gray

    'set some parameters 'near far  FOVangle,Aspect
    KAMERA.Set_ViewFrustum 10, 5000, 3.14 / 4, 1.01
    
    'set our camera
    
    KAMERA.Set_Position Tool.Vector(0, 50, -50)  'Starting Position
    KAMERA.Set_LookAt Tool.Vector(0, 50, 0)  'LooK

    'Activate Free_Look 6DOF CAMERA
    'Note That default is FPS (QUAKE-LIKE)
    KAMERA.Set_CameraStyle FREE_6DOF
End Sub




'this sub is the main loop for a game or 3d apllication
Sub gameLoop()

  'loop untill player press 'ESCAPE'
  'Nemo.Set_CullMode D3DCULL_NONE

    Nemo.Set_EngineRenderState D3DRS_ZENABLE, 1
    Nemo.Set_light 1
    
    
    Do

        '=====Keyboard handler can be added here
        Call GetKey
        DoEvents

       
        'start the 3d renderer
        Nemo.Begin3D
        '===============ADD game rendering mrthod here

        'draw our ground here
        Mesh.Render
        
        
        'render Particles
        
        SHOCK2.Render


        'animate this one
        FX_2.Set_EmiterPosition 50 + Sin(Timer * 2) * 150, 45, Cos(Timer * 2) * 150
        FX_2.Render

        SNOW.Render
        
        FIRE.Render
        
        SHOCK1.Render

       
       
        'show the FPS at pixel(5,10) color White
        Nemo.Draw_Text "FPS:" + Str(Nemo.Framesperseconde), 5, 10, &HFFFFFFFF
       
        Nemo.End3D
        'end the 3d renderer

        'check the player keyPressed
    Loop Until KEY.Get_KeyBoardKeyPressed(NEMO_KEY_ESCAPE)

    Call EndGame

End Sub

'----------------------------------------
'Name: GetKey
'----------------------------------------
Sub GetKey()

  
    'just move Forward
    If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_UP) Then _
       KAMERA.Move_Forward 1 * MoveSpeed
    
     'fast Forward
    If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_RCONTROL) Then _
       KAMERA.Move_Forward 4 * MoveSpeed
    
    'just move BackWard
    If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_DOWN) Then _
       KAMERA.Move_Backward 1 * MoveSpeed



  'Rotate left
    If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_LEFT) Then _
       KAMERA.Turn_Left 4 * RotationSpeed
    
    'Rotate right
    If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_RIGHT) Then _
       KAMERA.Turn_Right 4 * RotationSpeed

    'to take a snapshot
    If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_S) Then _
       Nemo.Take_SnapShot App.Path + "\Shot.bmp"

   'use mouse rotation
   KAMERA.RotateByMouse , , 0
   
   KAMERA.Update
End Sub



Sub EndGame()

  'end of the demo

    Set KEY = Nothing
    Set KAMERA = Nothing
    Set Mesh = Nothing
   
    Nemo.Free  'free resources used by the engine
    Set Nemo = Nothing
    End

End Sub

Private Sub Form_Unload(Cancel As Integer)

    EndGame

End Sub


