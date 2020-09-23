VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'main engine
Dim Nemo As New NemoX


'the house mesh
Dim HOUSE As New cNemo_Mesh






'----------------------------------------
'Name: GameLoop
'----------------------------------------
Sub GameLoop()
Dim VC As D3DVECTOR
'position camera
Nemo.Camera_SetPosition Vector(120#, 8#, -440#), _
                                Vector(120#, 8#, 441#)
                                
'set camera rotation
Nemo.Camera_SetRotation 0, 0, 0


'set bilinear texture filtering
Nemo.Set_EngineTextureFilter NEMO_FILTER_BILINEAR
 

'set the view 45°
Nemo.Set_ViewFrustum 10, 5500, 45 * RAD
     

'main loop
Do
 
  
  DoEvents
  Call Me.GetKey
  If Nemo.Get_KeyPress(DIK_ESCAPE) Then GoTo End_it
 
  'calculate frustum
  GLOB.SetUpFrustum
  
    
    If Nemo.GetD3dDevice Is Nothing Then Exit Sub


   
    Nemo.Clear3D
    
    Nemo.GetD3dDevice.BeginScene
    
     'check for collision
      If HOUSE.CheckCollisionSliding(Nemo.Camera_GetPosition, VC, 15) Then Nemo.Camera_set_EYE VC 'Nemo.Camera_Recall
    'render our mesh
    HOUSE.Render
     
     
     'update FPS
     Nemo.Draw_Text "FPS=" + Str(Nemo.Framesperseconde), 10, 1
  
     Nemo.GetD3dDevice.EndScene
    
     
    Nemo.Flip
 
  
Loop

End_it:
Form_Unload 0
End Sub








'----------------------------------------
'Name: GetKey
'----------------------------------------
Sub GetKey()

'Nemo.Camera_SetPositionEX vC

'If XF.Get_Colision(vC) Then Nemo.Camera_Recall


If Nemo.Get_KeyPress(NEMO_KEY_LEFT) Then _
    Nemo.Camera_Turn_Left 1.5 / 50
If Nemo.Get_KeyPress(NEMO_KEY_RIGHT) Then _
    Nemo.Camera_Turn_Right 1.5 / 50
    
    
    If Nemo.Get_KeyPress(NEMO_KEY_UP) Then _
    Nemo.Camera_Move_Foward 5
    If Nemo.Get_KeyPress(NEMO_KEY_RCONTROL) Then _
    Nemo.Camera_Move_Foward 10
   If Nemo.Get_KeyPress(NEMO_KEY_DOWN) Then _
    Nemo.Camera_Move_Backward 5



  If Nemo.Get_KeyPress(NEMO_KEY_ADD) Then _
    Nemo.Camera_Strafe_UP 1
If Nemo.Get_KeyPress(NEMO_KEY_SUBTRACT) Then _
    Nemo.Camera_Strafe_DOWN 1


 If Nemo.Get_KeyPress(NEMO_KEY_NUMPAD8) Then
    Nemo.Camera_Turn_UP 1 / 50
End If
If Nemo.Get_KeyPress(NEMO_KEY_NUMPAD2) Then _
    Nemo.Camera_Turn_DOWN 1 / 50
    
If Nemo.Get_KeyPress(NEMO_KEY_SPACE) Then _
 Nemo.Camera_SetPosition Vector(0#, 8#, -8#), _
                                Vector(0#, 8#, 500#)

If Nemo.Get_KeyPress(NEMO_KEY_S) Then _
 Nemo.Take_SnapShot App.Path + "\Shot.bmp"




End Sub



'----------------------------------------
'Name: Form_Load
'Object: Form
'Event: Load
'----------------------------------------
Private Sub Form_Load()
 Me.Show
  
  
  DoEvents ' Let the PC do what it has to do
  
  'InitD3D ' Initialize Direct3D
  Nemo.INIT_ShowDeviceDLG Me.hwnd
  'If Not Nemo.Initialize(Me.hwnd) Then
    'Form_Unload 0
  'End If
  
  
  Geo
End Sub



'----------------------------------------
'Name: Geo
'----------------------------------------
Sub Geo()

 

HOUSE.LoadMesh App.Path + "\MESH.nmsh"

 GameLoop
End Sub


'----------------------------------------
'Name: Form_Unload
'Object: Form
'Event: Unload
'----------------------------------------
Private Sub Form_Unload(Cancel As Integer)
 Nemo.Free
 
 End
End Sub








