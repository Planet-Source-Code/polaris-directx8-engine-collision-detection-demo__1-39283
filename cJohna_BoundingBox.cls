VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNemo_BoundingBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'   ============================================================
'    ----------------------------------------------------------
'     Application Name:
'     Developer/Programmer: Johna
'    ----------------------------------------------------------
'     Module Name: cNemo_BoundingBox
'     Module File: cJohna_BoundingBox.cls
'     Module Type: Class
'     Module Description:
'    ----------------------------------------------------------
'     © johna 2002
'    ----------------------------------------------------------
'   ============================================================



'==================================================
'  1/02/2002
'
'  CNEMO_Bounding box class
'    for handle with AABB colision
'
'
'-3/02/2002
'  give up builtin transformation tools
'   rotate euler ect...
'
'
'
'  Now update for being used with directx8
'  22-03-2002
'=========================================

Option Explicit

'for translating and scaling
Private Type VerTEX_PARAM
  vPosition As D3DVECTOR
  Vscal As D3DVECTOR
  Vrotate As D3DVECTOR
  Vsize As D3DVECTOR
  vCenter As D3DVECTOR
  vOrg As D3DVECTOR
  radius As Single
  OrgBOX As D3DRMBOX
  OrgSize As D3DVECTOR
  TransVert() As D3DVERTEX
  StaticVert() As D3DVERTEX
  VC(7) As D3DVECTOR
  Name As String
End Type

Private OBJET As VerTEX_PARAM

'for colision Detection
Private Const X_PI = 3.13987644547878
Private Const X_RAD = X_PI / 180

Private Const EPSILON = 0.0001
Private Type IntersectTest
  T As Double
  U As Double
  V As Double
End Type


Public Enum NEMO_COLISION_RET
    NORTH_COL = &H0
    SOUTH_COL = &H1
    EAST_COL = &H2
    WEST_COL = &H3
    TOP_COL = &H4
    DOWN_COL = &H5
    NONE = &HFFCCDD

End Enum


Private BBOX As D3DRMBOX
'Private Bvert() As D3DVERTEX 'for the CubeBOX

Private BoxMatrix As D3DMATRIX




'----------------------------------------
'Name: Get Name
'----------------------------------------
'----------------------------------------
'Name: Get Name
'Description:
'----------------------------------------
Property Get Name() As String
  Name = OBJET.Name
End Property


'----------------------------------------
'Name: CreateBox
'----------------------------------------
'----------------------------------------
'Name: CreateBox
'Description:
'----------------------------------------
Sub CreateBox(vBoxMin As D3DVECTOR, vBoxMax As D3DVECTOR, Optional Name As String)
  BBOX.min = vBoxMin
  BBOX.max = vBoxMax
  OBJET.Vscal = Vector(1, 1, 1)
  OBJET.OrgBOX = BBOX
  OBJET.OrgSize = VICTOR(BBOX.max.x - BBOX.min.x, BBOX.max.y - BBOX.min.y, BBOX.max.z - BBOX.min.z)
  OBJET.Name = Name
  Compute
End Sub




'----------------------------------------
'Name: CreateBoxEX
'----------------------------------------
'----------------------------------------
'Name: CreateBoxEX
'Description:
'----------------------------------------
Sub CreateBoxEX(vBOX As NemoBoundBOX, Optional Name As String)
  BBOX.min = Vector(vBOX.Vmin.x, vBOX.Vmin.y, vBOX.Vmin.z)
  BBOX.max = Vector(vBOX.Vmax.x, vBOX.Vmax.y, vBOX.Vmax.z)
  
  
  OBJET.Vscal = Vector(1, 1, 1)
  OBJET.OrgBOX = BBOX
  OBJET.OrgSize = VICTOR(BBOX.max.x - BBOX.min.x, BBOX.max.y - BBOX.min.y, BBOX.max.z - BBOX.min.z)
  OBJET.Name = Name
  Compute
End Sub


'----------------------------------------
'Name: CreateBoxFromVertices
'----------------------------------------
'----------------------------------------
'Name: CreateBoxFromVertices
'Description:
'----------------------------------------
Sub CreateBoxFromVertices(lpVERT() As D3DVERTEX, Optional Name As String)
Dim RET As D3DRMBOX
    OBJECT_getBoundingBOX lpVERT(), RET
    CreateBox RET.min, RET.max
End Sub


'Sub CombineWithEX(BoxtoCombine As D3DRMBOX)

'If (BoxtoCombine.max.X > maxx) Then bBOX.max.X = BoxtoCombine.max.X
        'If (BoxtoCombine.max.Y > maxy) Then bBOX.max.Y = BoxtoCombine.max.Y
        'If (BoxtoCombine.max.Z > MaxZ) Then bBOX.max.Z = BoxtoCombine.max.Z
        'If (bBOX.minx < minx) Then bBOX.min.X = bBOX.minx
        'If (bBOX.MinY < MinY) Then bBOX.min.Y = bBOX.MinY
        'If (bBOX.MinZ < MinZ) Then bBOX.min.Z = bBOX.MinZ
'End Sub




'----------------------------------------
'Name: CombineWith
'----------------------------------------
'----------------------------------------
'Name: CombineWith
'Description:
'----------------------------------------
Sub CombineWith(Vmin As D3DVECTOR, Vmax As D3DVECTOR)
Dim BoxtoCombine As D3DRMBOX

BoxtoCombine.min = Vmin
BoxtoCombine.max = Vmax


If (BoxtoCombine.max.x > BBOX.max.x) Then BBOX.max.x = BoxtoCombine.max.x
        If (BoxtoCombine.max.y > BBOX.max.y) Then BBOX.max.y = BoxtoCombine.max.y
        If (BoxtoCombine.max.z > BBOX.max.z) Then BBOX.max.z = BoxtoCombine.max.z
        If (BBOX.min.x < BoxtoCombine.min.x) Then BBOX.min.x = BBOX.min.x
        If (BBOX.min.y < BoxtoCombine.min.y) Then BBOX.min.y = BBOX.min.y
        If (BBOX.min.z < BoxtoCombine.min.z) Then BBOX.min.z = BBOX.min.z
End Sub




'Compute all parameters
'----------------------------------------
'Name: Compute3
'----------------------------------------
'----------------------------------------
'Name: Compute3
'Description:
'----------------------------------------
Private Sub Compute3()
  
  
  BoxMatrix = GLOB.MATRIX_Scale_Rotate_MOV(OBJET.Vscal, OBJET.Vrotate, OBJET.vPosition)
  

  
  
  'BoxMatrix = CreateMatrix(OBJET.Vrotate, OBJET.Vscal, OBJET.vPosition)
  GLOB.VerTEX_MatMULTILPY OBJET.StaticVert, BoxMatrix, OBJET.TransVert
  
  'get the transfomed box
  OBJECT_getBoundingBOX OBJET.TransVert, BBOX
  
  OBJET.radius = MAX_3(BBOX.max.x - BBOX.min.x, BBOX.max.y - BBOX.min.y, BBOX.max.z - BBOX.min.z)
  OBJET.vCenter = Vmiddle(BBOX.max, BBOX.min)
  OBJET.Vsize = Vector(BBOX.max.x - BBOX.min.x, BBOX.max.y - BBOX.min.y, BBOX.max.z - BBOX.min.z)
  

End Sub


'Sub RenDer_TransparentBOX(Engine As NEMO_DX7)

'glob.LpGLOBAL_NEMO.FX_alphaBlend_FrameTransparency 1
 'glob.LpGLOBAL_NEMO.OBJECT_Scale_Rotate_MovEX OBJET.Vscal, OBJET.Vrotate, OBJET.vPosition
 'glob.LpGLOBAL_NEMO.RenderCUBE OBJET.StaticVert
 'glob.LpGLOBAL_NEMO.CLearMATRIX
'glob.LpGLOBAL_NEMO.FX_alphaBlend_FrameTransparency 0
'End Sub


'This one is Faster than the previous one
'----------------------------------------
'Name: Compute
'----------------------------------------
'----------------------------------------
'Name: Compute
'Description:
'----------------------------------------
Private Sub Compute()




    OBJET.VC(0) = VICTOR(OBJET.OrgBOX.max.x, OBJET.OrgBOX.min.y, OBJET.OrgBOX.min.z)
    OBJET.VC(1) = VICTOR(OBJET.OrgBOX.max.x, OBJET.OrgBOX.min.y, OBJET.OrgBOX.max.z)
    OBJET.VC(2) = VICTOR(OBJET.OrgBOX.min.x, OBJET.OrgBOX.min.y, OBJET.OrgBOX.max.z)
    OBJET.VC(3) = VICTOR(OBJET.OrgBOX.min.x, OBJET.OrgBOX.min.y, OBJET.OrgBOX.min.z)
    OBJET.VC(4) = VICTOR(OBJET.OrgBOX.max.x, OBJET.OrgBOX.max.y, OBJET.OrgBOX.min.z)
    OBJET.VC(5) = VICTOR(OBJET.OrgBOX.max.x, OBJET.OrgBOX.max.y, OBJET.OrgBOX.max.z)
    OBJET.VC(6) = VICTOR(OBJET.OrgBOX.min.x, OBJET.OrgBOX.max.y, OBJET.OrgBOX.max.z)
    OBJET.VC(7) = VICTOR(OBJET.OrgBOX.min.x, OBJET.OrgBOX.max.y, OBJET.OrgBOX.min.z)


  'BoxMatrix = GLOB.MATRIX_Scale_Rotate_MOV(OBJET.Vscal, OBJET.Vrotate, OBJET.vPosition)

  BoxMatrix = GLOB.Matrix_GetEX(OBJET.Vscal, OBJET.Vrotate, OBJET.vPosition)
  
 
    D3DXVec3TransformCoord OBJET.VC(0), OBJET.VC(0), BoxMatrix
    D3DXVec3TransformCoord OBJET.VC(1), OBJET.VC(1), BoxMatrix
    D3DXVec3TransformCoord OBJET.VC(2), OBJET.VC(2), BoxMatrix
    D3DXVec3TransformCoord OBJET.VC(3), OBJET.VC(3), BoxMatrix
    D3DXVec3TransformCoord OBJET.VC(4), OBJET.VC(4), BoxMatrix
    D3DXVec3TransformCoord OBJET.VC(5), OBJET.VC(5), BoxMatrix
    D3DXVec3TransformCoord OBJET.VC(6), OBJET.VC(6), BoxMatrix
    D3DXVec3TransformCoord OBJET.VC(7), OBJET.VC(7), BoxMatrix


  'K = objet.Vc(0).X
  
  'update the box
    BBOX.max.x = maxiHeight(OBJET.VC(0).x, OBJET.VC(1).x, OBJET.VC(2).x, OBJET.VC(3).x, OBJET.VC(4).x, OBJET.VC(5).x, OBJET.VC(6).x, OBJET.VC(7).x)
    BBOX.max.y = maxiHeight(OBJET.VC(0).y, OBJET.VC(1).y, OBJET.VC(2).y, OBJET.VC(3).y, OBJET.VC(4).y, OBJET.VC(5).y, OBJET.VC(6).y, OBJET.VC(7).y)
    BBOX.max.z = maxiHeight(OBJET.VC(0).z, OBJET.VC(1).z, OBJET.VC(2).z, OBJET.VC(3).z, OBJET.VC(4).z, OBJET.VC(5).z, OBJET.VC(6).z, OBJET.VC(7).z)

     BBOX.min.x = MiniHeight(OBJET.VC(0).x, OBJET.VC(1).x, OBJET.VC(2).x, OBJET.VC(3).x, OBJET.VC(4).x, OBJET.VC(5).x, OBJET.VC(6).x, OBJET.VC(7).x)
     BBOX.min.y = MiniHeight(OBJET.VC(0).y, OBJET.VC(1).y, OBJET.VC(2).y, OBJET.VC(3).y, OBJET.VC(4).y, OBJET.VC(5).y, OBJET.VC(6).y, OBJET.VC(7).y)
     BBOX.min.z = MiniHeight(OBJET.VC(0).z, OBJET.VC(1).z, OBJET.VC(2).z, OBJET.VC(3).z, OBJET.VC(4).z, OBJET.VC(5).z, OBJET.VC(6).z, OBJET.VC(7).z)

  
  OBJET.radius = MAX_3(BBOX.max.x - BBOX.min.x, BBOX.max.y - BBOX.min.y, BBOX.max.z - BBOX.min.z) / 2
  OBJET.vCenter = Vmiddle(BBOX.max, BBOX.min)
  OBJET.Vsize = Vector(BBOX.max.x - BBOX.min.x, BBOX.max.y - BBOX.min.y, BBOX.max.z - BBOX.min.z)
  
 
End Sub


'----------------------------------------
'Name: PushVertex
'----------------------------------------
'----------------------------------------
'Name: PushVertex
'Description:
'----------------------------------------
Sub PushVertex(VERT As D3DVECTOR)

    If VERT.x < BBOX.min.x Then BBOX.min.x = VERT.x
    If VERT.y < BBOX.min.y Then BBOX.min.y = VERT.y
    If VERT.z < BBOX.min.z Then BBOX.min.z = VERT.z
    
    
    If VERT.x > BBOX.max.x Then BBOX.max.x = VERT.x
    If VERT.y > BBOX.max.y Then BBOX.max.y = VERT.y
    If VERT.z > BBOX.max.z Then BBOX.max.z = VERT.z

End Sub


'----------------------------------------
'Name: ApplyMatrix
'----------------------------------------
'----------------------------------------
'Name: ApplyMatrix
'Description:
'----------------------------------------
Sub ApplyMatrix(oBjectMat As D3DMATRIX)
  'BoxMatrix = CreateMatrix(OBJET.Vrotate, OBJET.Vscal, OBJET.vPosition)
  GLOB.VerTEX_MatMULTILPY OBJET.StaticVert, oBjectMat, OBJET.TransVert
  '
  BoxMatrix = oBjectMat
  'get the transfomed box
  OBJECT_getBoundingBOX OBJET.TransVert, BBOX
  
End Sub

'Out parameters
'----------------------------------------
'Name: Get_BoxRadius
'Object: Get
'Event: BoxRadius
'----------------------------------------
'----------------------------------------
'Name: Get_BoxRadius
'Object: Get
'Event: BoxRadius
'Description:
'----------------------------------------
Function Get_BoxRadius() As Single
   Get_BoxRadius = OBJET.radius
End Function


'----------------------------------------
'Name: Get_BoxSize
'Object: Get
'Event: BoxSize
'----------------------------------------
'----------------------------------------
'Name: Get_BoxSize
'Object: Get
'Event: BoxSize
'Description:
'----------------------------------------
Function Get_BoxSize() As D3DVECTOR
  Get_BoxSize = OBJET.Vsize
End Function



'----------------------------------------
'Name: Get_BoxCenter
'Object: Get
'Event: BoxCenter
'----------------------------------------
'----------------------------------------
'Name: Get_BoxCenter
'Object: Get
'Event: BoxCenter
'Description:
'----------------------------------------
Function Get_BoxCenter() As D3DVECTOR
       Get_BoxCenter = OBJET.vCenter
End Function


'----------------------------------------
'Name: Get_Position
'Object: Get
'Event: Position
'----------------------------------------
'----------------------------------------
'Name: Get_Position
'Object: Get
'Event: Position
'Description:
'----------------------------------------
Function Get_Position() As D3DVECTOR
  Get_Position = OBJET.vPosition
End Function



'----------------------------------------
'Name: Get_Rotation
'Object: Get
'Event: Rotation
'----------------------------------------
'----------------------------------------
'Name: Get_Rotation
'Object: Get
'Event: Rotation
'Description:
'----------------------------------------
Function Get_Rotation() As D3DVECTOR
  Get_Rotation = OBJET.Vrotate
End Function



'----------------------------------------
'Name: Get_Scale
'Object: Get
'Event: Scale
'----------------------------------------
'----------------------------------------
'Name: Get_Scale
'Object: Get
'Event: Scale
'Description:
'----------------------------------------
Function Get_Scale() As D3DVECTOR
  Get_Scale = OBJET.Vscal
End Function



'----------------------------------------
'Name: Get_BoxMatrix
'Object: Get
'Event: BoxMatrix
'----------------------------------------
'----------------------------------------
'Name: Get_BoxMatrix
'Object: Get
'Event: BoxMatrix
'Description:
'----------------------------------------
Function Get_BoxMatrix() As D3DMATRIX
  Get_BoxMatrix = BoxMatrix
End Function


'----------------------------------------
'Name: Get_BoundingBOX
'Object: Get
'Event: BoundingBOX
'----------------------------------------
'----------------------------------------
'Name: Get_BoundingBOX
'Object: Get
'Event: BoundingBOX
'Description:
'----------------------------------------
Sub Get_BoundingBOX(ByRef BoxMin As D3DVECTOR, BoxMax As D3DVECTOR)
  BoxMin = BBOX.min
  BoxMax = BBOX.max

End Sub


'----------------------------------------
'Name: Get BoxMin
'----------------------------------------
'----------------------------------------
'Name: Get BoxMin
'Description:
'----------------------------------------
Public Property Get BoxMin() As D3DVECTOR
      BoxMin = BBOX.min
      
End Property


'----------------------------------------
'Name: Get BoxMax
'----------------------------------------
'----------------------------------------
'Name: Get BoxMax
'Description:
'----------------------------------------
Public Property Get BoxMax() As D3DVECTOR
      BoxMax = BBOX.max
      
End Property


'----------------------------------------
'Name: Get_BoundingSphere
'Object: Get
'Event: BoundingSphere
'----------------------------------------
'----------------------------------------
'Name: Get_BoundingSphere
'Object: Get
'Event: BoundingSphere
'Description:
'----------------------------------------
Sub Get_BoundingSphere(ByRef RetSphereCenter As D3DVECTOR, ByRef RetRadius)
  RetSphereCenter = OBJET.vCenter
  RetRadius = OBJET.radius
End Sub

'Sub Get_BoundingBox(RetBOX As D3DRMBOX)
   'RetBOX = bBOX
'End Sub


'Transformations
'----------------------------------------
'Name: Set_Position
'Object: Set
'Event: Position
'----------------------------------------
'----------------------------------------
'Name: Set_Position
'Object: Set
'Event: Position
'Description:
'----------------------------------------
Sub Set_Position(x, y, z)
  OBJET.vPosition.x = x
  OBJET.vPosition.y = y
  OBJET.vPosition.z = z
  
  
  'bBOX.min.X = 'bBOX.min.X + X
  'bBOX.min.Y = 'bBOX.min.Y + Y
  'bBOX.min.Z = 'bBOX.min.Z + Z
  
  'bBOX.max.X = 'bBOX.max.X + X
  'bBOX.max.Y = 'bBOX.max.Y + Y
  'bBOX.max.Z = 'bBOX.max.Z + Z
  
  'fill the box and sphere parameters
    Compute
End Sub


'----------------------------------------
'Name: Set_Rotation
'Object: Set
'Event: Rotation
'----------------------------------------
'----------------------------------------
'Name: Set_Rotation
'Object: Set
'Event: Rotation
'Description:
'----------------------------------------
Sub Set_Rotation(x As Single, y As Single, z As Single)
  
  
  
  
  OBJET.Vrotate.x = x
  OBJET.Vrotate.y = y
  OBJET.Vrotate.z = z
  
  'fill the box and sphere parameters
    Compute
  
End Sub



'----------------------------------------
'Name: Set_Scale
'Object: Set
'Event: Scale
'----------------------------------------
'----------------------------------------
'Name: Set_Scale
'Object: Set
'Event: Scale
'Description:
'----------------------------------------
 Sub Set_Scale(x, y, z)
  Dim tmp As Single
  
  OBJET.Vscal.x = x
  OBJET.Vscal.y = y
  OBJET.Vscal.z = z
  
 
    'fill the box and sphere parameters
    Compute
End Sub



'===================================
'*********************Transformation PARAM*****************'




'----------------------------------------
'Name: Get Vscal
'----------------------------------------
'----------------------------------------
'Name: Get Vscal
'Description:
'----------------------------------------
Public Property Get Vscal() As D3DVECTOR
      Vscal = OBJET.Vscal
      Compute
End Property


'----------------------------------------
'Name: Let Vscal
'----------------------------------------
'----------------------------------------
'Name: Let Vscal
'Description:
'----------------------------------------
Public Property Let Vscal(vNew As D3DVECTOR)
      OBJET.Vscal = vNew
      Compute

End Property


'----------------------------------------
'Name: Get vPosition
'----------------------------------------
'----------------------------------------
'Name: Get vPosition
'Description:
'----------------------------------------
Public Property Get vPosition() As D3DVECTOR
      vPosition = OBJET.vPosition
      
End Property


'----------------------------------------
'Name: Let vPosition
'----------------------------------------
'----------------------------------------
'Name: Let vPosition
'Description:
'----------------------------------------
Public Property Let vPosition(vNew As D3DVECTOR)
      OBJET.vPosition = vNew
      Compute
End Property



'----------------------------------------
'Name: Get Vrotation
'----------------------------------------
'----------------------------------------
'Name: Get Vrotation
'Description:
'----------------------------------------
Public Property Get Vrotation() As D3DVECTOR
      Vrotation = OBJET.Vrotate
      
End Property


'----------------------------------------
'Name: Let Vrotation
'----------------------------------------
'----------------------------------------
'Name: Let Vrotation
'Description:
'----------------------------------------
Public Property Let Vrotation(vNew As D3DVECTOR)
      OBJET.Vrotate = vNew
      Compute
End Property





'----------------------------------------
'Name: Set_Scale2
'Object: Set
'Event: Scale2
'----------------------------------------
'----------------------------------------
'Name: Set_Scale2
'Object: Set
'Event: Scale2
'Description:
'----------------------------------------
Private Sub Set_Scale2(x, y, z)
  Dim tmp As Single
  
  OBJET.Vscal.x = x
  OBJET.Vscal.y = y
  OBJET.Vscal.z = z
  
  BBOX.min.x = BBOX.min.x * x
  BBOX.min.y = BBOX.min.y * y
  BBOX.min.z = BBOX.min.z * z
  
  BBOX.max.x = BBOX.max.x * x
  BBOX.max.y = BBOX.max.y * y
  BBOX.max.z = BBOX.max.z * z
  
  
  'check for inverted values
    If (BBOX.max.x < BBOX.min.x) Then
        tmp = BBOX.max.x
        BBOX.max.x = BBOX.min.x
        BBOX.min.x = tmp
    End If
    
    If (BBOX.max.y < BBOX.min.y) Then
        tmp = BBOX.max.y
        BBOX.max.y = BBOX.min.y
        BBOX.min.y = tmp
    End If
    
    If (BBOX.max.z < BBOX.min.z) Then
        tmp = BBOX.max.z
        BBOX.max.z = BBOX.min.z
        BBOX.min.z = tmp
    End If
    
    'fill the box and sphere parameters
    Compute
End Sub


'Sub Apply_Matrix(TransformationMatrix As D3DMATRIX)


'first get Rotate angles






'End Sub


'Sub Apply_Quaternion(Quat As D3DQUATERNION)



'End Sub


'tools functions
'----------------------------------------
'Name: maxiHeight
'----------------------------------------
'----------------------------------------
'Name: maxiHeight
'Description:
'----------------------------------------
Private Function maxiHeight(n1, n2, N3, N4, N5, N6, N7, N8)
    Dim maxi: maxi = n1
    
    If (n2 > maxi) Then maxi = n2
    If (N3 > maxi) Then maxi = N3
    If (N4 > maxi) Then maxi = N4
    If (N5 > maxi) Then maxi = N5
    If (N6 > maxi) Then maxi = N6
    If (N7 > maxi) Then maxi = N7
    If (N8 > maxi) Then maxi = N8
    maxiHeight = maxi
End Function


'----------------------------------------
'Name: MiniHeight
'----------------------------------------
'----------------------------------------
'Name: MiniHeight
'Description:
'----------------------------------------
Private Function MiniHeight(n1, n2, N3, N4, N5, N6, N7, N8)

Dim Mini: Mini = n1

    If (n2 < Mini) Then Mini = n2
    If (N3 < Mini) Then Mini = N3
    If (N4 < Mini) Then Mini = N4
    If (N5 < Mini) Then Mini = N5
    If (N6 < Mini) Then Mini = N6
    If (N7 < Mini) Then Mini = N7
    If (N8 < Mini) Then Mini = N8
    MiniHeight = Mini

End Function


'colision Method
'tools
'----------------------------------------
'Name: SetBoundingVertex
'----------------------------------------
'----------------------------------------
'Name: SetBoundingVertex
'Description:
'----------------------------------------
Private Sub SetBoundingVertex()

'Create the Bounding Cube Vextex
 'CreateCubeEX Bvert(), VICTOR(OBJET.Vsize.X, OBJET.Vsize.Y, OBJET.Vsize.Z)



End Sub



'return an array with Transformed Vertext
'----------------------------------------
'Name: Get_BoundingVertext
'Object: Get
'Event: BoundingVertext
'----------------------------------------
'----------------------------------------
'Name: Get_BoundingVertext
'Object: Get
'Event: BoundingVertext
'Description:
'----------------------------------------
Public Sub Get_BoundingVertext(ByRef RetVertArray() As D3DVERTEX)
'Dim Ret() As D3DVERTEX
  'VerTEX_MatMULTILPY Bvert(), BoxMatrix, Ret()
  
  ReDim RetVertArray(LBound(OBJET.TransVert) To UBound(OBJET.TransVert))
  RetVertArray = OBJET.TransVert
End Sub


'Point intersection
'----------------------------------------
'Name: IS_Inside
'Object: IS
'Event: Inside
'----------------------------------------
'----------------------------------------
'Name: IS_Inside
'Object: IS
'Event: Inside
'Description:
'----------------------------------------
Function IS_Inside(Tpoint As D3DVECTOR) As Boolean
Dim ST
ST = Name
  If (Tpoint.x >= BBOX.min.x And _
      Tpoint.x <= BBOX.max.x And _
      Tpoint.y >= BBOX.min.y And _
      Tpoint.y <= BBOX.max.y And _
      Tpoint.z >= BBOX.min.z And _
      Tpoint.z <= BBOX.max.z) Then IS_Inside = True
End Function


'----------------------------------------
'Name: IS_InsideEX
'Object: IS
'Event: InsideEX
'----------------------------------------
'----------------------------------------
'Name: IS_InsideEX
'Object: IS
'Event: InsideEX
'Description:
'----------------------------------------
Function IS_InsideEX(x, y, z) As Boolean
  If (x >= BBOX.min.x And x <= BBOX.max.x And y >= BBOX.min.y And y <= BBOX.max.y And z >= BBOX.min.z And z <= BBOX.max.z) Then IS_InsideEX = True
End Function



'----------------------------------------
'Name: Is_RayIntersectEX
'Object: Is
'Event: RayIntersectEX
'----------------------------------------
'----------------------------------------
'Name: Is_RayIntersectEX
'Object: Is
'Event: RayIntersectEX
'Description:
'----------------------------------------
Function Is_RayIntersectEX(vPos As D3DVECTOR, vDir As D3DVECTOR) As Boolean

If GLOB.RayTri(OBJET.VC(0), OBJET.VC(1), OBJET.VC(2), vDir, vPos, 0, 0, 0) Then Is_RayIntersectEX = True

If GLOB.RayTri(OBJET.VC(1), OBJET.VC(2), OBJET.VC(3), vDir, vPos, 0, 0, 0) Then Is_RayIntersectEX = True

If GLOB.RayTri(OBJET.VC(2), OBJET.VC(3), OBJET.VC(4), vDir, vPos, 0, 0, 0) Then Is_RayIntersectEX = True

If GLOB.RayTri(OBJET.VC(3), OBJET.VC(4), OBJET.VC(5), vDir, vPos, 0, 0, 0) Then Is_RayIntersectEX = True

If GLOB.RayTri(OBJET.VC(4), OBJET.VC(5), OBJET.VC(6), vDir, vPos, 0, 0, 0) Then Is_RayIntersectEX = True

If GLOB.RayTri(OBJET.VC(5), OBJET.VC(6), OBJET.VC(7), vDir, vPos, 0, 0, 0) Then Is_RayIntersectEX = True

If GLOB.RayTri(OBJET.VC(3), OBJET.VC(1), OBJET.VC(5), vDir, vPos, 0, 0, 0) Then Is_RayIntersectEX = True

If GLOB.RayTri(OBJET.VC(1), OBJET.VC(5), OBJET.VC(6), vDir, vPos, 0, 0, 0) Then Is_RayIntersectEX = True

If GLOB.RayTri(OBJET.VC(2), OBJET.VC(0), OBJET.VC(4), vDir, vPos, 0, 0, 0) Then Is_RayIntersectEX = True

If GLOB.RayTri(OBJET.VC(0), OBJET.VC(4), OBJET.VC(7), vDir, vPos, 0, 0, 0) Then Is_RayIntersectEX = True

End Function


'rayIntersection
'----------------------------------------
'Name: IS_RayIntersect
'Object: IS
'Event: RayIntersect
'----------------------------------------
'----------------------------------------
'Name: IS_RayIntersect
'Object: IS
'Event: RayIntersect
'Description:
'----------------------------------------
Function IS_RayIntersect(RayStart As D3DVECTOR, RayEnd As D3DVECTOR) As Boolean


  '// an infinite ray would intersect the box but if it is not
  '// we must check if the ray is out side the box
  If (RayStart.x > BBOX.max.x And RayEnd.x > BBOX.max.x) Then
     IS_RayIntersect = False
     Exit Function
  End If
  If (RayStart.y > BBOX.max.y And RayEnd.y > BBOX.max.y) Then
     IS_RayIntersect = False
     Exit Function
  End If
  If (RayStart.z > BBOX.max.z And RayEnd.z > BBOX.max.z) Then
     IS_RayIntersect = False
     Exit Function
  End If
  If (RayStart.x < BBOX.min.x And RayEnd.x < BBOX.min.x) Then
     IS_RayIntersect = False
     Exit Function
  End If
  If (RayStart.y < BBOX.min.y And RayEnd.y < BBOX.min.y) Then
     IS_RayIntersect = False
     Exit Function
  End If
  If (RayStart.z < BBOX.min.z And RayEnd.z < BBOX.min.z) Then
     IS_RayIntersect = False
     Exit Function
  End If
  
    If (IS_InsideEX(RayStart.x, RayStart.y, RayStart.z)) Then
     IS_RayIntersect = True
     Exit Function
  End If
    If (IS_InsideEX(RayEnd.x, RayEnd.y, RayEnd.z)) Then
     IS_RayIntersect = True
     Exit Function
  End If
    
    
    Dim Tx, Ty, Tz, Lent, Sx, Sy, SZ, Plen
    
     Tx = RayEnd.x - RayStart.x
     Ty = RayEnd.y - RayStart.y
     Tz = RayEnd.z - RayStart.z

     Lent = Sqr(Tx * Tx + Ty * Ty + Tz * Tz)

    If (Lent = 0#) Then
      IS_RayIntersect = False
      Exit Function
    End If


     Sx = (BBOX.max.x + BBOX.min.x) / 2 - RayStart.x
     Sy = (BBOX.max.y + BBOX.min.y) / 2 - RayStart.y
     SZ = (BBOX.max.z + BBOX.min.z) / 2 - RayStart.z

    '// normalize t
    Tx = Tx / Lent
    Ty = Ty / Lent
    Tz = Tz / Lent

     Plen = Tx * Sx + Ty * Sy + Tz * SZ

    '// t is the bbox center projection point on the ray
    Tx = Tx * Plen
    Ty = Ty * Plen
    Tz = Tz * Plen

    If (IS_InsideEX(Tx + RayStart.x, Ty + RayStart.y, Tz + RayStart.z)) Then
       IS_RayIntersect = True
     Exit Function
  End If
    


End Function




'line intersection
'----------------------------------------
'Name: IS_LineIntersect
'Object: IS
'Event: LineIntersect
'----------------------------------------
'----------------------------------------
'Name: IS_LineIntersect
'Object: IS
'Event: LineIntersect
'Description:
'----------------------------------------
Private Function IS_LineIntersect(x1min, x1max, x2min, x2max) As Boolean

If (x1min <= x2min And x2min <= x1max) Then
       IS_LineIntersect = True
     Exit Function
  End If
    If (x1min <= x2max And x2max <= x1max) Then
       IS_LineIntersect = True
     Exit Function
  End If
    If (x2min <= x1min And x1min <= x2max) Then
       IS_LineIntersect = True
     Exit Function
  End If
    If (x2min <= x1max And x1max <= x2max) Then
       IS_LineIntersect = True
     Exit Function
  End If

End Function




'Sphere bounding colision
'----------------------------------------
'Name: IS_SphereIntersect
'Object: IS
'Event: SphereIntersect
'----------------------------------------
'----------------------------------------
'Name: IS_SphereIntersect
'Object: IS
'Event: SphereIntersect
'Description:
'----------------------------------------
Function IS_SphereIntersect(VphereCenter As D3DVECTOR, SphereRadius) As Boolean
    Dim smaxX, smaxY, smaxZ, sminX, sminY, sminZ
     smaxX = VphereCenter.x + SphereRadius
     smaxY = VphereCenter.y + SphereRadius
     smaxZ = VphereCenter.z + SphereRadius
     sminX = VphereCenter.x - SphereRadius
     sminY = VphereCenter.y - SphereRadius
     sminZ = VphereCenter.z - SphereRadius
    If (IS_LineIntersect(BBOX.min.x, BBOX.max.x, sminX, smaxX) And _
        IS_LineIntersect(BBOX.min.y, BBOX.max.y, sminY, smaxY) And _
        IS_LineIntersect(BBOX.min.z, BBOX.max.z, sminZ, smaxZ)) Then IS_SphereIntersect = True


End Function


'BOX BOX colision
'Function IS_BoxIntersect(cBOX As D3DRMBOX) As Boolean

    'If (IS_LineIntersect(bBOX.min.X, bBOX.max.X, cBOX.min.X, cBOX.max.X) And _
        IS_LineIntersect(bBOX.min.Y, bBOX.max.Y, cBOX.min.Y, cBOX.max.Y) And _
        IS_LineIntersect(bBOX.min.Z, bBOX.max.Z, cBOX.min.Z, cBOX.max.Z)) Then IS_BoxIntersect = True
    
'End Function



'Box Tri Tri
'----------------------------------------
'Name: IS_TriIntersect
'Object: IS
'Event: TriIntersect
'----------------------------------------
'----------------------------------------
'Name: IS_TriIntersect
'Object: IS
'Event: TriIntersect
'Description:
'----------------------------------------
Function IS_TriIntersect(p1 As D3DVECTOR, p2 As D3DVECTOR, P3 As D3DVECTOR) As Boolean
  
  If IS_Inside(p1) Then
     IS_TriIntersect = True
     Exit Function
  End If
  
   If IS_Inside(p2) Then
     IS_TriIntersect = True
     Exit Function
  End If
  
   If IS_Inside(P3) Then
     IS_TriIntersect = True
     Exit Function
  End If
  
  
End Function



'----------------------------------------
'Name: IS_TriIntersectEX
'Object: IS
'Event: TriIntersectEX
'----------------------------------------
'----------------------------------------
'Name: IS_TriIntersectEX
'Object: IS
'Event: TriIntersectEX
'Description:
'----------------------------------------
Function IS_TriIntersectEX(p1 As D3DVERTEX, p2 As D3DVERTEX, P3 As D3DVERTEX) As Boolean
  
  If IS_InsideEX(p1.x, p1.y, p1.z) Then
     IS_TriIntersectEX = True
     Exit Function
  End If
  
If IS_InsideEX(p2.x, p2.y, p2.z) Then
     IS_TriIntersectEX = True
     Exit Function
  End If
  
  If IS_InsideEX(P3.x, P3.y, P3.z) Then
     IS_TriIntersectEX = True
     Exit Function
  End If
  
  
End Function





















'----------------------------------------
'Name: OBJECT_getBoundingBOX
'Object: OBJECT
'Event: getBoundingBOX
'----------------------------------------
'----------------------------------------
'Name: OBJECT_getBoundingBOX
'Object: OBJECT
'Event: getBoundingBOX
'Description:
'----------------------------------------
Private Sub OBJECT_getBoundingBOX(lpVERT() As D3DVERTEX, ByRef RETBOX As D3DRMBOX)
Dim i, j, ZZ
Dim VEC As D3DVECTOR

' bounding box
   Dim min As D3DVECTOR
   Dim max As D3DVECTOR
   min = Vector(99999999, 999999999, 999999999)
   max = Vector(-999999999, -999999999, -999999999)
   
 i = LBound(lpVERT)
 j = UBound(lpVERT)
 
   
For ZZ = i To j
   With lpVERT(ZZ)
       If .x > max.x Then max.x = .x
       If .y > max.y Then max.y = .y
       If .z > max.z Then max.z = .z
       If .x < min.x Then min.x = .x
       If .y < min.y Then min.y = .y
       If .z < min.z Then min.z = .z
       
      End With
Next ZZ
RETBOX.max = max
RETBOX.min = min
End Sub







