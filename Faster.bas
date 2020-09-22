Attribute VB_Name = "Engine"
Option Explicit


Type VertexDis
    X As Single
    Y As Single
    Z As Single
End Type

Type Coord
    X As Single
    Y As Single
    Z As Single
    Target As Integer
End Type

Type SkelitonDis
    Origin As VertexDis
    Target As Integer
    Name As String
    Huuh As String
End Type

Type WeaponDis
    Name As String
    Type As String
    Joint As Integer
    Vangle As Integer
    HAngle As Integer
End Type

Type Objects
    ID As String
    VertexCount As Integer
    Vertex() As Coord
    FaceCount As Integer
    Face() As Integer
    EdgeCount() As Integer
    SkelitonCount As Integer
    Skeliton() As SkelitonDis
    WeaponCount As Byte
    Weapon() As WeaponDis
    ForceSkeliton As Boolean
End Type

Type Morph
    Origin As VertexDis
    Angle As VertexDis
    Scale As VertexDis
End Type


Type CamaraDis
    Origin As VertexDis
    Angle As VertexDis
End Type

'###########################################

Type WorldDis
    Morph() As Morph
    Target As Integer
    Angle As VertexDis
    Origin As VertexDis
    Push As VertexDis
    ModelID As Integer
    Used As Boolean
    Speed As Single
    Timer As Integer
    Colour As Long
End Type

'###########################################

Dim Comp() As Objects
Public SINe(-361 To 361) As Double
Public COSine(-361 To 361) As Double
Const PI = 3.14159265358979
Dim TempVert(2500) As Coord
Dim Frames As Integer, Tima As Single
Dim TempMorph() As Morph
Public Xf As Integer, Yf As Integer
Const Zeye = 800
Dim BigSkel As Integer
Public WorldCount
Public NewFrame As Boolean
Dim Rotates As Coord

Public Function SetWorldToModel(WorldID, ModelID, World() As WorldDis) As Boolean
On Error GoTo Couldnt_Set_Model
    ReDim World(WorldID).Morph(Comp(ModelID).SkelitonCount) As Morph
    World(WorldID).ModelID = ModelID
    World(WorldID).Used = True
    SetWorldToModel = True
Couldnt_Set_Model:
End Function

Public Sub Make_LookUp()
    Dim I As Integer
    For I = -361 To 361
        SINe(I) = Sin(I / 180 * PI)
        COSine(I) = Cos(I / 180 * PI)
    Next
    ReDim Comp(1) As Objects
End Sub

Public Function LoadCompressedModel(FileName As String, ModelID As Integer) As Boolean
        On Error GoTo NotAComp
        Dim Test, N As Integer, FaceEdge As Integer, M As Integer, X As Integer
        Dim ForceSkel As Integer, Skel As Integer, Y As Integer, Z As Integer, WeKnow As Integer
        Dim Gun As Integer
        If ModelID > UBound(Comp()) Then ReDim Preserve Comp(ModelID) As Objects
        Open FileName For Input As #1
            Input #1, Test, Comp(ModelID).ID
            Line Input #1, Test
            Line Input #1, Test
            Line Input #1, Test
            Input #1, Comp(ModelID).VertexCount
            ReDim Comp(ModelID).Vertex(Comp(ModelID).VertexCount) As Coord
            Input #1, Comp(ModelID).FaceCount
            ReDim Comp(ModelID).Face(Comp(ModelID).FaceCount, 14) As Integer
            ReDim Comp(ModelID).EdgeCount(Comp(ModelID).FaceCount) As Integer
            For N = 1 To Comp(ModelID).VertexCount
                Input #1, Comp(ModelID).Vertex(N).X
                Input #1, Comp(ModelID).Vertex(N).Y
                Input #1, Comp(ModelID).Vertex(N).Z
                Input #1, Comp(ModelID).Vertex(N).Target
            Next N
            For N = 1 To Comp(ModelID).FaceCount
                Input #1, FaceEdge
                Comp(ModelID).EdgeCount(N) = FaceEdge
                For M = 1 To FaceEdge
                    Input #1, X: Comp(ModelID).Face(N, M) = X + 1
                Next M
            Next N
            Input #1, ForceSkel
            Input #1, Skel
            If ForceSkel = 1 Then Comp(ModelID).ForceSkeliton = True
            If Skel > BigSkel Then BigSkel = Skel
            Comp(ModelID).SkelitonCount = Skel
            ReDim Comp(ModelID).Skeliton(Skel) As SkelitonDis
            For N = 1 To Skel
                Input #1, X
                Input #1, Y
                Input #1, Z
                Input #1, WeKnow
                Comp(ModelID).Skeliton(WeKnow).Origin.X = X
                Comp(ModelID).Skeliton(WeKnow).Origin.Y = Y
                Comp(ModelID).Skeliton(WeKnow).Origin.Z = Z
                Input #1, Comp(ModelID).Skeliton(WeKnow).Target
                Input #1, Comp(ModelID).Skeliton(WeKnow).Name
                Input #1, Comp(ModelID).Skeliton(WeKnow).Huuh
            Next N
            Input #1, Gun
            Comp(ModelID).WeaponCount = Gun
            ReDim Comp(ModelID).Weapon(Gun) As WeaponDis
            For N = 1 To Gun
                Input #1, Comp(ModelID).Weapon(N).Joint
                Input #1, Comp(ModelID).Weapon(N).Name
                Input #1, Comp(ModelID).Weapon(N).Type
                Input #1, Comp(ModelID).Weapon(N).Vangle
                Input #1, Comp(ModelID).Weapon(N).HAngle
            Next N
        Close
        LoadCompressedModel = True
        Exit Function
NotAComp:
    MsgBox "Model failed to load:" & vbNewLine & FileName, vbCritical, "Error"
End Function


Public Sub RunEngine(Window As PictureBox, Camara As CamaraDis, World() As WorldDis)
    Dim Unit As Integer, N As Integer
    Window.Cls
    For N = LBound(World()) To UBound(World())
        If World(N).Used = True Then
            ReDim TempMorph(Comp(World(N).ModelID).VertexCount) As Morph
            If Rotate(N, Camara, World()) = False Then DrawObject Window, N, World()
        End If
    Next N
End Sub

Private Function Rotate(WorldID As Integer, Camara As CamaraDis, World() As WorldDis) As Boolean
    Dim Oangle1 As Integer, OAngle2 As Integer, Oangle3 As Integer, nn As Integer
    Dim WAngle1 As Integer, WAngle2 As Integer, WAngle3 As Integer, Target As Byte
    Dim Mangle1 As Integer, Mangle2 As Integer, Mangle3 As Integer
    Dim X As Integer, Y As Integer, Z As Integer, XRotated As Integer, YRotated As Integer, ZRotated As Integer
    Dim Spun As Coord
    WAngle1 = Camara.Angle.X Mod 360
    WAngle2 = (360 - Camara.Angle.Y) Mod 360
    WAngle3 = Camara.Angle.Z Mod 360
    X = World(WorldID).Origin.X - Camara.Origin.X
    Y = World(WorldID).Origin.Y - Camara.Origin.Y
    Z = World(WorldID).Origin.Z - Camara.Origin.Z
    XRotated = COSine(WAngle2) * X - SINe(WAngle2) * Z:    YRotated = Y
    ZRotated = SINe(WAngle2) * X + COSine(WAngle2) * Z
    X = XRotated: Y = YRotated: Z = ZRotated:    XRotated = X
    YRotated = COSine(WAngle1) * Y - SINe(WAngle1) * Z
    ZRotated = SINe(WAngle1) * Y + COSine(WAngle1) * Z
    X = XRotated: Y = YRotated: Z = ZRotated
    XRotated = COSine(WAngle3) * X - SINe(WAngle3) * Y
    YRotated = SINe(WAngle3) * X + COSine(WAngle3) * Y
    ZRotated = Z
    Spun.X = XRotated:    Spun.Y = YRotated:    Spun.Z = ZRotated
    If Spun.Z > 650 Then Rotate = True: Exit Function
    Oangle1 = World(WorldID).Angle.X Mod 360
    OAngle2 = World(WorldID).Angle.Y Mod 360
    Oangle3 = World(WorldID).Angle.Z Mod 360
    MorphSkeliton WorldID, World()
    For nn = 1 To Comp(World(WorldID).ModelID).VertexCount
        Target = Comp(World(WorldID).ModelID).Vertex(nn).Target
        Mangle1 = TempMorph(Target).Angle.X Mod 360
        Mangle2 = TempMorph(Target).Angle.Y Mod 360
        Mangle3 = TempMorph(Target).Angle.Z Mod 360
        X = Comp(World(WorldID).ModelID).Vertex(nn).X - Comp(World(WorldID).ModelID).Skeliton(Target).Origin.X
        Y = Comp(World(WorldID).ModelID).Vertex(nn).Y - Comp(World(WorldID).ModelID).Skeliton(Target).Origin.Y
        Z = Comp(World(WorldID).ModelID).Vertex(nn).Z - Comp(World(WorldID).ModelID).Skeliton(Target).Origin.Z
        XRotated = (COSine(Mangle3) * X - SINe(Mangle3) * Y)
        YRotated = (SINe(Mangle3) * X + COSine(Mangle3) * Y)
        ZRotated = Z: X = XRotated: Y = YRotated: Z = ZRotated: XRotated = X
        YRotated = COSine(Mangle1) * Y - SINe(Mangle1) * Z
        ZRotated = SINe(Mangle1) * Y + COSine(Mangle1) * Z
        X = XRotated: Y = YRotated: Z = ZRotated
        XRotated = (COSine(Mangle2) * X - SINe(Mangle2) * Z) * (TempMorph(Target).Scale.X + 1) + TempMorph(Target).Origin.X
        YRotated = Y * (TempMorph(Target).Scale.Y + 1) + TempMorph(Target).Origin.Y
        ZRotated = (SINe(Mangle2) * X + COSine(Mangle2) * Z) * (TempMorph(Target).Scale.Z + 1) + TempMorph(Target).Origin.Z
        X = XRotated: Y = YRotated: Z = ZRotated
        XRotated = COSine(Oangle3) * X - SINe(Oangle3) * Y
        YRotated = SINe(Oangle3) * X + COSine(Oangle3) * Y
        ZRotated = Z:  X = XRotated: Y = YRotated: Z = ZRotated: XRotated = X
        YRotated = COSine(Oangle1) * Y - SINe(Oangle1) * Z
        ZRotated = SINe(Oangle1) * Y + COSine(Oangle1) * Z
        X = XRotated: Y = YRotated: Z = ZRotated
        XRotated = COSine(OAngle2) * X - SINe(OAngle2) * Z
        YRotated = Y: ZRotated = SINe(OAngle2) * X + COSine(OAngle2) * Z: X = XRotated: Y = YRotated: Z = ZRotated
        XRotated = COSine(WAngle2) * X - SINe(WAngle2) * Z: YRotated = Y
        ZRotated = SINe(WAngle2) * X + COSine(WAngle2) * Z
        X = XRotated: Y = YRotated: Z = ZRotated: XRotated = X
        YRotated = COSine(WAngle1) * Y - SINe(WAngle1) * Z
        ZRotated = SINe(WAngle1) * Y + COSine(WAngle1) * Z
        X = XRotated: Y = YRotated: Z = ZRotated
        XRotated = COSine(WAngle3) * X - SINe(WAngle3) * Y
        YRotated = SINe(WAngle3) * X + COSine(WAngle3) * Y: ZRotated = Z
        TempVert(nn).X = XRotated + Spun.X
        TempVert(nn).Y = YRotated + Spun.Y
        TempVert(nn).Z = ZRotated + Spun.Z
    Next nn
End Function

Private Sub MorphSkeliton(WorldID, World() As WorldDis)
    Dim N As Integer, Target As Byte, Cx As Integer, Cy As Integer, Cz As Integer, sAngle1 As Integer, sAngle2 As Integer, sAngle3 As Integer
    Dim X As Integer, Y As Integer, Z As Integer
    For N = 1 To Comp(World(WorldID).ModelID).SkelitonCount
        Target = Comp(World(WorldID).ModelID).Skeliton(N).Target
        If Target <> 0 Then
            TempMorph(N).Origin.X = Comp(World(WorldID).ModelID).Skeliton(N).Origin.X
            TempMorph(N).Origin.Y = Comp(World(WorldID).ModelID).Skeliton(N).Origin.Y
            TempMorph(N).Origin.Z = Comp(World(WorldID).ModelID).Skeliton(N).Origin.Z
            Cx = TempMorph(Target).Origin.X
            Cy = TempMorph(Target).Origin.Y
            Cz = TempMorph(Target).Origin.Z
            sAngle1 = TempMorph(Target).Angle.X
            sAngle2 = TempMorph(Target).Angle.Y
            sAngle3 = TempMorph(Target).Angle.Z
            X = TempMorph(Target).Origin.X + Comp(World(WorldID).ModelID).Skeliton(N).Origin.X - Comp(World(WorldID).ModelID).Skeliton(Target).Origin.X + World(WorldID).Morph(N).Origin.X
            Y = TempMorph(Target).Origin.Y + Comp(World(WorldID).ModelID).Skeliton(N).Origin.Y - Comp(World(WorldID).ModelID).Skeliton(Target).Origin.Y + World(WorldID).Morph(N).Origin.Y
            Z = TempMorph(Target).Origin.Z + Comp(World(WorldID).ModelID).Skeliton(N).Origin.Z - Comp(World(WorldID).ModelID).Skeliton(Target).Origin.Z + World(WorldID).Morph(N).Origin.Z
            TempMorph(N).Angle.X = World(WorldID).Morph(N).Angle.X + sAngle1
            TempMorph(N).Angle.Y = World(WorldID).Morph(N).Angle.Y + sAngle2
            TempMorph(N).Angle.Z = World(WorldID).Morph(N).Angle.Z + sAngle3
            RotateSkeliton sAngle1, sAngle2, sAngle3, X, Y, Z, Cx, Cy, Cz
            TempMorph(N).Origin.X = Rotates.X
            TempMorph(N).Origin.Y = Rotates.Y
            TempMorph(N).Origin.Z = Rotates.Z
            TempMorph(N).Scale.X = World(WorldID).Morph(N).Scale.X
            TempMorph(N).Scale.Y = World(WorldID).Morph(N).Scale.Y
            TempMorph(N).Scale.Z = World(WorldID).Morph(N).Scale.Z
        End If
    Next N
End Sub

Private Sub RotateSkeliton(tAngle1, tAngle2, tAngle3, X, Y, Z, Cx, Cy, Cz)
    Dim XRotated As Integer, YRotated As Integer, ZRotated As Integer
    tAngle1 = (tAngle1 Mod 360)
    tAngle2 = (tAngle2 Mod 360)
    tAngle3 = (tAngle3 Mod 360)
    X = X - Cx:    Y = Y - Cy:    Z = Z - Cz
    XRotated = COSine(tAngle3) * X - SINe(tAngle3) * Y
    YRotated = SINe(tAngle3) * X + COSine(tAngle3) * Y:    ZRotated = Z
    X = XRotated:    Y = YRotated:    Z = ZRotated
    YRotated = COSine(tAngle1) * Y - SINe(tAngle1) * Z
    ZRotated = SINe(tAngle1) * Y + COSine(tAngle1) * Z: XRotated = X
    X = XRotated:    Y = YRotated:    Z = ZRotated:    XRotated = X
    XRotated = COSine(tAngle2) * X - SINe(tAngle2) * Z:    YRotated = Y
    ZRotated = SINe(tAngle2) * X + COSine(tAngle2) * Z
    Rotates.X = XRotated + Cx:    Rotates.Y = YRotated + Cy:    Rotates.Z = ZRotated + Cz
End Sub

Private Sub DrawObject(Window As PictureBox, WorldID As Integer, World() As WorldDis)
    Dim Ner(16, 2) As Integer, N As Integer, M As Byte, s1 As Integer, Cola As Long
    Dim Xx1 As Integer, Yy1 As Integer, Xx2 As Integer, Yy2 As Integer
    On Error GoTo Yikes
    Cola = World(WorldID).Colour
    For N = 1 To Comp(World(WorldID).ModelID).FaceCount
        For M = 1 To Comp(World(WorldID).ModelID).EdgeCount(N)
            s1 = Comp(World(WorldID).ModelID).Face(N, M)
            Ner(M, 1) = Xf + Int(TempVert(s1).X * (Zeye / (Zeye - TempVert(s1).Z)))
            Ner(M, 2) = Yf + Int(TempVert(s1).Y * (Zeye / (Zeye - TempVert(s1).Z)))
        Next M
        If FaceNormal(Ner(), M) < 0 Then
            For M = 1 To Comp(World(WorldID).ModelID).EdgeCount(N) - 1
                Xx1 = Ner(M, 1): Yy1 = Ner(M, 2)
                Xx2 = Ner(M + 1, 1): Yy2 = Ner(M + 1, 2)
                Window.Line (Xx1, Yy1)-(Xx2, Yy2), Cola
            Next M
            Xx1 = Ner(M, 1): Yy1 = Ner(M, 2)
            Xx2 = Ner(1, 1): Yy2 = Ner(1, 2)
            Window.Line (Xx1, Yy1)-(Xx2, Yy2), Cola
        End If
    Next N
Yikes:
End Sub

Private Function FaceNormal(Ner() As Integer, Edges As Byte)
    On Error GoTo FailedFace
        Select Case Edges
            Case 4, 5, 6, 7, 8, 9, 10
                FaceNormal = ((Ner(1, 2) - Ner(3, 2)) * (Ner(2, 1) - Ner(1, 1))) - ((Ner(1, 1) - Ner(3, 1)) * (Ner(2, 2) - Ner(1, 2)))
            Case 11, 12, 13, 14, 15, 16, 17
                FaceNormal = ((Ner(1, 2) - Ner(9, 2)) * (Ner(6, 1) - Ner(1, 1))) - ((Ner(1, 1) - Ner(9, 1)) * (Ner(6, 2) - Ner(1, 2)))
        
        
        End Select
        Exit Function
FailedFace:
End Function
