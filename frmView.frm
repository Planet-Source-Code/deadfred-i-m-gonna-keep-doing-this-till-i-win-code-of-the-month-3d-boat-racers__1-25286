VERSION 5.00
Begin VB.Form frmView 
   Caption         =   "Animation Viewer"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer LapPass 
      Left            =   4320
      Top             =   3840
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   4320
   End
   Begin VB.PictureBox view 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   1080
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
   Begin VB.Shape Check 
      FillColor       =   &H000000FF&
      Height          =   735
      Index           =   5
      Left            =   120
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "You"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Boris the evil"
      Height          =   255
      Left            =   -120
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Boris 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.Label yourz 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Check 
      FillColor       =   &H000000FF&
      Height          =   735
      Index           =   4
      Left            =   120
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   735
   End
   Begin VB.Shape Check 
      FillColor       =   &H000000FF&
      Height          =   735
      Index           =   3
      Left            =   120
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   735
   End
   Begin VB.Shape Check 
      FillColor       =   &H000000FF&
      Height          =   735
      Index           =   2
      Left            =   120
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   735
   End
   Begin VB.Shape Check 
      FillColor       =   &H000000FF&
      Height          =   735
      Index           =   1
      Left            =   120
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   735
   End
   Begin VB.Shape Check 
      FillColor       =   &H000000FF&
      Height          =   735
      Index           =   0
      Left            =   120
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   735
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Quit As Boolean, Frames As Integer
Dim Camara As CamaraDis
Dim EvilCamara As CamaraDis
Dim HelicopterCam As CamaraDis
Dim TowerCam As CamaraDis
Dim World(35) As WorldDis
Dim VKey(255) As Boolean
Dim CamaraMode As Integer
Const Pie As Single = ((22 / 7) * 18.23)

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    VKey(KeyCode) = True
    If KeyCode = 112 Then CamaraMode = 1
    If KeyCode = 113 Then CamaraMode = 2
    If KeyCode = 114 Then CamaraMode = 3
    If KeyCode = 115 Then CamaraMode = 4
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    VKey(KeyCode) = False
    If KeyCode = 112 Then CamaraMode = 0
    If KeyCode = 113 Then CamaraMode = 0
    If KeyCode = 114 Then CamaraMode = 0
    If KeyCode = 115 Then CamaraMode = 0
End Sub

Private Sub Form_Load()
    Engine.Make_LookUp
    If Engine.LoadCompressedModel(App.Path & "\boat.dat", 1) = False Then End
    If Engine.LoadCompressedModel(App.Path & "\helicopter.dat", 2) = False Then End
    If Engine.LoadCompressedModel(App.Path & "\start.dat", 3) = False Then End
    If Engine.LoadCompressedModel(App.Path & "\check.dat", 4) = False Then End
    If Engine.LoadCompressedModel(App.Path & "\Highlight.dat", 5) = False Then End
    If Engine.LoadCompressedModel(App.Path & "\Arrow.dat", 6) = False Then End
    If Engine.LoadCompressedModel(App.Path & "\Tower.dat", 7) = False Then End
    If Engine.LoadCompressedModel(App.Path & "\Wake.dat", 8) = False Then End
    If Engine.LoadCompressedModel(App.Path & "\Jerry.dat", 9) = False Then End
    Show
    Engine.SetWorldToModel 1, 1, World()
    Engine.SetWorldToModel 2, 2, World()
    Engine.SetWorldToModel 3, 5, World()
    Engine.SetWorldToModel 4, 9, World()
    Engine.SetWorldToModel 5, 3, World()
    Engine.SetWorldToModel 6, 4, World()
    Engine.SetWorldToModel 7, 4, World()
    Engine.SetWorldToModel 8, 4, World()
    Engine.SetWorldToModel 9, 4, World()
    Engine.SetWorldToModel 10, 4, World()
    Engine.SetWorldToModel 12, 6, World()
    Engine.SetWorldToModel 13, 7, World()
    
    World(3).Colour = vbBlue
    'Set out camaras
    Camara.Origin.Y = -100
    Camara.Angle.X = -10
    EvilCamara.Angle.X = -10
    EvilCamara.Origin.Y = -100
    HelicopterCam.Origin.Y = -400
    TowerCam.Origin.Y = -300
    
    'Set out Boats
    World(1).Origin.X = -100
    World(4).Origin.X = 100
    World(1).Target = 5
    World(4).Target = 5
    
    'Set out the Track
    World(5).Origin.Z = 0
    World(5).Origin.X = 0
    World(5).Angle.Y = 0
    World(6).Origin.Z = -4000
    World(6).Origin.X = 0
    World(6).Angle.Y = 0
    World(7).Origin.Z = -4000
    World(7).Origin.X = 6000
    World(7).Angle.Y = 90
    
    World(8).Origin.Z = -500
    World(8).Origin.X = 7000
    World(8).Angle.Y = 45
    
    World(9).Origin.Z = -500
    World(9).Origin.X = 10000
    World(9).Angle.Y = 90
    
    World(10).Origin.Z = 2000
    World(10).Origin.X = 3000
    World(10).Angle.Y = 90

    'Set  out camara tower
    World(13).Origin.X = 2000
    World(13).Origin.Y = -100

    'Set the helicopter
    World(2).Origin.Y = -400
    World(2).Angle.X = 20

    'Set pointy thing
    World(12).Origin.Y = -250
    World(12).Origin.X = World(1).Origin.X
    World(12).Origin.Z = World(1).Origin.Z
    World(12).Angle.X = 20

    World(3).Origin.X = World(World(1).Target).Origin.X
    World(3).Origin.Z = World(World(1).Target).Origin.Z

    TowerCam.Origin.X = World(13).Origin.X
    TowerCam.Origin.Z = World(13).Origin.Z


    Do
        Frames = Frames + 1
        DoEvents
        MoveFlags
        MoveHelicopter
        MoveEvilBoat
        RemoveWake
        If CamaraMode = 3 Or CamaraMode = 4 Then AlingTowerCAm
        RunKeys
        Select Case CamaraMode
            Case 0:        Engine.RunEngine view, Camara, World()
            Case 1:        Engine.RunEngine view, EvilCamara, World()
            Case 2:        Engine.RunEngine view, HelicopterCam, World()
            Case 3, 4:    Engine.RunEngine view, TowerCam, World()
        End Select
    Loop While Quit = False
    End
End Sub

Private Sub AlingTowerCAm()
    If CamaraMode = 3 Then
        Angle = GetAngle(World(13).Origin.X - World(1).Origin.X, World(1).Origin.Z - World(13).Origin.Z)
    Else
        Angle = GetAngle(World(13).Origin.X - World(4).Origin.X, World(4).Origin.Z - World(13).Origin.Z)
    End If
    
    TowerCam.Angle.Y = Angle

End Sub

Private Function GetWake()
    For N = 15 To 35
        If World(N).Used = False Then GetWake = N: Exit Function
    Next N
End Function

Private Sub RemoveWake()
    For N = 15 To 35
        If World(N).Used = True Then

            World(N).Timer = World(N).Timer + 1
            If World(N).Timer = 20 Then

                World(N).Timer = 0
                World(N).Used = False
            End If
        End If
    Next N
End Sub

Private Sub MoveEvilBoat()
    Target = World(4).Target
    Dista = Dist(World(4).Origin.X, World(4).Origin.Z, World(Target).Origin.X, World(Target).Origin.Z)
    If Dista < 150 Then
        World(4).Target = World(4).Target + 1
        If World(4).Target = 12 Then World(4).Target = 5: Boris = Boris + 1
    End If
    World(4).Speed = World(4).Speed + 5
    World(4).Speed = World(4).Speed * 0.95
    Angle = GetAngle(World(4).Origin.X - World(Target).Origin.X, World(Target).Origin.Z - World(4).Origin.Z)
    If World(4).Angle.Y > Angle Then World(4).Angle.Y = World(4).Angle.Y - 5: World(4).Angle.Z = World(4).Angle.Z - (World(4).Speed * 0.25)
    If World(4).Angle.Y < Angle Then World(4).Angle.Y = World(4).Angle.Y + 5: World(4).Angle.Z = World(4).Angle.Z + (World(4).Speed * 0.25)
    World(4).Angle.Z = World(4).Angle.Z * 0.3
    X = SINe(World(4).Angle.Y) * World(4).Speed
    Z = COSine(World(4).Angle.Y) * World(4).Speed
    World(4).Origin.X = World(4).Origin.X + X
    World(4).Origin.Z = World(4).Origin.Z - Z
    EvilCamara.Origin.X = World(4).Origin.X
    EvilCamara.Origin.Z = World(4).Origin.Z
    EvilCamara.Angle.Y = World(4).Angle.Y
    
    World(4).Timer = World(4).Timer + 1
    If World(4).Timer = 8 And World(4).Speed <> 0 Then
        World(4).Timer = 0
        wake = GetWake: If wake = 0 Then Exit Sub
        Engine.SetWorldToModel wake, 8, World()
        World(wake).Origin.X = World(4).Origin.X
        World(wake).Origin.Z = World(4).Origin.Z
        World(wake).Angle.Y = World(4).Angle.Y
        World(wake).Colour = vbBlue
    End If
    
End Sub


Private Sub MoveHelicopter()
    World(2).Morph(2).Angle.X = (World(2).Morph(2).Angle.X + 10) Mod 360
    World(2).Morph(3).Angle.Y = (World(2).Morph(3).Angle.Y + 20) Mod 360
    Target = World(2).Target
    If Dist(World(2).Origin.X, World(2).Origin.Z, World(Target).Origin.X, World(Target).Origin.Z) < 150 Then
        World(2).Target = Int(Rnd * 4) + 6
    End If
    Angle = GetAngle(World(2).Origin.X - World(Target).Origin.X, World(Target).Origin.Z - World(2).Origin.Z)
    If World(2).Angle.Y > Angle Then World(2).Angle.Y = World(2).Angle.Y - 5: World(2).Angle.Z = World(2).Angle.Z - (World(2).Speed * 0.25)
    If World(2).Angle.Y < Angle Then World(2).Angle.Y = World(2).Angle.Y + 5: World(2).Angle.Z = World(2).Angle.Z + (World(2).Speed * 0.25)
    World(2).Angle.Z = World(2).Angle.Z * 0.3
    X = SINe(World(2).Angle.Y) * 20
    Z = COSine(World(2).Angle.Y) * 20
    World(2).Origin.X = World(2).Origin.X + X
    World(2).Origin.Z = World(2).Origin.Z - Z
    HelicopterCam.Origin.X = World(2).Origin.X
    HelicopterCam.Origin.Z = World(2).Origin.Z
    HelicopterCam.Angle.Y = World(2).Angle.Y
End Sub


Private Sub RunKeys()
    If VKey(38) = True Then
        World(1).Speed = World(1).Speed - 5
    End If
    If VKey(40) = True Then
        World(1).Speed = World(1).Speed + 5
    End If
    World(1).Speed = World(1).Speed * 0.95
    If World(1).Speed > 0 Then World(1).Speed = World(1).Speed * 0.8
    X = SINe(World(1).Angle.Y) * World(1).Speed
    Z = COSine(World(1).Angle.Y) * World(1).Speed
    World(1).Origin.X = World(1).Origin.X - X
    World(1).Origin.Z = World(1).Origin.Z + Z
    If VKey(37) = True Then
        World(1).Angle.Y = (World(1).Angle.Y - 10) Mod 360
        World(1).Angle.Z = World(1).Angle.Z + (World(1).Speed * 0.25)
    End If
    If VKey(39) = True Then
        World(1).Angle.Y = (World(1).Angle.Y + 10) Mod 360
        World(1).Angle.Z = World(1).Angle.Z - (World(1).Speed * 0.25)
    End If
    World(1).Angle.Z = World(1).Angle.Z * 0.3
    If Dist(World(1).Origin.X, World(1).Origin.Z, World(World(1).Target).Origin.X, World(World(1).Target).Origin.Z) < 200 Then
        World(1).Target = World(1).Target + 1
        If World(1).Target = 12 Then
            World(1).Target = 6
            yourz = yourz + 1
            LapPass.Interval = 1
            LapPass.Tag = 25
        End If
        Check(World(1).Target - 6).FillStyle = 0
        World(3).Origin.X = World(World(1).Target).Origin.X
        World(3).Origin.Z = World(World(1).Target).Origin.Z
    End If
    Camara.Angle.Y = World(1).Angle.Y
    Camara.Origin.X = World(1).Origin.X
    Camara.Origin.Z = World(1).Origin.Z
    
    CrashDis = Dist(World(1).Origin.X, World(1).Origin.Z, World(4).Origin.X, World(4).Origin.Z)
    If CrashDis < 100 Then
        Angle = GetAngle(World(1).Origin.X - World(4).Origin.X, World(4).Origin.Z - World(1).Origin.Z)
        pz = SINe(Angle) * (100 - CrashDis) * 0.75
        px = COSine(Angle) * (100 - CrashDis) * 0.75
        World(1).Push.X = px
        World(1).Push.Z = -pz
        World(4).Push.X = -px
        World(4).Push.Z = pz
        World(1).Speed = -World(1).Speed
        World(4).Speed = -World(4).Speed
    End If
    Target = World(1).Target
    Angle = GetAngle(World(1).Origin.X - World(Target).Origin.X, World(Target).Origin.Z - World(1).Origin.Z)
    World(1).Origin.X = World(1).Origin.X + World(1).Push.X
    World(1).Origin.Z = World(1).Origin.Z + World(1).Push.Z
    World(4).Origin.X = World(4).Origin.X + World(4).Push.X
    World(4).Origin.Z = World(4).Origin.Z + World(4).Push.Z
    World(1).Push.X = World(1).Push.X * 0.4
    World(1).Push.Z = World(1).Push.Z * 0.4
    World(4).Push.X = World(4).Push.X * 0.4
    World(4).Push.Z = World(4).Push.Z * 0.4
    World(12).Angle.Y = Angle
    World(12).Origin.X = World(1).Origin.X
    World(12).Origin.Z = World(1).Origin.Z
    
    World(1).Timer = World(1).Timer + 1
    If World(1).Timer = 8 And World(1).Speed <> 0 Then
        World(1).Timer = 0
        wake = GetWake: If wake = 0 Then Exit Sub
        Engine.SetWorldToModel wake, 8, World()
        World(wake).Origin.X = World(1).Origin.X
        World(wake).Origin.Z = World(1).Origin.Z
        World(wake).Angle.Y = World(1).Angle.Y
        World(wake).Colour = vbBlue
    End If
    
End Sub


Private Sub MoveFlags()
    For N = 3 To 5: World(5).Morph(N).Angle.Y = (Rnd * 20) - 10
    World(5).Morph(N).Angle.Z = (Rnd * 4) - 2:   Next
    For N = 7 To 9: World(5).Morph(N).Angle.Y = (Rnd * 20) - 10
    World(5).Morph(N).Angle.Z = (Rnd * 4) - 2:   Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quit = True
End Sub

Private Sub LapPass_Timer()
    Check(Int(Rnd * 6)).FillColor = vbGreen
    Check(Int(Rnd * 6)).FillColor = vbBlue
    Check(Int(Rnd * 6)).FillColor = vbRed
    Check(Int(Rnd * 6)).FillColor = vbYellow
    LapPass.Tag = LapPass.Tag - 1
    If LapPass.Tag = 0 Then
        LapPass.Interval = 0
        For nn = 0 To 5
            Check(nn).FillStyle = 1
            Check(nn).FillColor = vbRed
        Next nn
        Check(0).FillStyle = 0
    End If
End Sub

Private Sub Timer1_Timer()
    Me.Caption = Frames
    Frames = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    view.Height = Me.ScaleHeight
    view.Width = Me.ScaleWidth
    Engine.Xf = view.ScaleWidth / 2
    Engine.Yf = view.ScaleHeight / 2
End Sub

Private Function Dist(X1, Y1, X2, Y2) As Single
    Dist = Sqr(((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2))
End Function

Private Function GetAngle(X, Y)
    Dim ang As Integer
    If Y = 0 And X > 0 Then
        ang = 90
    ElseIf Y > 0 And X = 0 Then
        ang = 180
    ElseIf Y = 0 And X < 0 Then
        ang = 270
    ElseIf X = 0 And Y = 0 Then
        ang = 0
    Else
        ang = Abs(Atn(X / Y) * Pie)
        If Y > 0 And X < 0 Then
            ang = ang + 180
        ElseIf Y > 0 And X > 0 Then
            ang = 90 + (90 - ang)
        ElseIf Y < 0 And X < 0 Then
            ang = 90 + (90 - ang) + 180
        End If
    End If
    If ang > 180 Then ang = -(360 - ang)
    GetAngle = -ang
End Function

