VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Create1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Sailcut8   Creating a new sail"
   ClientHeight    =   7260
   ClientLeft      =   5475
   ClientTop       =   2175
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "create1.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form3"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7260
   ScaleWidth      =   9795
   Begin VB.VScrollBar YardAScroll 
      Height          =   1575
      LargeChange     =   10
      Left            =   0
      Max             =   0
      Min             =   89
      TabIndex        =   32
      Top             =   360
      Value           =   75
      Width           =   255
   End
   Begin VB.VScrollBar LoLeechScroll 
      Height          =   1575
      LargeChange     =   50
      Left            =   5640
      Max             =   100
      Min             =   2000
      SmallChange     =   5
      TabIndex        =   31
      Top             =   1800
      Value           =   1000
      Width           =   255
   End
   Begin VB.VScrollBar UpLuffScroll 
      Height          =   1815
      LargeChange     =   50
      Left            =   0
      Max             =   50
      Min             =   1000
      SmallChange     =   5
      TabIndex        =   30
      Top             =   2040
      Value           =   150
      Width           =   255
   End
   Begin VB.HScrollBar NBpanelScroll 
      Height          =   252
      LargeChange     =   4
      Left            =   8160
      Max             =   32
      TabIndex        =   16
      Top             =   4800
      Value           =   4
      Width           =   1332
   End
   Begin VB.HScrollBar NHpanelScroll 
      Height          =   255
      LargeChange     =   2
      Left            =   8160
      Max             =   9
      Min             =   1
      TabIndex        =   21
      Top             =   4440
      Value           =   2
      Width           =   1335
   End
   Begin VB.HScrollBar LYardScroll 
      Height          =   255
      LargeChange     =   100
      Left            =   360
      Max             =   10000
      Min             =   100
      SmallChange     =   10
      TabIndex        =   4
      Top             =   120
      Value           =   2500
      Width           =   2535
   End
   Begin VB.VScrollBar FootAScroll 
      Height          =   1815
      LargeChange     =   5
      Left            =   5640
      Max             =   0
      Min             =   40
      TabIndex        =   6
      Top             =   4200
      Width           =   255
   End
   Begin VB.VScrollBar LoLuffScroll 
      Height          =   2055
      LargeChange     =   50
      Left            =   0
      Max             =   100
      Min             =   2000
      SmallChange     =   5
      TabIndex        =   5
      Top             =   3960
      Value           =   700
      Width           =   255
   End
   Begin VB.HScrollBar LBattenScroll 
      Height          =   255
      LargeChange     =   100
      Left            =   360
      Max             =   10000
      Min             =   100
      SmallChange     =   10
      TabIndex        =   7
      Top             =   6000
      Value           =   2800
      Width           =   2415
   End
   Begin VB.HScrollBar DepthScroll 
      Height          =   252
      LargeChange     =   5
      Left            =   8160
      Max             =   12
      TabIndex        =   24
      Top             =   2880
      Value           =   6
      Width           =   1332
   End
   Begin VB.HScrollBar TwistScroll 
      Height          =   252
      LargeChange     =   5
      Left            =   8160
      Max             =   24
      TabIndex        =   23
      Top             =   4080
      Value           =   12
      Width           =   1332
   End
   Begin VB.HScrollBar RPdepthScroll 
      Height          =   240
      LargeChange     =   5
      Left            =   6480
      Max             =   96
      Min             =   25
      TabIndex        =   22
      Top             =   3600
      Value           =   43
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   5160
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.CommandButton Cancel 
      Appearance      =   0  'Flat
      Caption         =   "Quit  Create1"
      Height          =   372
      Left            =   6600
      TabIndex        =   3
      Top             =   6240
      Width           =   2532
   End
   Begin VB.CommandButton Develop 
      Appearance      =   0  'Flat
      Caption         =   "develop"
      Height          =   375
      Left            =   6600
      TabIndex        =   13
      Top             =   5760
      Width           =   2535
   End
   Begin VB.TextBox SailName 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   285
      Left            =   8160
      TabIndex        =   14
      Text            =   "NEW"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   240
      ScaleHeight     =   5625
      ScaleWidth      =   5385
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label8(1)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   37
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Yard L.  L8(0)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   36
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Leech L6(1)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   17
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lo Leech W L6(0)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   35
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "YardA L5(1)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   34
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Yard angle  L5(0)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   33
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Up Luff L  L1(0)"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   29
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Luff L1(1)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   28
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Surface L7(1)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   12
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Surface L7(0)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   25
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label NBpanelLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "NB panel "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   26
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label NHpanelLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "NH panel "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   27
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Foot A L4(1)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "batten L3(1)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   19
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Luff L2(1)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label RPdepthLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "depth RP"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      TabIndex        =   20
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label TwistLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Twist  "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   18
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label sailLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sail  L20"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label DepthLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Depth "
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Foot angle  L4(0)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Batten L.  L3(0)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   8
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lo Luff L  L2(0)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Menu fmenu 
      Caption         =   "&File       "
      Begin VB.Menu fnew 
         Caption         =   "&New"
      End
      Begin VB.Menu fsep1 
         Caption         =   "-"
      End
      Begin VB.Menu fopen 
         Caption         =   "&Open"
      End
      Begin VB.Menu fsep2 
         Caption         =   "-"
      End
      Begin VB.Menu fprint 
         Caption         =   "&Print"
         Begin VB.Menu fprint_data 
            Caption         =   "Data"
         End
         Begin VB.Menu fprint_XYZ 
            Caption         =   "Panels X-Y"
         End
      End
      Begin VB.Menu fsep3 
         Caption         =   "-"
      End
      Begin VB.Menu fsave 
         Caption         =   "&Save"
         Begin VB.Menu fsave_data 
            Caption         =   "Data"
         End
         Begin VB.Menu fSail_DXF 
            Caption         =   "Sail DXF"
         End
         Begin VB.Menu fsave_DXF 
            Caption         =   "Panels DXF"
         End
         Begin VB.Menu fsave_XYZ 
            Caption         =   "Panels X-Y"
         End
         Begin VB.Menu fsave_vrml1 
            Caption         =   "VRML1.0"
         End
      End
      Begin VB.Menu fsep4 
         Caption         =   "-"
      End
      Begin VB.Menu fquit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu clothMenu 
      Caption         =   "&Cloth    "
      Begin VB.Menu clothOpen 
         Caption         =   "Open"
      End
   End
End
Attribute VB_Name = "Create1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  ' 21 mars 2004
 
Private Sub Cancel_Click()
    
    Menu1.Show

End Sub '--------------------------------------------------

Private Sub chkBatten()

    If LBattenScroll.Value < LoLuffScroll.Value Then
        LBattenScroll.Value = LoLuffScroll.Value
        Label3(1).BackColor = &HFFFF&
        '-----
    Else
        Label3(1).BackColor = &H80000005
        '-----
    End If

End Sub '--------------------------------------------------

Private Sub clean_click()

    Picture1.Cls
    'Call compute
    Call echelle(xo, yo, ks)
    Call dessinPlan
    Call dessinAR

End Sub ' clean -------------------------------------------

Private Sub clothOpen_Click()
    
    Cloth1.Show

End Sub  '---------------------------------------------

Private Sub compute()
    Dim n%, i%, W%
    Dim a!, b!, c!
    Dim alfa!, beta!, gama!, delta!
    Dim dir1!, dir2!
    Dim xi!
    Dim h!, dl!, l!
    Erase px, py, pz
    '-----
    n = 0
    px(0, 0) = 0
    py(0, 0) = 0
    pz(0, 0) = 0
    '----- reset couleurs
    NBpanelLabel.BackColor = &H80000005
    NHpanelLabel.BackColor = &H80000005
    Label1(0).BackColor = &H80000005
    Label1(1).BackColor = &H80000005
    Label2(0).BackColor = &H80000005
    Label2(1).BackColor = &H80000005
    Label6(0).BackColor = &H80000005
    Label6(1).BackColor = &H80000005
    Label8(0).BackColor = &H80000005
    Label8(1).BackColor = &H80000005
    '----- baton bordure
    For i = 1 To 20
        px(0, i) = px(0, 0) + LBatten * i / 20 * Cos(AFoot / RAD)
        py(0, i) = py(0, 0) + LBatten * i / 20 * Sin(AFoot / RAD)
        pz(0, i) = 0
    Next i
    
    '----- batons panneaux bas
    W = LoLeech

    For n = 1 To nBpanel
        alfa = Atn((py(2 * (n - 1), 20) - py(2 * (n - 1), 0)) / (px(2 * (n - 1), 20) - px(2 * (n - 1), 0)))
        
        px(2 * n, 0) = 0
        py(2 * n, 0) = py(2 * (n - 1), 0) + LoLuff
        pz(2 * n, 0) = 0
        
        beta = Atn((W - LoLuff * Cos(alfa)) / LBatten)
        '----- check angle
        If (alfa + beta) >= 1.3 Then 'angle max
            beta = 1.3 - alfa
            NBpanelLabel.BackColor = &HFFFF&
        End If
        
        If (alfa + beta) >= AYard / RAD Then
            beta = AYard / RAD - alfa
            NBpanelLabel.BackColor = &HFF&
        End If
        '-----
        For i = 1 To 20
            px(2 * n, i) = px(2 * n, 0) + LBatten * i / 20 * Cos(alfa + beta)
            py(2 * n, i) = py(2 * n, 0) + LBatten * i / 20 * Sin(alfa + beta)
            pz(2 * n, i) = 0
        Next i

        AHead = alfa + beta
        '----- check pour cassure dans la chute en bas
        If n > 2 Then
          dir1 = directionXY(px(2 * (n - 1), 20), py(2 * (n - 1), 20), px(2 * n, 20), py(2 * n, 20))
          dir2 = directionXY(px(2 * (n - 2), 20), py(2 * (n - 2), 20), px(2 * (n - 1), 20), py(2 * (n - 1), 20))
          If dir1 <= dir2 Then
            If NBpanelLabel.BackColor = &HFF& Then
                Label2(0).BackColor = &HFF&
                Label2(1).BackColor = &HFF&
                Label6(0).BackColor = &HFF&
                Label6(1).BackColor = &HFF&
              Else
                NBpanelLabel.BackColor = &HFFFF&
                Label2(0).BackColor = &HFFFF&
                Label2(1).BackColor = &HFFFF&
                Label6(0).BackColor = &HFFFF&
                Label6(1).BackColor = &HFFFF&
            End If
          End If
        End If
        '-----
    Next n

    '----- lignes intermédiaires basses avec profile
    'Call CubicP(RPdepth, a, b, c)  ' profile des sections
    For n = 1 To nBpanel
        For i = 0 To 20
          px(2 * n - 1, i) = (px(2 * n - 2, i) + px(2 * n, i)) / 2
          py(2 * n - 1, i) = (py(2 * n - 2, i) + py(2 * n, i)) / 2
          pz(2 * n - 1, i) = (pz(2 * n - 2, i) + pz(2 * n, i)) / 2
          xi = (i / 20)
          If n < nBpanel Then
            pz(2 * n - 1, i) = pz(2 * n - 1, i) + Mdepth * LBatten * profileP(RPdepth, xi)
            Else ' half depth
            pz(2 * n - 1, i) = pz(2 * n - 1, i) + Mdepth / 2 * LBatten * profileP(RPdepth, xi)
          End If
        Next i
    Next n

    '----- batons panneaux de tête
    alfa = Atn((py(2 * (n - 1), 20) - py(2 * (n - 1), 0)) / (px(2 * (n - 1), 20) - px(2 * (n - 1), 0)))
    delta = (AYard / RAD - alfa) / nHpanel
    
    If beta < 0 Then
        beta = 0
        AYard = alfa * RAD
        YardAScroll.Value = AYard
    End If

    dl = (LYard - LBatten)

    For n = (nBpanel + 1) To (nBpanel + nHpanel)
        
        alfa = alfa + delta
        l = LBatten + dl * ((n - nBpanel) / nHpanel) ^ 1.5

        px(2 * n, 0) = 0
        py(2 * n, 0) = py(2 * (n - 1), 0) + UpLuff
        pz(2 * n, 0) = 0
        
        For i = 1 To 20
            px(2 * n, i) = px(2 * n, 0) + l * i / 20 * Cos(alfa)
            py(2 * n, i) = py(2 * n, 0) + l * i / 20 * Sin(alfa)
            pz(2 * n, i) = 0
        Next i
        '----- check pour cassure dans la chute en tete
        If nBpanel > 1 Then
        dir1 = directionXY(px(2 * (n - 1), 20), py(2 * (n - 1), 20), px(2 * n, 20), py(2 * n, 20))
        dir2 = directionXY(px(2 * (n - 2), 20), py(2 * (n - 2), 20), px(2 * (n - 1), 20), py(2 * (n - 1), 20))
        If dir1 <= dir2 Then
            NHpanelLabel.BackColor = &HFFFF&
            Label1(0).BackColor = &HFFFF&
            Label1(1).BackColor = &HFFFF&
            Label8(0).BackColor = &HFFFF&
            Label8(1).BackColor = &HFFFF&
        End If
        End If
        '-----
    Next n

    Xpeak = px(2 * (nBpanel + nHpanel), 20)
    Ypeak = py(2 * (nBpanel + nHpanel), 20)
    Zpeak = pz(2 * (nBpanel + nHpanel), 20)

    '----- ligne intermédiaire de tête
    For n = (nBpanel + 1) To (nBpanel + nHpanel)
        For i = 0 To 20
            px(2 * n - 1, i) = (px(2 * n - 2, i) + px(2 * n, i)) / 2
            py(2 * n - 1, i) = (py(2 * n - 2, i) + py(2 * n, i)) / 2
            pz(2 * n - 1, i) = (pz(2 * n - 2, i) + pz(2 * n, i)) / 2
        Next i
    Next n

    '----- surface
    Surface = 0
    For n = 1 To 2 * (nBpanel + nHpanel)
        For i = 1 To 20
            a = distance2D(px(n - 1, i - 1), py(n - 1, i - 1), px(n, i - 1), py(n, i - 1))
            b = distance2D(px(n - 1, i - 1), py(n - 1, i - 1), px(n, i), py(n, i))
            c = distance2D(px(n, i - 1), py(n, i - 1), px(n, i), py(n, i))
            Call Triangle(a, b, c, alfa, beta, gama)
            Surface = Surface + Sin(gama) * (a * b) / 2
            '-----
            a = distance2D(px(n - 1, i), py(n - 1, i), px(n, i), py(n, i))
            c = distance2D(px(n - 1, i - 1), py(n - 1, i - 1), px(n - 1, i), py(n - 1, i))
            Call Triangle(a, b, c, alfa, beta, gama)
            Surface = Surface + Sin(gama) * (a * b) / 2
        Next i
    Next n
    Surface = Surface / 1000000
    
    '----- adding twist
    For n = 0 To 2 * (nBpanel + nHpanel)
        h = (py(n, 20) - py(0, 20)) / (Ypeak - py(0, 20) + 0.001)
        For i = 0 To 20
            Rot2D 0, 0, px(n, i), pz(n, i), h * Atwist
        Next i
    Next n
    
    '-----
End Sub ' compute ----------------------------------------

Private Sub defaut()
    ' valeurs par défaut de la voile
    
    Sail = "NEW_JUNK"
    SailName.Text = Sail
    Sailname_lostfocus

    UpLuffScroll.Value = 150   'mm
    LoLuffScroll.Value = 700   'mm
    LoLeechScroll.Value = 1000 'mm
    LBattenScroll.Value = 2800 'mm
    LYardScroll.Value = 2500   'mm

    FootAScroll.Value = 0  'deg
    YardAScroll.Value = 75  'deg

    DepthScroll.Value = 6    '%
    RPdepthScroll.Value = 43 '%
    
    TwistScroll.Value = 12  'degré

    NHpanelScroll.Value = 2
    NBpanelScroll.Value = 4

    ClothW = 900 'mm
    SeamW = 50   'mm
    SeamT = 15   'mm

    genre1 = "JunkSail"
    genre2 = "default_2"
    genre3 = "default-3"

End Sub ' defaut -----------------------------------------

Private Sub DepthScroll_Change()

    Mdepth = DepthScroll.Value / 100
    
    form_activate

End Sub ' Sub DepthScroll_Change() ------------------------

Private Sub dessinDev()
    Dim couleur As Long
    Dim n%, i%
    Dim xb!, yb!
   
'----- dessin developpement
For n = 1 To nBpanel + nHpanel
    If n > nBpanel Then  'partie haute
        couleur = RGB(200, 0, 200)
    Else
        couleur = RGB(200, 0, 0)
    End If

    Picture1.DrawWidth = 1
    xb = (Picture1.Width - 300) - (1.1 * LBatten * ks)
    'xb = 1.3 * LBatten * ks
    
    ' calculer origine de dessin du panneau
    yb = -200
    For i = 1 To n
        yb = yb + Abs(pmy(i - 1, 20) - ply(i - 1, 20)) * ks + 80
    Next i
    
    yb = yo + yb

    ' dessiner bord inférieur
    Picture1.PSet (xb + plx(n, 0) * ks, yb + ply(n, 0) * ks), couleur
    For i = 1 To 20
        Picture1.Line -(xb + plx(n, i) * ks, yb + ply(n, i) * ks), couleur
    Next i

    Picture1.PSet (xb + plx(n, 0) * ks, yb + ply(n, 0) * ks), couleur
        Picture1.Line -(xb + pcx(n, 0) * ks, yb + pcy(n, 0) * ks), couleur
    
    For i = 0 To 20
        Picture1.Line -(xb + pmx(n, i) * ks, yb + pmy(n, i) * ks), couleur
    Next i
        
        Picture1.Line -(xb + pcx(n, 20) * ks, yb + pcy(n, 20) * ks), couleur
        Picture1.Line -(xb + plx(n, 20) * ks, yb + ply(n, 20) * ks), couleur
    
    For i = 0 To 20
        Picture1.PSet (xb + plx(n, i) * ks, yb + ply(n, i) * ks)
        Picture1.Line -(xb + pcx(n, i) * ks, yb + pcy(n, i) * ks), RGB(0, 127, 0)
        Picture1.Line -(xb + pmx(n, i) * ks, yb + pmy(n, i) * ks), RGB(65, 127, 0)
    Next i
    '-----
Next n
'-----
Picture1.DrawWidth = 1

End Sub ' dessinDev ---------------------------------------

Private Sub dessinAR()
    ' dessine le plan de voile
    Dim couleur As Long
    Dim n%, i%
    Dim xa!, ya!
    '--- nouvelle origine
    xa = xo + 1.15 * LBatten * ks
    ya = yo
'----- dessin en vue arriere voile
    ' bordure
    Picture1.DrawWidth = 2
    Picture1.PSet (xa + pz(0, 0) * ks, ya + py(0, 0) * ks)
    For i = 0 To 20
        Picture1.Line -(xa + pz(0, i) * ks, ya + py(0, i) * ks)
    Next i

'----- batons
For n = 2 To 2 * (nBpanel + nHpanel) Step 2
    If n > 2 * nBpanel Then
        couleur = RGB(200, 0, 200) 'tete
      ElseIf n = 2 * nBpanel Then
        couleur = RGB(0, 0, 0) 'transition bas-tete
      Else
        couleur = RGB(200, 0, 0) 'bas
    End If

    Picture1.DrawWidth = 1
    Picture1.PSet (xa + pz(n - 2, 0) * ks, ya + py(n - 2, 0) * ks)
    Picture1.Line -(xa + pz(n, 0) * ks, ya + py(n, 0) * ks)

    Picture1.DrawWidth = 2
    For i = 1 To 20
        Picture1.Line -(xa + pz(n, i) * ks, ya + py(n, i) * ks), couleur
    Next i
    
    Picture1.DrawWidth = 1
    Picture1.Line -(xa + pz(n - 2, 20) * ks, ya + py(n - 2, 20) * ks)
Next n

'----- lignes intermédiaires
For n = 1 To 2 * (nBpanel + nHpanel) Step 2
    Picture1.DrawWidth = 1
    Picture1.PSet (xa + pz(n, 0) * ks, ya + py(n, 0) * ks)
    For i = 1 To 20
        Picture1.Line -(xa + pz(n, i) * ks, ya + py(n, i) * ks), RGB(0, 200, 250)
    Next i
Next n
'-----
Picture1.DrawWidth = 1

End Sub ' dessinAR -----------------

Private Sub dessinPlan()
    ' dessine le plan de voile
    Dim couleur As Long
    Dim n%, i%
    Dim X!, Y!
    
'----- dessin profile
    Picture1.DrawWidth = 1
    Picture1.PSet (xo, 50)
    For i = 0 To 50
        X = i / 50
        Y = Mdepth * LBatten * profileP(RPdepth, X)
        Picture1.Line -(xo + (LBatten * X) * ks, 50 + Y * ks), RGB(0, 127, 0)
    Next i
        Picture1.Line -(xo, 50), RGB(0, 0, 0)
        
'----- dessin plan de voile
    ' bordure
    Picture1.DrawWidth = 2
    Picture1.PSet (xo + px(0, 0) * ks, yo + py(0, 0) * ks)
    For i = 0 To 20
        Picture1.Line -(xo + px(0, i) * ks, yo + py(0, i) * ks)
    Next i

'----- batons
For n = 2 To 2 * (nBpanel + nHpanel) Step 2
    If n > 2 * nBpanel Then
        couleur = RGB(200, 0, 200) 'tete
      ElseIf n = 2 * nBpanel Then
        couleur = RGB(0, 0, 0) 'transition bas-tete
      Else
        couleur = RGB(200, 0, 0) 'bas
    End If

    Picture1.DrawWidth = 1
    Picture1.PSet (xo + px(n - 2, 0) * ks, yo + py(n - 2, 0) * ks)
    Picture1.Line -(xo + px(n, 0) * ks, yo + py(n, 0) * ks)

    Picture1.DrawWidth = 2
    For i = 1 To 20
        Picture1.Line -(xo + px(n, i) * ks, yo + py(n, i) * ks), couleur
    Next i
    
    Picture1.DrawWidth = 1
    Picture1.Line -(xo + px(n - 2, 20) * ks, yo + py(n - 2, 20) * ks)
Next n

'----- lignes intermediaires
For n = 1 To 2 * (nBpanel + nHpanel) Step 2
    Picture1.DrawWidth = 1
    Picture1.PSet (xo + px(n, 0) * ks, yo + py(n, 0) * ks)
    For i = 1 To 20
        Picture1.Line -(xo + px(n, i) * ks, yo + py(n, i) * ks), RGB(0, 200, 250)
    Next i
Next n
'-----
Picture1.DrawWidth = 1

End Sub '--------------------------------------------------

Private Sub Develop_Click()
    ' developpement

    MousePointer = 11 'change pointer to hour glass
    
    Call dessinDev
        
    MousePointer = 0

End Sub '------------------------------------------------

Private Sub develop_hi(pan)
    '--- développer panneaux haut
    
    Dim p%, i%, lo%, hi%
    Dim alfa!, beta!, gama!
    Dim a!, b!, c!, r!
    '-----
    p = 2 * nBpanel + 2 * pan
    lo = 2 * nBpanel + 2 * pan - 2
    hi = 2 * nBpanel + 2 * pan

    plx(p, 0) = 0
    ply(p, 0) = 0

    alfa = Atn((py(lo, 20) - py(lo, 0)) / (px(lo, 20) - px(lo, 0)))

    For i = 1 To 20   'bord inférieur
        r = Sqr((px(lo, i) - px(lo, 0)) ^ 2 + (py(lo, i) - py(lo, 0)) ^ 2)
        beta = directionXY(px(lo, 0), py(lo, 0), px(lo, i), py(lo, i))
        plx(p, i) = r * Cos(beta - alfa)
        ply(p, i) = r * Sin(beta - alfa)
    Next i

    For i = 0 To 20   'bord supérieur
        r = Sqr((px(hi, i) - px(lo, 0)) ^ 2 + (py(hi, i) - py(lo, 0)) ^ 2)
        beta = directionXY(px(lo, 0), py(lo, 0), px(hi, i), py(hi, i))
        pmx(p, i) = r * Cos(beta - alfa)
        pmy(p, i) = r * Sin(beta - alfa)
    Next i
    '-----

End Sub ' develop_hi -------------------------------------

Private Sub develop2(pan%)
    'développer panneaux bas
    
    Dim p%, i%, n%
    Dim alfa!, beta!, gama!
    Dim a!, b!, c!
    Dim h!, v!

    n = 2 * pan
    
    pcx(pan, 0) = 0
    pcy(pan, 0) = 0
    
    '--- points courants haut
    
    For i = 1 To 20
        b = distance3D(px(n - 1, i - 1), py(n - 1, i - 1), pz(n - 1, i - 1), px(n - 1, i), py(n - 1, i), pz(n - 1, i))
        a = distance3D(px(n - 1, i - 1), py(n - 1, i - 1), pz(n - 1, i - 1), px(n, i), py(n, i), pz(n, i))
        c = distance3D(px(n - 1, i), py(n - 1, i), pz(n - 1, i), px(n, i), py(n, i), pz(n, i))
        'Debug.Print a, b, c
        Call Triangle(a, b, c, alfa, beta, gama)
        'Debug.Print a, b, c, gama
        pcx(pan, i) = pcx(pan, i - 1) + b
        pcy(pan, i) = 0
        pmx(pan, i) = pcx(pan, i - 1) + a * Cos(gama)
        pmy(pan, i) = pcy(pan, i - 1) + a * Sin(gama)
    Next i
    
    '--- premier triangle haut
    b = distance3D(px(n - 1, 0), py(n - 1, 0), pz(n - 1, 0), px(n, 1), py(n, 1), pz(n, 1))
    a = distance3D(px(n - 1, 0), py(n - 1, 0), pz(n - 1, 0), px(n, 0), py(n, 0), pz(n, 0))
    c = distance3D(px(n, 0), py(n, 0), pz(n, 0), px(n, 1), py(n, 1), pz(n, 1))
    Call Triangle(a, b, c, alfa, beta, gama)
    alfa = directionXY(pcx(pan, 0), pcy(pan, 0), pmx(pan, 1), pmy(pan, 1))
    pmx(pan, 0) = pcx(pan, 0) + a * Cos(alfa + gama)
    pmy(pan, 0) = pcy(pan, 0) + a * Sin(alfa + gama)
  
    '----- points courants bas
    For i = 1 To 20
        b = pcx(pan, i) - pcx(pan, i - 1)
        a = distance3D(px(n - 1, i - 1), py(n - 1, i - 1), pz(n - 1, i - 1), px(n - 2, i), py(n - 2, i), pz(n - 2, i))
        c = distance3D(px(n - 1, i), py(n - 1, i), pz(n - 1, i), px(n - 2, i), py(n - 2, i), pz(n - 2, i))
        'Debug.Print a, b, c
        Call Triangle(a, b, c, alfa, beta, gama)
        'Debug.Print a, b, c, gama
        plx(pan, i) = pcx(pan, i - 1) + a * Cos(-gama)
        ply(pan, i) = pcy(pan, i - 1) + a * Sin(-gama)
    Next i
    
    '--- premier triangle bas
    b = distance3D(px(n - 1, 0), py(n - 1, 0), pz(n - 1, 0), px(n - 2, 1), py(n - 2, 1), pz(n - 2, 1))
    a = distance3D(px(n - 1, 0), py(n - 1, 0), pz(n - 1, 0), px(n - 2, 0), py(n - 2, 0), pz(n - 2, 0))
    c = distance3D(px(n - 2, 0), py(n - 2, 0), pz(n - 2, 0), px(n - 2, 1), py(n - 2, 1), pz(n - 2, 1))
    Call Triangle(a, b, c, alfa, beta, gama)
    alfa = directionXY(pcx(pan, 0), pcy(pan, 0), plx(pan, 1), ply(pan, 1))
    plx(pan, 0) = pcx(pan, 0) + a * Cos(alfa - gama)
    ply(pan, 0) = pcy(pan, 0) + a * Sin(alfa - gama)
  
    '--- recadrage vertical
    v = 0
    For i = 0 To 20
        If ply(pan, i) < v Then
            v = ply(pan, i)
        End If
    Next i

    For i = 0 To 20
        pcy(pan, i) = pcy(pan, i) - v
        ply(pan, i) = ply(pan, i) - v
        pmy(pan, i) = pmy(pan, i) - v
    Next i
    
    '--- recadrage horizontal
    If plx(pan, 0) < h Then h = plx(pan, 0)
    If pcx(pan, 0) < h Then h = pcx(pan, 0)
    If pmx(pan, 0) < h Then h = pmx(pan, 0)

    For i = 0 To 20
        pcx(pan, i) = pcx(pan, i) - h
        plx(pan, i) = plx(pan, i) - h
        pmx(pan, i) = pmx(pan, i) - h
    Next i
    '---

End Sub ' develop2 -------------------------------------

Private Sub devjunk()
    Dim n%, p%, i%
    Dim alfa!, beta!, gama!
    Dim a!, b!, c!, r!
    Erase plx, ply, pmx, pmy

'----- copier les points xyz
For n = 1 To 2 * (nBpanel + nHpanel)
    For i = 0 To 20
        plx(n, i) = px(n - 1, i) - px(n - 1, 0)
        ply(n, i) = py(n - 1, i) - py(n - 1, 0)
        pmx(n, i) = px(n, i) - px(n - 1, 0)
        pmy(n, i) = py(n, i) - py(n - 1, 0)
    Next i
Next n

'----- recadrage
For n = 1 To 2 * (nBpanel + nHpanel)
    p = n Mod 2
    '-----
    Select Case p

    Case 1  ' panneau au dessus d'un baton

        alfa = Atn((ply(n, 20) - ply(n, 0)) / (plx(n, 20) - plx(n, 0)))
    
        For i = 1 To 20   'bord inferieur
            r = Sqr((plx(n, i) - plx(n, 0)) ^ 2 + (ply(n, i) - ply(n, 0)) ^ 2)
            beta = directionXY(plx(n, 0), ply(n, 0), plx(n, i), ply(n, i))
            plx(n, i) = r * Cos(beta - alfa)
            ply(n, i) = r * Sin(beta - alfa)
        Next i
    
        For i = 0 To 20   'bord superieur
            r = Sqr((pmx(n, i) - plx(n, 0)) ^ 2 + (pmy(n, i) - ply(n, 0)) ^ 2)
            beta = directionXY(plx(n, 0), ply(n, 0), pmx(n, i), pmy(n, i))
            pmx(n, i) = r * Cos(beta - alfa)
            pmy(n, i) = r * Sin(beta - alfa)
        Next i
    '-----

    Case 0  'panneau au dessous d'un baton

        alfa = Atn((pmy(n, 20) - pmy(n, 0)) / (pmx(n, 20) - pmx(n, 0)))
    
        For i = 1 To 20   'bord inferieur
            r = Sqr((plx(n, i) - plx(n, 0)) ^ 2 + (ply(n, i) - ply(n, 0)) ^ 2)
            beta = directionXY(plx(n, 0), ply(n, 0), plx(n, i), ply(n, i))
            plx(n, i) = r * Cos(beta - alfa)
            ply(n, i) = r * Sin(beta - alfa)
        Next i
    
        For i = 0 To 20   'bord superieur
            r = Sqr((pmx(n, i) - plx(n, 0)) ^ 2 + (pmy(n, i) - ply(n, 0)) ^ 2)
            beta = directionXY(plx(n, 0), ply(n, 0), pmx(n, i), pmy(n, i))
            pmx(n, i) = r * Cos(beta - alfa)
            pmy(n, i) = r * Sin(beta - alfa)
        Next i
        '--- recadrage vertical
        r = 0
        For i = 0 To 20
            If ply(n, i) < r Then
                r = ply(n, i)
            End If
        Next i
    
        For i = 0 To 20
            ply(n, i) = ply(n, i) - r
            pmy(n, i) = pmy(n, i) - r
        Next i
        
        '--- recadrage horizontal
        r = pmx(n, 0)
        For i = 0 To 20
            plx(n, i) = plx(n, i) - r
            pmx(n, i) = pmx(n, i) - r
        Next i

    '-----
    End Select
    '-----
Next n

End Sub '-----------------------------------------

Private Sub echelle(xo, yo, ks)
    ' compute scale factor - calcul facteur d'échelle
    ' return the origin coordinates  (xo,yo)
    ' and screen scale factor  (ks)

    Dim sw!, sh!, W!, h!, h1!, h2!, k1!, k3!
    
    sw = Picture1.Width
    sh = Picture1.Height
    
    Picture1.Scale (0, sh)-(sw, 0)
    
    xo = 200
    yo = 400
    '--- largeur
    k1 = (sw - 300) / (100 + LBatten)
    '--- hauteur
    h1 = nBpanel * (LoLuff + 50) + nHpanel * (UpLuff + 50) + LYard + 50
    h2 = (nHpanel + nBpanel) * (LoLeech * 1.13) - LBatten * Sin(AFoot / (2 * RAD))
    If h1 > h2 Then
        h = h1
    Else
        h = h2
    End If

    k3 = (sh - 300) / (100 + h)
    '---
    If k1 < k3 Then
        ks = k1
      Else
        ks = k3
    End If
    '---

End Sub '-----------------------------------------------

Private Sub fnew_Click()
    
    defaut

End Sub  '---------------------------------------------

Private Sub FootAScroll_Change()
    
    AFoot = FootAScroll.Value

    form_activate

End Sub '----------------------------------------------

Private Sub fopen_Click()
    ' load a file - charger un fichier
    Dim fichier$
'---
CMDialog1.CancelError = True

On Error GoTo errhandler1
    CMDialog1.CancelError = True
    CMDialog1.Filter = "SAILCUT6 files|*.sc8"
    CMDialog1.FileName = Sail + ".sc8"
    CMDialog1.Action = 1
    Sail$ = UCase$(CMDialog1.FileTitle)
    Sail$ = Left$(Sail$, (Len(Sail$) - 4))
    fichier$ = CMDialog1.FileName
    
    Call F_Lire(fichier$)

    SailName.BackColor = &H80000005
    SailName.Text = UCase(Sail)

    UpLuffScroll.Value = UpLuffScrl
    LoLuffScroll.Value = LoLuffScrl
    LoLeechScroll.Value = LoLeechScrl
    LBattenScroll.Value = LBattenScrl
    LYardScroll.Value = LyardScrl
    FootAScroll.Value = FootAScrl
    YardAScroll.Value = YardAScrl

    DepthScroll.Value = Mdepth * 100
    RPdepthScroll.Value = RPdepth * 100
    TwistScroll.Value = twistScrl
    
    NHpanelScroll.Value = nHpanel
    NBpanelScroll.Value = nBpanel
    '-----
    
    form_activate
    
    '-----
    Exit Sub

errhandler1:
    Close
    If Err <> 32755 Then
        MsgBox Error(Err), 0, "Create fopen"
    End If
        Create1.MousePointer = 0
    Exit Sub

End Sub ' fopen ------------------------------------------

Private Sub form_activate()
    ' Create1 Activate

    Create1.MousePointer = 11  ' pointeur hourglass
    
    redraw  ' inclue le calcul d'échelle

    SailName.Text = UCase(Sail)

    takeData

'----- afficher dimensions
    Label1(1).Caption = Str(Int(UpLuff)) + " mm"
    Label2(1).Caption = Str(Int(LoLuff)) + " mm"
    Label3(1).Caption = Str(Int(LBatten)) + " mm"
    Label4(1).Caption = Str(Int(AFoot)) + " deg"
    Label5(1).Caption = Str(Int(AYard)) + " deg"
    Label6(1).Caption = Str(Int(LoLeech)) + " mm"
    Label7(1).Caption = Format$(Surface, "0.#0") + " m2"
    Label8(1).Caption = Str(Int(LYard)) + " mm"

'----- afficher textes
Select Case langue
  Case 0 'français

    If Sail$ <> "NEW_JUNK" Then
        Create1.Caption = titre$ + "  Géométrie : modifier la voile : " + Sail$
    Else
        Create1.Caption = titre$ + "  Géométrie d'une nouvelle voile "
    End If
    
    Label1(0) = "Largeur coté guindant panneaux de tête"
    Label2(0) = "Largeur coté guindant panneaux bas"
    Label3(0) = "Longueur latte"
    Label4(0) = "Angle bordure"
    Label5(0) = "Angle yard"
    Label6(0) = "Largeur coté chute panneaux bas"
    Label7(0) = "Surface"
    Label8(0) = "Longueur Yard"

    sailLabel = "Nom fichier voile :"
    Develop.Caption = "Développement"
    Cancel.Caption = "Quitter"
    
    fmenu.Caption = "&Fichier"
    fnew.Caption = "&Nouveau"
    fopen.Caption = "&Ouvrir fichier"
    fprint.Caption = "Im&Primer"
    fprint_data.Caption = "Données"
    fprint_XYZ.Caption = "Panneaux XY"
    fsave.Caption = "&Sauver"
    fsave_data.Caption = "Données"
    fsave_XYZ.Caption = "Panneaux XY"
    fsave_vrml1.Caption = "VRML 1.0"
    fquit.Caption = "&Quitter"

    clothMenu.Caption = "&Tissus"
    clothOpen.Caption = "Ouvrir"
    
    DepthLabel.Caption = "Creux :" + Format$(Mdepth * 100, " #.0") + "%"
    RPdepthLabel.Caption = "Position du creux :" + Str(Int(RPdepth * 100)) + "%"
    TwistLabel.Caption = "Vrillage :" + Str$(twistScrl) + " deg."
    
    NHpanelLabel.Caption = "Panneaux tête :" + Str(nHpanel)
    NBpanelLabel.Caption = "Panneaux bas  :" + Str(nBpanel)
    
  Case 1 'anglais

    If Sail$ <> "NEW_JUNK" Then
        Create1.Caption = titre$ + "  Geometry : modifying the sail : " + Sail$
    Else
        Create1.Caption = titre$ + "  Geometry of a new sail "
    End If

    Label1(0) = "Luff width of upper panels"
    Label2(0) = "Luff width of lower panels"
    Label3(0) = "Batten length"
    Label4(0) = "Foot angle"
    Label5(0) = "Yard angle"
    Label6(0) = "Leech width of lower panels"
    Label7(0) = "Area"
    Label8(0) = "Yard Length"

    sailLabel = "Sail file name :"
    Develop.Caption = "Development"
    Cancel.Caption = "Quit"

    fmenu.Caption = "&File"
    fnew.Caption = "&New"
    fopen.Caption = "&Open file"
    fprint.Caption = "&Print"
    fprint_data.Caption = "Data"
    fprint_XYZ.Caption = "Panels XY"
    fsave.Caption = "&Save"
    fsave_data.Caption = "Data"
    fsave_XYZ.Caption = "Panels XY"
    fsave_vrml1.Caption = "VRML 1.0"
    fquit.Caption = "&Quit"

    clothMenu.Caption = "Clo&Th"
    clothOpen.Caption = "Open"

    DepthLabel.Caption = "Depth:" + Format$(Mdepth * 100, " #.0") + "%"
    RPdepthLabel.Caption = "Depth position :" + Str(Int(RPdepth * 100)) + "%"
    TwistLabel.Caption = "Twist :" + Str$(twistScrl) + " deg."

    NHpanelLabel.Caption = "Head panels   :" + Str(nHpanel)
    NBpanelLabel.Caption = "Bottom panels :" + Str(nBpanel)
        '-----
End Select
'------
    Create1.MousePointer = 0
    Menu1.MousePointer = 0

End Sub ' Form_Activate ----------------------------------

Private Sub Form_Load()
    ' Create1 Form Load
    Width = 0.9 * Screen.Width
    Height = 0.7 * Width
    Top = (Screen.Height - Height) / 2 - 100
    Left = (Screen.Width - Width) / 2 - 100

    fsave_vrml1.Visible = False
    clothMenu.Visible = False

    defaut

End Sub ' Form_Load --------------------------------------

Private Sub Form_Resize()

    Dim p1%, p2%, p3% 'positions of labels
    '-----
    YardAScroll.Left = 10
    LYardScroll.Top = 10
    
    Picture1.Left = YardAScroll.Left + YardAScroll.Width + 10
        Picture1.Top = LYardScroll.Top + LYardScroll.Height + 10
        Picture1.Height = Create1.Height - 1300
        Picture1.Width = Create1.Width - 4150
    
    '--- haut
    LYardScroll.Width = 0.6 * Picture1.Width
        LYardScroll.Left = Picture1.Left + 100
    '--- bas
    LBattenScroll.Width = LYardScroll.Width
        LBattenScroll.Left = LYardScroll.Left
        LBattenScroll.Top = Picture1.Top + Picture1.Height + 10
    
    '--- gauche
    YardAScroll.Height = Picture1.Height / 3 - 100
        YardAScroll.Top = Picture1.Top

    UpLuffScroll.Height = Picture1.Height / 3 - 100
        UpLuffScroll.Top = Picture1.Top + (Picture1.Height) / 3 + 50
        UpLuffScroll.Left = YardAScroll.Left
    
    LoLuffScroll.Height = Picture1.Height / 3 - 100
        LoLuffScroll.Top = Picture1.Top + 2 * (Picture1.Height) / 3 + 100
        LoLuffScroll.Left = YardAScroll.Left
    
    '--- droite
    LoLeechScroll.Height = Picture1.Height / 3 - 100
        LoLeechScroll.Top = Picture1.Top + (Picture1.Height) / 4 + 50
        LoLeechScroll.Left = Picture1.Left + Picture1.Width + 10

    FootAScroll.Height = Picture1.Height / 3 - 100
        FootAScroll.Top = Picture1.Top + 2 * (Picture1.Height) / 3 + 100
        FootAScroll.Left = Picture1.Left + Picture1.Width + 10
    
    Picture1.Scale (-0.02 * Picture1.Width, -0.04 * Picture1.Height)-(1.02 * Picture1.Width, 1.02 * Picture1.Height)
    
    p1 = FootAScroll.Left + FootAScroll.Width + 150
    p2 = p1 + 2000
    p3 = p2 + 1250
    
    Label1(0).Left = p1  'Upper Luff width
        Label1(0).Width = p2 - p1
        Label1(0).Top = 60
        Label1(0).Height = 400
        Label1(1).Left = p2
        Label1(1).Width = 1800
        Label1(1).Top = Label1(0).Top
        Label1(1).Height = 250
    
    Label2(0).Left = p1  'Lower Luff width
        Label2(0).Width = p2 - p1
        Label2(0).Top = Label1(0).Top + Label1(0).Height + 50
        Label2(0).Height = Label1(0).Height
        Label2(1).Left = p2
        Label2(1).Width = Label1(1).Width
        Label2(1).Top = Label2(0).Top
        Label2(1).Height = Label1(0).Height

    Label6(0).Left = p1 'Lower Leech width
        Label6(0).Width = p2 - p1
        Label6(0).Top = Label2(0).Top + Label2(0).Height + 50
        Label6(0).Height = Label1(0).Height
        Label6(1).Left = p2
        Label6(1).Width = Label2(1).Width
        Label6(1).Top = Label6(0).Top
        Label6(1).Height = Label6(0).Height
    
    Label3(0).Left = p1  'Batten length
        Label3(0).Width = p2 - p1
        Label3(0).Top = Label6(0).Top + Label6(0).Height + 50
        Label3(0).Height = 250
        Label3(1).Left = p2
        Label3(1).Width = Label2(1).Width
        Label3(1).Top = Label3(0).Top
        Label3(1).Height = Label3(0).Height

    Label8(0).Left = p1  'Yard length
        Label8(0).Width = p2 - p1
        Label8(0).Top = Label3(0).Top + Label3(0).Height + 50
        Label8(0).Height = 250
        Label8(1).Left = p2
        Label8(1).Width = Label2(1).Width
        Label8(1).Top = Label8(0).Top
        Label8(1).Height = Label8(0).Height

    Label4(0).Left = p1  'batten angle
        Label4(0).Width = p2 - p1
        Label4(0).Top = Label8(0).Top + Label8(0).Height + 50
        Label4(0).Height = 250
        Label4(1).Left = p2
        Label4(1).Width = Label2(1).Width
        Label4(1).Top = Label4(0).Top
        Label4(1).Height = Label4(0).Height

    Label5(0).Left = p1 'yard angle
        Label5(0).Width = p2 - p1
        Label5(0).Top = Label4(0).Top + Label4(0).Height + 50
        Label5(0).Height = 250
        Label5(1).Left = p2
        Label5(1).Width = Label2(1).Width
        Label5(1).Top = Label5(0).Top
        Label5(1).Height = Label5(0).Height
    
    Label7(0).Left = p1 'surface
        Label7(0).Width = p2 - p1
        Label7(0).Top = Label5(0).Top + Label5(0).Height + 50
        Label7(0).Height = 250
        Label7(1).Left = p2
        Label7(1).Width = Label2(1).Width
        Label7(1).Top = Label7(0).Top
        Label7(1).Height = Label7(0).Height
    
    DepthLabel.Left = p1 'depth
        DepthLabel.Width = Label2(0).Width
        DepthLabel.Top = Label7(0).Top + Label7(0).Height + 100
        DepthLabel.Width = p2 - p1
        DepthLabel.Height = 250
        DepthScroll.Top = DepthLabel.Top
        DepthScroll.Width = 1250
        DepthScroll.Left = p2
    
    RPdepthLabel.Left = p1 'RPdepth
        RPdepthLabel.Width = Label2(0).Width
        RPdepthLabel.Top = DepthLabel.Top + DepthLabel.Height + 50
        RPdepthLabel.Width = p3 - p1
        RPdepthLabel.Height = 250
        RPdepthScroll.Top = RPdepthLabel.Top + 300
        RPdepthScroll.Width = p3 - p1
        RPdepthScroll.Left = p1
    
    TwistLabel.Left = p1  'Leech twist
        TwistLabel.Width = p2 - p1
        TwistLabel.Top = RPdepthScroll.Top + RPdepthScroll.Height + 100
        TwistLabel.Height = 250
        TwistScroll.Top = TwistLabel.Top
        TwistScroll.Width = 1250
        TwistScroll.Left = p2

    NHpanelLabel.Left = p1 ' nombre panels tete
        NHpanelLabel.Top = TwistLabel.Top + TwistLabel.Height + 100
        NHpanelLabel.Height = 250
        NHpanelLabel.Width = p2 - p1
        NHpanelScroll.Left = p2
        NHpanelScroll.Top = NHpanelLabel.Top
        NHpanelScroll.Width = 1250

    NBpanelLabel.Left = p1 ' nombre panels bas
        NBpanelLabel.Top = NHpanelLabel.Top + NHpanelLabel.Height + 50
        NBpanelLabel.Height = 250
        NBpanelLabel.Width = p2 - p1
        NBpanelScroll.Left = p2
        NBpanelScroll.Top = NBpanelLabel.Top
        NBpanelScroll.Width = 1250
    
    sailLabel.Left = p1 'sailname
        sailLabel.Top = NBpanelLabel.Top + NBpanelLabel.Height + 250
        sailLabel.Width = p2 - p1
        sailLabel.Height = 250
        SailName.Left = p2
        SailName.Top = sailLabel.Top - 30
        SailName.Width = 1250

    Develop.Left = p1 + 30
        Develop.Width = p2 - p1 - 200
        Develop.Height = 350
        Develop.Top = LBattenScroll.Top - 100
    
    Cancel.Left = p2
        Cancel.Width = 1250
        Cancel.Height = Develop.Height
        Cancel.Top = Develop.Top
    '-----

    If Sail$ <> "NEW_JUNK" Then
        Create1.Caption = titre$ + "  Modifying the sail : " + Sail$
    Else
        Create1.Caption = titre$ + "  Creating a new sail"
    End If
    '-----
    redraw

End Sub ' Form_Resize ------------------------------------

Private Sub Form_Unload(Cancel As Integer)
    
    End

End Sub '-------------------------------------------------

Private Sub fprint_Click()
    
    MousePointer = 11 'change pointer to hour glass
    
    MousePointer = 0

End Sub  '-------------------------------------------------

Private Sub fprint_data_Click()

    Printer.Print Space(6); " ---------------"
    Printer.Print
    Printer.Print Space(6); titre$, " - Sail: "; Sail
    Printer.Print

    Printer.Print Space(6); "Batten length ="; LBattenScroll.Value
    Printer.Print Space(6); "Yard length   ="; LYardScroll.Value
    Printer.Print Space(6); "Foot Angle    ="; FootAScroll.Value; "deg"
    Printer.Print Space(6); "Yard Angle    ="; YardAScroll.Value; "deg"
    Printer.Print
    Printer.Print Space(6); "Head panel    = "; NHpanelScroll.Value
    Printer.Print Space(6); "Lower panels  = "; NBpanelScroll.Value
    Printer.Print Space(8); "(lower panels are split in 2 for development)"
    Printer.Print

    Printer.Print Space(6); "Upper panels luff width  ="; UpLuffScroll.Value
    Printer.Print Space(6); "Lower panels luff width  ="; LoLuffScroll.Value
    Printer.Print Space(6); "Lower panels leech width ="; LoLeechScroll.Value
    Printer.Print
    Printer.Print Space(6); "Lower panels depth ="; DepthScroll.Value; "%"
    Printer.Print Space(6); "    Depth position ="; RPdepthScroll.Value; "%"
    Printer.Print Space(6); "             Twist = "; TwistScroll.Value
    Printer.Print

    Printer.Print Space(6); " ---------------"
    Printer.EndDoc
End Sub    '-----------------------------------------------

Private Sub fprint_XYZ_Click()

    Dim n%, i%

    Const fmt2$ = "####0"

    For n = 1 To nBpanel + nHpanel
        Printer.Print
        If n > nBpanel Then
            Printer.Print Space(6); titre$, " - Sail: "; Sail, " Head panel  "; n
        Else
            Printer.Print Space(6); titre$; " - Sail: "; Sail, " Lower panel  "; n
        End If
        
        Printer.Print
        Printer.Print " ", "X lower edge", "Y lower edge", "X upper edge", "Y upper edge"
        Printer.Print
    
        For i = 0 To 20 Step 2
            Printer.Print " ", Format$(plx(n, i), fmt2), Format$(ply(n, i), fmt2), Format$(pmx(n, i), fmt2), Format$(pmy(n, i), fmt2)
        Next i
        
        Printer.Print
        Printer.Print " ", "X Mid luff", "Y Mid luff", "X Mid leech", "Y Mid leech"
        Printer.Print
        Printer.Print " ", Format$(pcx(n, 0), fmt2), Format$(pcy(n, 0), fmt2), Format$(pcx(n, 20), fmt2), Format$(pcy(n, 20), fmt2)
        Printer.Print
        Printer.Print "  ------------- "
        '-----
        If Printer.CurrentY > (Printer.Height / 2 - 200) Then
            Printer.NewPage
        End If
        '-----
    Next n
    '-----
    Printer.EndDoc

End Sub   '------------------------------------------------

Private Sub fquit_Click()

    Menu1.Show

End Sub '--------------------------------------------------

Private Sub fsave_data_Click()

    Dim fichier$
    '-----
    UpLuffScrl = UpLuffScroll.Value
    LoLuffScrl = LoLuffScroll.Value
    LoLeechScrl = LoLeechScroll.Value
    LBattenScrl = LBattenScroll.Value
    LyardScrl = LYardScroll.Value
    FootAScrl = FootAScroll.Value
    YardAScrl = YardAScroll.Value

    Mdepth = DepthScroll.Value / 100
    RPdepth = RPdepthScroll.Value / 100
    twistScrl = TwistScroll.Value
    nHpanel = NHpanelScroll.Value
    nBpanel = NBpanelScroll.Value
    '-----
    CMDialog1.CancelError = True
    
    On Error GoTo errhandler2
    '-----
    Sail$ = UCase$(Sail$)
    CMDialog1.Filter = "SAILCUT Files|*.sc8"
    CMDialog1.FileName = Sail$ + ".sc8"
    CMDialog1.Action = 2
    Sail$ = UCase$(CMDialog1.FileTitle)
    Sail$ = Left$(Sail$, (Len(Sail$) - 4))
    fichier$ = CMDialog1.FileName
    '-----
    Call F_Ecrire(fichier$)

    Exit Sub
    '-----
errhandler2:
    Close
    Screen.MousePointer = 0
    
    If Err <> 32755 Then
        MsgBox Error(Err), 0, "Create fsave_data"
    End If
    Exit Sub

End Sub ' fSave data -------------------------------------

Private Sub fSail_DXF_click()
    Dim fichier$
    Dim c%, i%, p%
    '-----
    CMDialog1.CancelError = True
    
    On Error GoTo errSailDXF
    '-----
    Sail$ = UCase$(Sail$)
    CMDialog1.Filter = "Sail DXF|*.DXF"
    CMDialog1.FileName = Sail$ + ".DXF"
    CMDialog1.Action = 2
    Sail$ = UCase$(CMDialog1.FileTitle)
    Sail$ = Left$(Sail$, (Len(Sail$) - 4))
    fichier$ = CMDialog1.FileName
    '-----
    Create1.MousePointer = 11
    Open fichier$ For Output As #1
        DXFHeader "junk rig sail " + Sail
        DXFSectionHeaderGeometry
    
        p = 0
        DXFPolyline 1, 7
        For i = 0 To 20
            DXFVertex 1, 7, px(p, i), py(p, i), pz(p, i)
        Next i
        DXFSequenceEnd
        
    For p = 1 To 2 * (nBpanel + nHpanel)
        DXFPolyline 1, 7
        DXFVertex 1, 7, px(p - 1, 0), py(p - 1, 0), pz(p - 1, 0)
        DXFVertex 1, 7, px(p, 0), py(p, 0), pz(p, 0)
        DXFSequenceEnd
        
        c = 7 - p Mod 2
        DXFPolyline 1, c
        For i = 0 To 20
            DXFVertex 1, c, px(p, i), py(p, i), pz(p, i)
        Next i
        DXFSequenceEnd
        
        DXFPolyline 1, 7
        DXFVertex 1, 7, px(p, 20), py(p, 20), pz(p, 20)
        DXFVertex 1, 7, px(p - 1, 20), py(p - 1, 20), pz(p - 1, 20)
        DXFSequenceEnd
    Next p
    
        DXFSectionEnd
        DXFEnd
        '-----
    Close #1
    Create1.MousePointer = 0
    Exit Sub
    '-----

errSailDXF:
    Close
    Create1.MousePointer = 0
    
    If Err <> 32755 Then
        MsgBox Error(Err), 0, "Create fSail_DXF"
    End If
    
    Exit Sub

End Sub ' fSail_DXF_click() ------------------------------

Private Sub fsave_DXF_Click()
    Dim fichier$
    '-----
    CMDialog1.CancelError = True
    
    On Error GoTo errSaveDXF
    '-----
    Sail$ = UCase$(Sail$)
    CMDialog1.Filter = "Panels DXF|*.DXF"
    CMDialog1.FileName = Sail$ + ".DXF"
    CMDialog1.Action = 2
    Sail$ = UCase$(CMDialog1.FileTitle)
    Sail$ = Left$(Sail$, (Len(Sail$) - 4))
    fichier$ = CMDialog1.FileName
    '-----
    Call F_DXF(fichier$)

    Exit Sub
    '-----

errSaveDXF:
    Close
    Create1.MousePointer = 0
    
    If Err <> 32755 Then
        MsgBox Error(Err), 0, "Create fsave_DXF"
    End If
    
    Exit Sub

End Sub ' fsave_DXF_Click --------------------------------

Private Sub fsave_vrml1_Click()
    ' VRML 1.0

Dim fichier$
'-----
CMDialog1.CancelError = True

On Error GoTo errhandvrml1
    '-----
    Sail$ = UCase$(Sail$)
    CMDialog1.Filter = "VRML Files|*.wrl"
    CMDialog1.FileName = Sail$ + ".wrl"
    CMDialog1.Action = 2
    Sail$ = UCase$(CMDialog1.FileTitle)
    Sail$ = Left$(Sail$, (Len(Sail$) - 4))
    fichier$ = CMDialog1.FileName
    '-----
    Call F_VRML1(fichier)

    Exit Sub
    '-----
errhandvrml1:
    Close
    If Err <> 32755 Then
        MsgBox Error(Err), 0, "Create fsave_VRML"
    End If
        Create1.MousePointer = 0
    Exit Sub

End Sub ' fSave VRML1 ------------------------------------

Private Sub fsave_XYZ_Click()
    Dim fichier$
    '-----
    CMDialog1.CancelError = True
    
    On Error GoTo errSaveXYZ
    '-----
    Sail$ = UCase$(Sail$)
    CMDialog1.Filter = "Panels X-Y|*.XYZ"
    CMDialog1.FileName = Sail$ + ".XYZ"
    CMDialog1.Action = 2
    Sail$ = UCase$(CMDialog1.FileTitle)
    Sail$ = Left$(Sail$, (Len(Sail$) - 4))
    fichier$ = CMDialog1.FileName
    '-----
    Call F_XY(fichier$)

    Exit Sub
    '-----

errSaveXYZ:
    Close
    Create1.MousePointer = 0
    
    If Err <> 32755 Then
        MsgBox Error(Err), 0, "Create fsave_XY"
    End If
    
    Exit Sub


End Sub ' fsave_XYZ_Click --------------------------------

Private Sub LBattenScroll_Change()
    ' longueur foot
    
    LBatten = LBattenScroll.Value

    form_activate

End Sub '-------------------------------------------------

Private Sub LoLeechScroll_Change()

    If LoLeechScroll.Value < LoLuffScroll.Value Then
        LoLeechScroll.Value = LoLuffScroll.Value
    End If

    LoLeech = LoLeechScroll.Value

    form_activate

End Sub ' ------------------------------------------------

Private Sub LoLuffScroll_Change()
        
        If LoLuffScroll.Value > LoLeechScroll.Value Then
            LoLuffScroll.Value = LoLeechScroll.Value
        End If
    
        LoLuff = LoLuffScroll.Value

        form_activate

End Sub  '------------------------------------------------

Private Sub LYardScroll_Change()
    
    If LYardScroll.Value < 0.5 * LBattenScroll.Value Then
        LYardScroll.Value = 0.5 * LBattenScroll.Value
    
    ElseIf LYardScroll.Value > 1.8 * LBattenScroll.Value Then
        LYardScroll.Value = 1.8 * LBattenScroll.Value
    End If

    LYard = LYardScroll.Value

    form_activate

End Sub    '----------------------------------------------

Private Sub NBpanelScroll_Change()
    
    NBpanelLabel.BackColor = &H80000005
    nBpanel = NBpanelScroll.Value
    
    form_activate

End Sub ' NBpanelScroll_Change ----------------------------

Private Sub NHpanelScroll_Change()
    
    nHpanel = NHpanelScroll.Value
    
    form_activate

End Sub ' ------------------------------------------------

Private Sub Picture1_Click()

    Picture1.Cls
    form_activate

End Sub '-------------------------------------------------

Private Sub redraw()
    Dim n%
    '-----
    Picture1.Cls
    Call compute
    'Call devjunk
    
    For n = 1 To nBpanel
        Call develop2(n)
    Next n
    
    For n = 1 To nHpanel
        Call develop2(nBpanel + n)
    Next n

    Call echelle(xo, yo, ks)
    Call dessinPlan
    Call dessinAR

End Sub ' redraw -----------------------------------------

Private Sub RPdepthScroll_Change()

    RPdepth = RPdepthScroll.Value / 100

    form_activate

End Sub  ' -----------------------------------------------

Private Sub SailName_KeyPress(keyascii As Integer)
    
    If keyascii = 13 Then
        Sailname_lostfocus
    End If

End Sub '-------------------------------------------------

Private Sub Sailname_lostfocus()

Dim i%, l%

    Sail = UCase$(SailName.Text)
    If Len(Sail) > 8 Then
        Beep
        SailName.BackColor = &HFFFF&
        Sail = Left$(Sail, 8)
    ElseIf Len(Sail) < 3 Then
        SailName.BackColor = &HFFFF&
        Sail = "SPI" + Sail
    Else
        SailName.BackColor = &H80000005
    End If
    
    If Left$(Sail, 1) < "A" Or Left$(Sail, 1) > "Z" Then
        Beep
        Sail = "A" + Right$(Sail, Len(Sail) - 1)
        SailName.BackColor = &HFF&
    End If

    For i = 1 To Len(Sail)
        l = Asc(Mid$(Sail, i, 1))
        If l = 95 Then
            'Next i
        ElseIf (l >= 48 And l <= 57) Or (l >= 65 And l <= 90) Then
            'Next i
        Else
            Mid$(Sail, i, 1) = "_"
            SailName.BackColor = &HFF&
        End If
    Next i
    '-----
    SailName.Text = UCase(Sail)

End Sub

Private Sub takeData()
    
    LYard = LYardScroll.Value
    LoLuff = LoLuffScroll.Value
    UpLuff = UpLuffScroll.Value
    LoLeech = LoLeechScroll.Value
    LBatten = LBattenScroll.Value
    
    AYard = YardAScroll.Value
    AFoot = FootAScroll.Value

    nBpanel = NBpanelScroll.Value
    nHpanel = NHpanelScroll.Value

    Mdepth = DepthScroll.Value / 100
    RPdepth = RPdepthScroll.Value / 100

    twistScrl = TwistScroll.Value
    Atwist = twistScrl / 57

End Sub ' takeData ------------------------------------

Private Sub TwistScroll_Change()
    ' vrillage

    twistScrl = TwistScroll.Value
    Atwist = twistScrl / 57
    '--
    form_activate

End Sub '-------------------------------------------------

Private Sub UpLuffScroll_Change()
        
    If UpLuffScroll.Value > LoLeechScroll.Value Then
        UpLuffScroll.Value = LoLeechScroll.Value
    End If

    UpLuff = UpLuffScroll.Value
    
    form_activate

End Sub  '------------------------------------------------

Private Sub YardAScroll_Change()
    
    If YardAScroll.Value < FootAScroll.Value Then
        YardAScroll.Value = FootAScroll.Value
    ElseIf YardAScroll.Value < AHead * RAD Then
        YardAScroll.Value = AHead * RAD + 0.5
    End If
    
    AYard = YardAScroll.Value

    form_activate

End Sub  '------------------------------------------------

