VERSION 5.00
Begin VB.Form Menu1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Sailcut8"
   ClientHeight    =   4530
   ClientLeft      =   2685
   ClientTop       =   3765
   ClientWidth     =   5520
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
   Icon            =   "menu1.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4530
   ScaleWidth      =   5520
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "English"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      Top             =   3960
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Français"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Terminate 
      Appearance      =   0  'Flat
      Caption         =   "TERMINATE and EXIT"
      Height          =   492
      Left            =   2760
      TabIndex        =   2
      Top             =   2880
      Width           =   2532
   End
   Begin VB.CommandButton Create 
      Appearance      =   0  'Flat
      Caption         =   "CREATE A NEW SAIL"
      Height          =   492
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
      Width           =   2532
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2172
      Left            =   120
      ScaleHeight     =   2145
      ScaleWidth      =   1905
      TabIndex        =   4
      Top             =   1680
      Width           =   1932
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1    Menu1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Menu1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit   '29 décembre 2002

Private Sub Create_Click()

    MousePointer = 11  'change pointer to hourglass
    Create1.Show

End Sub '--------------------------------------------------

Private Sub form_activate()
    
    Create1.MousePointer = 0
    Menu1.MousePointer = 0

End Sub '--------------------------------------------------

Private Sub Form_Load()
    Dim large, hauteur, haut, gauche As Single
    Dim i As Integer
    '-----
    timenow = Now
    
    Cls
    Menu1.Visible = False
    Menu1.Width = 6000
    Menu1.Height = 5800
    Menu1.Top = (Screen.Height - Menu1.Height) / 2
    Menu1.Left = (Screen.Width - Menu1.Width) / 2
    Menu1.Caption = titre + "       version :" + version
    Label1.Top = 200
    Label1.Height = 300
    Label1.Width = Menu1.Width * 0.9
    Label1.Left = (Menu1.Width - Label1.Width) / 2
    
    Label2.Top = Label1.Top + Label1.Height + 100
    Label2.Height = 2000
    Label2.Width = Label1.Width
    Label2.Left = Label1.Left
    
    Picture1.Top = Label2.Top + Label2.Height + 200
    Create.Top = Picture1.Top
    Create.Left = Menu1.Width / 2
    Create.Width = Menu1.Width / 2.2


    Terminate.Top = Create.Top + Create.Height + 250
    Terminate.Left = Create.Left
    Terminate.Width = Create.Width
    
    Option1(0).Left = Create.Left
    Option1(0).Top = Terminate.Top + Terminate.Height + 250
    
    Option1(1).Left = Create.Left
    Option1(1).Top = Option1(0).Top + Option1(0).Height + 100
    Option1(1).Value = True

    '---
    Picture1.Width = Menu1.Width / 2.4
    ks = Picture1.Width / 1500
    Picture1.DrawWidth = 6
    
        Picture1.PSet (1000 * ks, 300 + i)
        Picture1.Line -Step(-400 * ks, 500)
        Picture1.Line -Step(-150 * ks, 250)
        Picture1.Line -Step(-10 * ks, 100)
        Picture1.Line -Step(50 * ks, 100)
        Picture1.Line -Step(100 * ks, 50)
        Picture1.Line -Step(200 * ks, 50)
        Picture1.Line -Step(200 * ks, 0)
        Picture1.Line -Step(100 * ks, -50)
        Picture1.Line -Step(25 * ks, -100)
        Picture1.Line -Step(-25 * ks, -100)
        Picture1.Line -Step(-100 * ks, -100)
        Picture1.Line -Step(-90 * ks, -0)
        Picture1.Line -Step(-200 * ks, 100)
        Picture1.Line -Step(-100 * ks, 100)
        Picture1.Line -Step(-200 * ks, 300)
        Picture1.Line -Step(-200 * ks, 400)
        Picture1.PSet (900 * ks, 1350 + i)
        Picture1.Line -Step(300 * ks, 600)
    '------
    Menu1.Show
    '--- chargement des valeurs par défaut
    Sail = "NEW_JUNK"
    LBatten = 3000
    LYard = 3000
    AFoot = 15
    AYard = 70
    LoLeech = 1000
    ClothW = 900
    SeamW = 25
    nHpanel = 3
    nBpanel = 4
    RPdepth = 0.4

End Sub '-------------------------------------------------

Private Sub Form_Unload(Cancel As Integer)
    
    End

End Sub '-------------------------------------------------

Private Sub LOGO()
    ' logo RL
    Dim i As Integer
    Picture1.DrawWidth = 6
    'For i = 0 To 4 Step 2
        Picture1.PSet (1000, 300 + i)
        Picture1.Line -Step(-400, 500)
        Picture1.Line -Step(-150, 250)
        Picture1.Line -Step(-10, 100)
        Picture1.Line -Step(50, 100)
        Picture1.Line -Step(100, 50)
        Picture1.Line -Step(200, 50)
        Picture1.Line -Step(200, 0)
        Picture1.Line -Step(100, -50)
        Picture1.Line -Step(25, -100)
        Picture1.Line -Step(-25, -100)
        Picture1.Line -Step(-100, -100)
        Picture1.Line -Step(-90, -0)
        Picture1.Line -Step(-200, 100)
        Picture1.Line -Step(-100, 100)
        Picture1.Line -Step(-200, 300)
        Picture1.Line -Step(-200, 400)
        Picture1.PSet (900, 1350 + i)
        Picture1.Line -Step(300, 600)
    'Next i

End Sub    '----------------------------------------------

Private Sub Option1_Click(Index As Integer)
    ' choix langue
If Index = 0 Then
    langue = 0
    
    Label1.Caption = "Bienvenue aux maîtres voiliers"
    
    Label2.Caption = "Ce programme pour concevoir des voiles de jonques est mis à disposition pour évaluation et sans aucune garantie sous"
    Label2.Caption = Label2.Caption + Chr$(13) + "GNU General Public Licence version 2, tel que publié par la Free Software Foundation."
    Label2.Caption = Label2.Caption + Chr$(13)
    Label2.Caption = Label2.Caption + Chr$(13) + "Copyright et autres droits réservés par Robert Lainé"
    Label2.Caption = Label2.Caption + Chr$(13) + "Sailcut est une marque déposée" + Chr$(13)
    
    Create.Caption = "Débuter"
    Terminate.Caption = "Fin"
Else
    langue = 1

    Label1.Caption = "Welcome to all sailmakers"

    Label2.Caption = "This software for designing junk sails is made available for evaluation purpose and without any warranty under"
    Label2.Caption = Label2.Caption + Chr$(13) + "GNU General Public Licence version 2, as published by the Free Software Foundation."
    Label2.Caption = Label2.Caption + Chr$(13)
    Label2.Caption = Label2.Caption + Chr$(13) + "Copyright and other rights reserved by Robert Lainé"
    Label2.Caption = Label2.Caption + Chr$(13) + "Sailcut is a registered trademark" + Chr$(13)
    Create.Caption = "Begin"
    Terminate.Caption = "End"
    '------
End If

    Label2.Caption = Label2.Caption + Chr$(13) + "E-mail: robert.laine@sailcut.com"

End Sub ' Option1 langue ---------------------------------

Private Sub Terminate_Click()

    Menu1.Hide
    End

End Sub ' Terminate_click --------------------------------

