VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Create1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sailcut8   Creating a new sail"
   ClientHeight    =   7260
   ClientLeft      =   5460
   ClientTop       =   2160
   ClientWidth     =   10020
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7260
   ScaleWidth      =   10020
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   14
      Left            =   7800
      TabIndex        =   43
      Text            =   "(14)"
      Top             =   4560
      Visible         =   0   'False
      Width           =   715
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   364
      Left            =   6669
      TabIndex        =   41
      Top             =   5616
      Width           =   2470
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   13
      Left            =   7722
      TabIndex        =   40
      Text            =   "(13)"
      Top             =   3978
      Width           =   715
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   12
      Left            =   7722
      TabIndex        =   39
      Text            =   "(12)"
      Top             =   3744
      Width           =   715
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   11
      Left            =   7722
      TabIndex        =   38
      Text            =   "(11)"
      Top             =   3510
      Width           =   715
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   10
      Left            =   7722
      TabIndex        =   37
      Text            =   "(10)"
      Top             =   3276
      Width           =   715
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   9
      Left            =   7722
      TabIndex        =   36
      Text            =   "(9)"
      Top             =   2808
      Width           =   715
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   8
      Left            =   7722
      TabIndex        =   35
      Text            =   "(8)"
      Top             =   2457
      Width           =   715
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   7
      Left            =   7722
      TabIndex        =   34
      Text            =   "(7)"
      Top             =   2106
      Width           =   715
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   6
      Left            =   7722
      TabIndex        =   33
      Text            =   "(6)"
      Top             =   1872
      Width           =   715
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   5
      Left            =   7722
      TabIndex        =   32
      Text            =   "(5)"
      Top             =   1404
      Width           =   715
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   4
      Left            =   7722
      TabIndex        =   31
      Text            =   "(4)"
      Top             =   1053
      Width           =   715
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   3
      Left            =   7722
      TabIndex        =   30
      Text            =   "(3)"
      Top             =   819
      Width           =   715
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   2
      Left            =   7722
      TabIndex        =   29
      Text            =   "(2)"
      Top             =   468
      Width           =   715
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Index           =   1
      Left            =   7722
      TabIndex        =   28
      Text            =   "(1)"
      Top             =   117
      Width           =   715
   End
   Begin VB.HScrollBar YardAScroll 
      Height          =   364
      LargeChange     =   10
      Left            =   9477
      Max             =   89
      TabIndex        =   24
      Top             =   2106
      Width           =   364
   End
   Begin VB.HScrollBar LoLeechScroll 
      Height          =   364
      LargeChange     =   50
      Left            =   9477
      Max             =   2000
      Min             =   100
      SmallChange     =   5
      TabIndex        =   23
      Top             =   819
      Value           =   2000
      Width           =   364
   End
   Begin VB.HScrollBar UpLuffScroll 
      Height          =   364
      LargeChange     =   50
      Left            =   9477
      Max             =   1000
      Min             =   50
      SmallChange     =   5
      TabIndex        =   22
      Top             =   117
      Value           =   1000
      Width           =   364
   End
   Begin VB.HScrollBar NBpanelScroll 
      Height          =   255
      LargeChange     =   4
      Left            =   9711
      Max             =   32
      TabIndex        =   11
      Top             =   3978
      Value           =   4
      Width           =   390
   End
   Begin VB.HScrollBar NHpanelScroll 
      Height          =   255
      LargeChange     =   2
      Left            =   9711
      Max             =   9
      Min             =   1
      TabIndex        =   14
      Top             =   3744
      Value           =   2
      Width           =   403
   End
   Begin VB.HScrollBar LYardScroll 
      Height          =   364
      LargeChange     =   100
      Left            =   9477
      Max             =   10000
      Min             =   100
      SmallChange     =   10
      TabIndex        =   3
      Top             =   1404
      Value           =   2500
      Width           =   364
   End
   Begin VB.HScrollBar FootAScroll 
      Height          =   364
      LargeChange     =   5
      Left            =   9477
      Max             =   40
      TabIndex        =   5
      Top             =   1755
      Value           =   40
      Width           =   364
   End
   Begin VB.HScrollBar LoLuffScroll 
      Height          =   364
      LargeChange     =   50
      Left            =   9477
      Max             =   2000
      Min             =   100
      SmallChange     =   5
      TabIndex        =   4
      Top             =   468
      Value           =   2000
      Width           =   364
   End
   Begin VB.HScrollBar LBattenScroll 
      Height          =   364
      LargeChange     =   100
      Left            =   9477
      Max             =   10000
      Min             =   100
      SmallChange     =   10
      TabIndex        =   6
      Top             =   1053
      Value           =   2800
      Width           =   364
   End
   Begin VB.HScrollBar DepthScroll 
      Height          =   255
      LargeChange     =   5
      Left            =   9594
      Max             =   12
      TabIndex        =   17
      Top             =   2808
      Value           =   6
      Width           =   390
   End
   Begin VB.HScrollBar TwistScroll 
      Height          =   255
      LargeChange     =   5
      Left            =   9711
      Max             =   24
      TabIndex        =   16
      Top             =   3510
      Value           =   12
      Width           =   390
   End
   Begin VB.HScrollBar RPdepthScroll 
      Height          =   240
      LargeChange     =   5
      Left            =   9711
      Max             =   96
      Min             =   25
      TabIndex        =   15
      Top             =   3276
      Value           =   43
      Width           =   325
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
      TabIndex        =   2
      Top             =   6240
      Width           =   2532
   End
   Begin VB.CommandButton Develop 
      Appearance      =   0  'Flat
      Caption         =   "develop"
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   5148
      Width           =   2535
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
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "VanLornratio(14)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   6120
      TabIndex        =   42
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Yard L.  L1(5)"
      ForeColor       =   &H80000008&
      Height          =   260
      Index           =   5
      Left            =   6123
      TabIndex        =   27
      Top             =   1443
      Width           =   1573
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lo Leech W L1(3)"
      ForeColor       =   &H80000008&
      Height          =   260
      Index           =   3
      Left            =   6123
      TabIndex        =   26
      Top             =   845
      Width           =   1573
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Yard angle  L1(7)"
      ForeColor       =   &H80000008&
      Height          =   260
      Index           =   7
      Left            =   6123
      TabIndex        =   25
      Top             =   2158
      Width           =   1573
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Up Luff L  L1(1)"
      ForeColor       =   &H80000008&
      Height          =   377
      Index           =   1
      Left            =   6123
      TabIndex        =   21
      Top             =   117
      Width           =   1573
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Surface L1(8)"
      ForeColor       =   &H80000008&
      Height          =   260
      Index           =   8
      Left            =   6123
      TabIndex        =   18
      Top             =   2522
      Width           =   1573
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "NB panel L1(13)"
      ForeColor       =   &H80000008&
      Height          =   260
      Index           =   13
      Left            =   6201
      TabIndex        =   19
      Top             =   4095
      Width           =   1573
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "NHpanel L1(12)"
      ForeColor       =   &H80000008&
      Height          =   260
      Index           =   12
      Left            =   6201
      TabIndex        =   20
      Top             =   3861
      Width           =   1573
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Depth RP L1(10)"
      ForeColor       =   &H80000008&
      Height          =   260
      Index           =   10
      Left            =   6123
      TabIndex        =   13
      Top             =   3237
      Width           =   1456
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Twist L1(11) "
      ForeColor       =   &H80000008&
      Height          =   260
      Index           =   11
      Left            =   6123
      TabIndex        =   12
      Top             =   3510
      Width           =   1573
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Depth   L1(9)"
      ForeColor       =   &H80000008&
      Height          =   377
      Index           =   9
      Left            =   6123
      TabIndex        =   9
      Top             =   2886
      Width           =   1573
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Foot angle  L1(6)"
      ForeColor       =   &H80000008&
      Height          =   260
      Index           =   6
      Left            =   6123
      TabIndex        =   8
      Top             =   1872
      Width           =   1573
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Batten L.  L1(4)"
      ForeColor       =   &H80000008&
      Height          =   260
      Index           =   4
      Left            =   6123
      TabIndex        =   7
      Top             =   1079
      Width           =   1573
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lo Luff L  L1(2)"
      ForeColor       =   &H80000008&
      Height          =   260
      Index           =   2
      Left            =   6123
      TabIndex        =   1
      Top             =   481
      Width           =   1573
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
   Begin VB.Menu mnuSailStyle 
      Caption         =   "Sail Style"
      Begin VB.Menu mnuSailStyleStd 
         Caption         =   "Standard"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSailStyleVanLorn 
         Caption         =   "Van Lorn"
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

Option Explicit
' 20 December 2004******************************************************
' Copyright (C) 1997-2004 Robert Lainé
' Sailcut is a trademark registered by Robert Lainé
' See CREDITS file for a full list of contributors.
'
' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'***************************************************************************
 
Private Sub Cancel_Click()
Dim response As Long

IniKeysSet
If infoDirty = True Then
    response = MsgBox( _
    Prompt:="Save changes to " & Sail$ & "?", _
    Buttons:=vbYesNoCancel + vbQuestion + vbDefaultButton2, _
    Title:="Save" & Sail$)

    Select Case response
        Case vbYes
            fsave_data_Click
            Create1.Hide
            Menu1.Show
        Case vbNo
            Create1.Hide
            Menu1.Show
            Exit Sub
        Case vbCancel
            Exit Sub
    End Select
Else
    Create1.Hide
    Menu1.Show
End If

End Sub '--------------------------------------------------

Private Sub chkBatten()

    If LBattenScroll.value < LoLuffScroll.value Then
        LBattenScroll.value = LoLuffScroll.value
        infoDirty = True
        Label1(4).BackColor = vbYellow
        '-----
    Else
        Label1(4).BackColor = vbWindowBackground
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
Private Sub CheckFlags()
Dim i%
    ' reset label colors
    For i = 1 To 13
        Label1(i).BackColor = vbWindowBackground
    Next i
    If SailStyle = 0 Then
        'check if top of lower panels angle is ok
        If AngleMaxOK = 1 Then Label1(13).BackColor = vbYellow
        If AngleMaxOK = 2 Then Label1(13).BackColor = vbRed
        
        If AngleDir = False Then
            If Label1(13).BackColor = vbRed Then
                Label1(2).BackColor = vbRed
                Label1(3).BackColor = vbRed
              
              Else
                Label1(13).BackColor = vbYellow
                Label1(2).BackColor = vbYellow
                Label1(3).BackColor = vbYellow
                
            End If
        End If
    End If
    
    If YardReset = True Then YardAScroll.value = AYard
    If UpLeechGood = False Then
        Label1(12).BackColor = vbYellow
        Label1(1).BackColor = vbYellow
        Label1(5).BackColor = vbYellow
    End If
    

End Sub ' checkflags

Private Sub defaut()
    ' default sail
    
    
    Sail = "NEW_JUNK"
        
    UpLuffScroll.value = 150   'mm
    LoLuffScroll.value = 700   'mm
    LoLeechScroll.value = 1000 'mm
    LBattenScroll.value = 2800 'mm
    LYardScroll.value = 2500   'mm

    FootAScroll.value = 0  'deg
    YardAScroll.value = 75  'deg

    DepthScroll.value = 6    '%
    RPdepthScroll.value = 43 '%
    
    TwistScroll.value = 12  'degré

    NHpanelScroll.value = 2
    NBpanelScroll.value = 4

    ClothW = 900 'mm
    SeamW = 50   'mm
    SeamT = 15   'mm

    genre1 = "JunkSail"
    genre2 = "default_2"
    genre3 = "default-3"

End Sub ' defaut -----------------------------------------
Private Sub VanLornDefault()
    ' default sail
    
    
    Sail = "New Van Lorn Junk"
        
    'UpLuffScroll.value = 1  'mm
    LoLuffScroll.value = 1000   'mm
    LoLeechScroll.value = 1000 'mm
    LBattenScroll.value = 3300 'mm
    LYardScroll.value = 3300   'mm

    FootAScroll.value = 5  'deg
    YardAScroll.value = 35  'deg

    DepthScroll.value = 6    '%
    RPdepthScroll.value = 43 '%
    
    TwistScroll.value = 12  'degré

    NHpanelScroll.value = 2
    NBpanelScroll.value = 4

    ClothW = 900 'mm
    SeamW = 50   'mm
    SeamT = 15   'mm

    genre1 = "JunkSail"
    genre2 = "default_2"
    genre3 = "default-3"

End Sub '

Private Sub DepthScroll_Change()

    Mdepth(0) = DepthScroll.value / 100
    infoDirty = True
    
    If Me.Visible And Not infoRead Then form_activate

End Sub ' Sub DepthScroll_Change() ------------------------

Private Sub dessinDev()
    Dim couleur As Long
    Dim n, i   As Integer
    Dim xb, yb   As Single
   
'----- drawing development
For n = 1 To nBpanel + nHpanel
    If n > nBpanel Then  'upper part
        couleur = RGB(200, 0, 200)
    Else
        couleur = RGB(200, 0, 0)
    End If

    Picture1.DrawWidth = 1
    xb = (Picture1.Width - 300) - (1.1 * LBatten * ks)
    'xb = 1.3 * LBatten * ks
    
    ' compute origin of panel drawing
    yb = -200
    For i = 1 To n
        yb = yb + Abs(pmy(i - 1, 20) - ply(i - 1, 20)) * ks + 80
    Next i
    
    yb = yo + yb

    ' draw lower edge
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
    ' draw sail seen from behind
    Dim couleur As Long
    Dim n, i As Integer
    Dim xa, ya As Single
    '--- new origin
    xa = xo + 1.15 * LBatten * ks
    ya = yo
    
    '---- foot
    Picture1.DrawWidth = 2
    Picture1.PSet (xa + pz(0, 0) * ks, ya + py(0, 0) * ks)
    For i = 0 To 20
        Picture1.Line -(xa + pz(0, i) * ks, ya + py(0, i) * ks)
    Next i

'----- batten
For n = 2 To 2 * (nBpanel + nHpanel) Step 2
    If n > 2 * nBpanel Then
        couleur = RGB(200, 0, 200) 'head part
      ElseIf n = 2 * nBpanel Then
        couleur = RGB(0, 0, 0) 'transition lower-head part
      Else
        couleur = RGB(200, 0, 0) 'lower part
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

'----- intermediate lines
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
    ' draw sail
    Dim couleur As Long
    Dim n, i As Integer
    Dim x, y As Single
    
'----- draw profile below sail
    Picture1.DrawWidth = 1
    Picture1.PSet (xo, 50)
    For i = 0 To 50
        x = i / 50
        y = Mdepth(0) * LBatten * profileP(RPdepth, x)
        Picture1.Line -(xo + (LBatten * x) * ks, 50 + y * ks), RGB(0, 127, 0)
    Next i
        Picture1.Line -(xo, 50), RGB(0, 0, 0)
        
'----- draw sail
    ' foot
    Picture1.DrawWidth = 2
    Picture1.PSet (xo + px(0, 0) * ks, yo + py(0, 0) * ks)
    For i = 0 To 20
        Picture1.Line -(xo + px(0, i) * ks, yo + py(0, i) * ks)
    Next i

'----- batten
For n = 2 To 2 * (nBpanel + nHpanel) Step 2
    If n > 2 * nBpanel Then
        couleur = RGB(200, 0, 200) ' head part
      ElseIf n = 2 * nBpanel Then
        couleur = RGB(0, 0, 0) 'transition head-lower
      Else
        couleur = RGB(200, 0, 0) ' lower part
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

'----- intermediate lines
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
    ' development of sail

    MousePointer = 11 'change pointer to hour glass
    
    Call dessinDev
        
    MousePointer = 0

End Sub '------------------------------------------------

Private Sub develop_hi(pan)
    '--- develop head panels
    
    Dim p, i, lo, hi As Integer
    Dim alfa, beta, gama As Single
    Dim a, b, c, r As Single
    '-----
    p = 2 * nBpanel + 2 * pan
    lo = 2 * nBpanel + 2 * pan - 2
    hi = 2 * nBpanel + 2 * pan

    plx(p, 0) = 0
    ply(p, 0) = 0

    alfa = Atn((py(lo, 20) - py(lo, 0)) / (px(lo, 20) - px(lo, 0)))

    For i = 1 To 20   'lower edge
        r = Sqr((px(lo, i) - px(lo, 0)) ^ 2 + (py(lo, i) - py(lo, 0)) ^ 2)
        beta = directionXY(px(lo, 0), py(lo, 0), px(lo, i), py(lo, i))
        plx(p, i) = r * Cos(beta - alfa)
        ply(p, i) = r * Sin(beta - alfa)
    Next i

    For i = 0 To 20   'upper edge
        r = Sqr((px(hi, i) - px(lo, 0)) ^ 2 + (py(hi, i) - py(lo, 0)) ^ 2)
        beta = directionXY(px(lo, 0), py(lo, 0), px(hi, i), py(hi, i))
        pmx(p, i) = r * Cos(beta - alfa)
        pmy(p, i) = r * Sin(beta - alfa)
    Next i
    '-----

End Sub ' develop_hi -------------------------------------

Private Sub develop2(pan%)
    'develop lower panels
    
    Dim p, i, n As Integer
    Dim alfa, beta, gama As Single
    Dim a, b, c As Single
    Dim h, v As Single

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
    
    '--- first triangle up
    b = distance3D(px(n - 1, 0), py(n - 1, 0), pz(n - 1, 0), px(n, 1), py(n, 1), pz(n, 1))
    a = distance3D(px(n - 1, 0), py(n - 1, 0), pz(n - 1, 0), px(n, 0), py(n, 0), pz(n, 0))
    c = distance3D(px(n, 0), py(n, 0), pz(n, 0), px(n, 1), py(n, 1), pz(n, 1))
    Call Triangle(a, b, c, alfa, beta, gama)
    alfa = directionXY(pcx(pan, 0), pcy(pan, 0), pmx(pan, 1), pmy(pan, 1))
    pmx(pan, 0) = pcx(pan, 0) + a * Cos(alfa + gama)
    pmy(pan, 0) = pcy(pan, 0) + a * Sin(alfa + gama)
  
    '----- lower points
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
    
    '--- first triangle low
    b = distance3D(px(n - 1, 0), py(n - 1, 0), pz(n - 1, 0), px(n - 2, 1), py(n - 2, 1), pz(n - 2, 1))
    a = distance3D(px(n - 1, 0), py(n - 1, 0), pz(n - 1, 0), px(n - 2, 0), py(n - 2, 0), pz(n - 2, 0))
    c = distance3D(px(n - 2, 0), py(n - 2, 0), pz(n - 2, 0), px(n - 2, 1), py(n - 2, 1), pz(n - 2, 1))
    Call Triangle(a, b, c, alfa, beta, gama)
    alfa = directionXY(pcx(pan, 0), pcy(pan, 0), plx(pan, 1), ply(pan, 1))
    plx(pan, 0) = pcx(pan, 0) + a * Cos(alfa - gama)
    ply(pan, 0) = pcy(pan, 0) + a * Sin(alfa - gama)
  
    '--- reframing vertically
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
    
    '--- reframing horizontally
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
    Dim n, p, i As Integer
    Dim alfa, beta, gama As Single
    Dim a, b, c, r As Single
    Erase plx, ply, pmx, pmy

'----- copy points xyz
For n = 1 To 2 * (nBpanel + nHpanel)
    For i = 0 To 20
        plx(n, i) = px(n - 1, i) - px(n - 1, 0)
        ply(n, i) = py(n - 1, i) - py(n - 1, 0)
        pmx(n, i) = px(n, i) - px(n - 1, 0)
        pmy(n, i) = py(n, i) - py(n - 1, 0)
    Next i
Next n

'----- reframing
For n = 1 To 2 * (nBpanel + nHpanel)
    p = n Mod 2
    '-----
    Select Case p

    Case 1  ' panel above a baton

        alfa = Atn((ply(n, 20) - ply(n, 0)) / (plx(n, 20) - plx(n, 0)))
    
        For i = 1 To 20   'lower edge
            r = Sqr((plx(n, i) - plx(n, 0)) ^ 2 + (ply(n, i) - ply(n, 0)) ^ 2)
            beta = directionXY(plx(n, 0), ply(n, 0), plx(n, i), ply(n, i))
            plx(n, i) = r * Cos(beta - alfa)
            ply(n, i) = r * Sin(beta - alfa)
        Next i
    
        For i = 0 To 20   'upper edge
            r = Sqr((pmx(n, i) - plx(n, 0)) ^ 2 + (pmy(n, i) - ply(n, 0)) ^ 2)
            beta = directionXY(plx(n, 0), ply(n, 0), pmx(n, i), pmy(n, i))
            pmx(n, i) = r * Cos(beta - alfa)
            pmy(n, i) = r * Sin(beta - alfa)
        Next i
    '-----

    Case 0  'panel below baton

        alfa = Atn((pmy(n, 20) - pmy(n, 0)) / (pmx(n, 20) - pmx(n, 0)))
    
        For i = 1 To 20   'lower edge - bord inferieur
            r = Sqr((plx(n, i) - plx(n, 0)) ^ 2 + (ply(n, i) - ply(n, 0)) ^ 2)
            beta = directionXY(plx(n, 0), ply(n, 0), plx(n, i), ply(n, i))
            plx(n, i) = r * Cos(beta - alfa)
            ply(n, i) = r * Sin(beta - alfa)
        Next i
    
        For i = 0 To 20   'upper edge - bord superieur
            r = Sqr((pmx(n, i) - plx(n, 0)) ^ 2 + (pmy(n, i) - ply(n, 0)) ^ 2)
            beta = directionXY(plx(n, 0), ply(n, 0), pmx(n, i), pmy(n, i))
            pmx(n, i) = r * Cos(beta - alfa)
            pmy(n, i) = r * Sin(beta - alfa)
        Next i
        '--- reframing vertically
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
        
        '--- reframing horizontally
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

    Dim sw, sh, w, h, h1, h2, k1, k3 As Single
    
    sw = Picture1.Width
    sh = Picture1.Height
    
    Picture1.Scale (0, sh)-(sw, 0)
    
    xo = 200
    yo = 400
    '--- width - largeur
    k1 = (sw - 300) / (100 + LBatten)
    '--- height - hauteur
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
    
    AFoot = FootAScroll.value
    infoDirty = True

    If Me.Visible And Not infoRead Then form_activate

End Sub '----------------------------------------------

Private Sub fopen_Click()
    ' load a file - charger un fichier
    Dim fichier$
'---
CMDialog1.CancelError = True
infoRead = True

On Error GoTo errhandler1
    CMDialog1.CancelError = True
    CMDialog1.Filter = "SAILCUT8 files|*.sc8"
    CMDialog1.filename = Sail + ".sc8"
    CMDialog1.Action = 1
    Sail$ = CMDialog1.FileTitle
    Sail$ = Left$(Sail$, (Len(Sail$) - 4))
    fichier$ = CMDialog1.filename
    
    Call F_Lire(fichier$)

'    Text1(14).BackColor = vbWindowBackground
'    Text1(14).Text = Sail

    UpLuffScroll.value = UpLuffScrl
    LoLuffScroll.value = LoLuffScrl
    LoLeechScroll.value = LoLeechScrl
    LBattenScroll.value = LBattenScrl
    LYardScroll.value = LyardScrl
    FootAScroll.value = FootAScrl
    YardAScroll.value = YardAScrl

    DepthScroll.value = Mdepth(0) * 100
    RPdepthScroll.value = RPdepth * 100
    TwistScroll.value = twistScrl
    
    NHpanelScroll.value = nHpanel
    NBpanelScroll.value = nBpanel
    '-----
    
    form_activate
    infoRead = False
    infoDirty = False
    
    '-----
    Exit Sub

errhandler1:
    Close
    If Err <> 32755 Then
        MsgBox Error(Err), 0, "Create fopen"
    End If
        Create1.MousePointer = 0
        infoRead = False
    Exit Sub

End Sub ' fopen ------------------------------------------

Private Sub form_activate()
        Dim j As Integer
'------Create1 Activate

    Create1.MousePointer = 11  ' pointer type hourglass
    
    redraw  ' including scale factor computation

    
'----- display dimensions
        mSCVars_LoadArray
        mUnits_LoadArray


'----- display texts
        Dim i As Integer
        i = Langue
        
            If UCase(Sail$) <> "NEW_JUNK" Then
                Create1.Caption = titre$ & Lang(9, i) & Sail$
            Else
                Create1.Caption = titre$ & Lang(10, i)
            End If
        
        For j = 1 To 13
            Label1(j).Caption = Lang(10 + j, i)
            Text1(j).Text = SCVars(j)
        Next j
        
        
        ChangeUnitType
        



    
    Develop.Caption = Lang(25, i)
    Save.Caption = Lang(41, i)
    Cancel.Caption = Lang(26, i)
    
    fmenu.Caption = Lang(27, i)
    fnew.Caption = Lang(28, i)
    fopen.Caption = Lang(29, i)
    
    fprint.Caption = Lang(31, i)
    fprint_data.Caption = Lang(32, i)
    fprint_XYZ.Caption = Lang(33, i)

    fsave_DXF.Caption = Lang(30, i)
    fsave_XYZ.Caption = Lang(33, i)
    fsave.Caption = Lang(34, i)
    fsave_data.Caption = Lang(35, i)
    fSail_DXF.Caption = Lang(36, i)
    fsave_vrml1.Caption = Lang(37, i)
    
    fquit.Caption = Lang(38, i)

    clothMenu.Caption = Lang(39, i)
    clothOpen.Caption = Lang(40, i)
 
'------
    
    Create1.MousePointer = 0
    'Menu1.MousePointer = 0

End Sub ' Form_Activate ----------------------------------

Private Sub Form_Load()
    ' Create1 Form Load
    ' Check if INI file had diff values
    
    If WinWidth <> 0 Then
        Width = WinWidth
    Else
        Width = 0.9 * Screen.Width
    End If

    If WinHeight <> 0 Then
        Height = WinHeight
    Else
        Height = 0.7 * Width
    End If

    If WinTop <> 0 Then
        Top = WinTop
    Else
        Top = (Screen.Height - Height) / 2 - 100
    End If

    If WinLeft <> 0 Then
        Left = WinLeft
    Else
        Left = (Screen.Width - Width) / 2 - 100
    End If
    
    fsave_vrml1.Visible = True
    clothMenu.Visible = False
    
    If SailStyle = 1 Then
        VanLornDefault
    Else
        defaut      ' load default sail
    End If
    
End Sub ' Form_Load --------------------------------------

Private Sub Form_Resize()

    Dim p1 As Integer   'positions of labels
    Dim intLabelTop As Integer
    Dim i As Integer

    '-----

    With Picture1
        .Left = 60
        .Top = 60
        .Height = Create1.Height - 850
        .Width = Create1.Width - 4200
    End With
        
    Picture1.Scale (-0.02 * Picture1.Width, -0.04 * Picture1.Height)-(1.02 * Picture1.Width, 1.02 * Picture1.Height)
    
    p1 = Picture1.Left + Picture1.Width + 150
    intLabelTop = 60    'set base LabelTop
    
For i = 1 To 13
      
    With Label1(i)
        .Left = p1  'Upper Luff width
        .Width = 1700
        .Top = intLabelTop
        .Height = 375
    End With
    
    intLabelTop = intLabelTop + Label1(i).Height + 20
    
    
Next i
'***************************************************************
' Index 14 - Sail$ - filename not needed to be displayed
'****************************************************************


    With UpLuffScroll
        .Width = 800
        .Left = Text1(1).Left + Text1(1).Width + 20
        .Top = Text1(1).Top
        .Height = 200
        .Tag = "1"
    End With

    With LoLuffScroll
        .Width = 800
        .Left = Text1(1).Left + Text1(1).Width + 20
        .Top = Text1(2).Top
        .Height = 200
        .Tag = "2"
    End With

    '--- right
    With LoLeechScroll
        .Width = 800
        .Left = Text1(1).Left + Text1(1).Width + 20
        .Top = Text1(3).Top
        .Height = 200
        .Tag = "3"
    End With

    '--- lower
    With LBattenScroll
        .Width = 800
        .Left = Text1(1).Left + Text1(1).Width + 20
        .Top = Text1(4).Top
        .Height = 200
        .Tag = "4"
    End With


    '--- upper
    With LYardScroll
        .Width = 800
        .Left = Text1(1).Left + Text1(1).Width + 20
        .Top = Text1(5).Top
        .Height = 200
        .Tag = "5"
    End With

    With FootAScroll
        .Width = 800
        .Left = Text1(1).Left + Text1(1).Width + 20
        .Top = Text1(6).Top
        .Height = 200
        .Tag = "6"
    End With

    '--- left
    With YardAScroll
        .Width = 800
        .Left = Text1(1).Left + Text1(1).Width + 20
        .Top = Text1(7).Top
        .Height = 200
        .Tag = "7"
    End With


    With DepthScroll
        .Width = 800
        .Left = Text1(1).Left + Text1(1).Width + 20
        .Top = Text1(9).Top
        .Height = 200
        .Tag = "9"
    End With

With RPdepthScroll
        .Width = 800
        .Left = Text1(1).Left + Text1(1).Width + 20
        .Top = Text1(10).Top
        .Height = 200
        .Tag = "10"
End With


With TwistScroll        'Leech twist
        .Width = 800
        .Left = Text1(1).Left + Text1(1).Width + 20
        .Top = Text1(11).Top
        .Height = 200
        .Tag = "11"
End With


With NHpanelScroll          ' number of head panels
        .Width = 800
        .Left = Text1(1).Left + Text1(1).Width + 20
        .Top = Text1(12).Top
        .Height = 200
        .Tag = "12"
End With



With NBpanelScroll      ' number of lower panels
        .Width = 800
        .Left = Text1(1).Left + Text1(1).Width + 20
        .Top = Text1(13).Top
        .Height = 200
        .Tag = "13"
End With

If SailStyle = 1 Then       ' if van lorn style
    With Label1(14)
        .Left = p1  'Upper Luff width
        .Width = 1700
        .Top = intLabelTop
        .Height = 375
        .Visible = True
        .Caption = "Foot / Luff :"
    End With
Else
    Label1(14).Visible = False
End If
    

With Develop
    .Left = Label1(13).Left
    .Width = 1550
    .Height = 450
    .Top = Picture1.Top + Picture1.Height - 550
End With

With Save
    .Left = Develop.Left + Develop.Width + 50
    .Width = 1050
    .Height = Develop.Height
    .Top = Develop.Top
End With
    
With Cancel
    .Left = Save.Left + Save.Width + 50
    .Width = 1000
    .Height = Develop.Height
    .Top = Develop.Top
End With

If Sail$ <> "NEW_JUNK" Then
    Create1.Caption = titre$ & Lang(9, Langue) & Sail$
Else
    Create1.Caption = titre$ & Lang(10, Langue)
End If

SailStyleChange
ChangeUnitType
    
redraw

End Sub ' Form_Resize ------------------------------------

Private Sub Form_Unload(Cancel As Integer)
'*********************************************************************************
'9 Nov 2004 - SBS- Added saving setting to INI file
'*********************************************************************************
    IniKeysSet
    Unload Me
    End

End Sub '-------------------------------------------------

Private Sub fprint_Click()
    
    MousePointer = 11 'change pointer to hour glass
    
    MousePointer = 0

End Sub  '-------------------------------------------------

Private Sub fprint_data_Click()
'*********************************************************************************************************
'1 Novenber 2004- sbs -Changed to allow printing Imperial as well as metric.
'Depends on current UnitType.
'**********************************************************************************************************

    Printer.Print Space(6); " ---------------"
    Printer.Print
    Printer.Print Space(6); titre$, " - Sail: "; Sail
    Printer.Print

    Printer.Print Space(6); "Batten length = "; UnitConvert(SCVars(4), UnitType) 'LBattenScroll.Value
    Printer.Print Space(6); "Yard length   = "; UnitConvert(SCVars(5), UnitType) 'LYardScroll.Value
    Printer.Print Space(6); "Foot Angle    ="; SCVars(6); Chr(176) 'FootAScroll.Value; "deg"
    Printer.Print Space(6); "Yard Angle    ="; SCVars(7); Chr(176) 'YardAScroll.Value; "deg"
    Printer.Print
    Printer.Print Space(6); "Head panel    = "; SCVars(12)    'NHpanelScroll.Value
    Printer.Print Space(6); "Lower panels  = "; SCVars(13)    'NBpanelScroll.Value
    Printer.Print Space(8); "(lower panels are split in 2 for development)"
    Printer.Print

    Printer.Print Space(6); "Area  = "; UnitConvertA(SCVars(8), UnitType); UnitCaption(8, UnitType)
    Printer.Print
    Printer.Print Space(6); "Upper panels luff width  = "; UnitConvert(SCVars(1), UnitType)
    Printer.Print Space(6); "Lower panels luff width  = "; UnitConvert(SCVars(2), UnitType)  'LoLuffScroll.Value
    Printer.Print Space(6); "Lower panels leech width = "; UnitConvert(SCVars(3), UnitType)  'LoLeechScroll.Value
    Printer.Print
    Printer.Print Space(6); "Lower panels depth ="; SCVars(9); "%"  'DepthScroll.Value; "%"
    Printer.Print Space(6); "    Depth position ="; SCVars(10); "%"   'RPdepthScroll.Value; "%"
    Printer.Print Space(6); "             Twist = "; SCVars(11); "%"   'TwistScroll.Value
    Printer.Print

    Printer.Print Space(6); " ---------------"
    Printer.EndDoc
End Sub    '-----------------------------------------------

Private Sub fprint_XYZ_Click()
'*******************************************************************************************************
'1 Novenber 2004- sbs -Changed to allow printing Imperial as well as metric.
'Depends on current UnitType. Also added tabbing to keep everything neat...
'*******************************************************************************************************

    Dim n%, i%, j%
    
    j = UnitType

    'Const fmt2$ = "####0"

    For n = 1 To nBpanel + nHpanel
        Printer.Print
        If n > nBpanel Then
            Printer.Print Space(6); titre$, " - Sail: "; Sail, " Head panel  "; n
        Else
            Printer.Print Space(6); titre$; " - Sail: "; Sail, " Lower panel  "; n
        End If
        
        Printer.Print
        Printer.Print Tab; "X lower edge", "Y lower edge", Tab; "X upper edge", "Y upper edge"
        Printer.Print
    
        For i = 0 To 20 Step 2
            Printer.Print Tab; UnitConvert(CStr(plx(n, i)), j), Tab; UnitConvert(CStr(ply(n, i)), j), Tab; UnitConvert(CStr(pmx(n, i)), j), Tab; UnitConvert(CStr(pmy(n, i)), j)
        Next i
        
        Printer.Print
        Printer.Print " ", "X Mid luff", Tab; "Y Mid luff", Tab; "X Mid leech", Tab; "Y Mid leech"
        Printer.Print
        Printer.Print " ", UnitConvert(CStr(pcx(n, 0)), j), Tab; UnitConvert(CStr(pcy(n, 0)), j), Tab; UnitConvert(CStr(pcx(n, 20)), j), Tab; UnitConvert(CStr(pcy(n, 20)), j)
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

    'Menu1.Show
    Cancel_Click

End Sub '--------------------------------------------------


Private Sub fsave_data_Click()

    Dim fichier$
    '-----
    UpLuffScrl = UpLuffScroll.value
    LoLuffScrl = LoLuffScroll.value
    LoLeechScrl = LoLeechScroll.value
    LBattenScrl = LBattenScroll.value
    LyardScrl = LYardScroll.value
    FootAScrl = FootAScroll.value
    YardAScrl = YardAScroll.value

    Mdepth(0) = DepthScroll.value / 100
    RPdepth = RPdepthScroll.value / 100
    twistScrl = TwistScroll.value
    nHpanel = NHpanelScroll.value
    nBpanel = NBpanelScroll.value
    '-----
    CMDialog1.CancelError = True
    
    On Error GoTo errhandler2
    '-----
    Sail$ = Sail$
    CMDialog1.Filter = "SAILCUT Files|*.sc8"
    CMDialog1.filename = Sail$ + ".sc8"
    CMDialog1.Action = 2
    Sail$ = CMDialog1.FileTitle
    Sail$ = Left$(Sail$, (Len(Sail$) - 4))
    fichier$ = CMDialog1.filename
    '-----
    Call F_Ecrire(fichier$)
    infoDirty = False
    SCVars(14) = Sail$
    Form_Resize
    IniKeysSet

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
    Dim c%, i%, p%, j%
    '-----
    CMDialog1.CancelError = True
    
    On Error GoTo errSailDXF
    '-----
    j = UnitType
    If j = 3 Then j = 2     ' no mixed ft - inch
    Sail$ = Sail$
    CMDialog1.Filter = "Sail DXF|*.DXF"
    CMDialog1.filename = Sail$ + ".DXF"
    CMDialog1.Action = 2
    Sail$ = CMDialog1.FileTitle
    Sail$ = Left$(Sail$, (Len(Sail$) - 4))
    fichier$ = CMDialog1.filename
    '-----
    Create1.MousePointer = 11
    Open fichier$ For Output As #1
        DXFHeader "junk rig sail " + Sail
        DXFSectionHeaderGeometry
    
        p = 0
        DXFPolyline 1, 7
        For i = 0 To 20
            DXFVertex 1, 7, UnitConvert(CStr(px(p, i)), j), UnitConvert(CStr(py(p, i)), j), UnitConvert(CStr(pz(p, i)), j)
        Next i
        DXFSequenceEnd
        
    For p = 1 To 2 * (nBpanel + nHpanel)
        DXFPolyline 1, 7
        DXFVertex 1, 7, UnitConvert(CStr(px(p - 1, 0)), j), UnitConvert(CStr(py(p - 1, 0)), j), UnitConvert(CStr(pz(p - 1, 0)), j)
        DXFVertex 1, 7, UnitConvert(CStr(px(p, 0)), j), UnitConvert(CStr(py(p, 0)), j), UnitConvert(CStr(pz(p, 0)), j)
        DXFSequenceEnd
        
        c = 7 - p Mod 2
        DXFPolyline 1, c
        For i = 0 To 20
            DXFVertex 1, c, UnitConvert(CStr(px(p, i)), j), UnitConvert(CStr(py(p, i)), j), UnitConvert(CStr(pz(p, i)), j)
        Next i
        DXFSequenceEnd
        
        DXFPolyline 1, 7
        DXFVertex 1, 7, UnitConvert(CStr(px(p, 20)), j), UnitConvert(CStr(py(p, 20)), j), UnitConvert(CStr(pz(p, 20)), j)
        DXFVertex 1, 7, UnitConvert(CStr(px(p - 1, 20)), j), UnitConvert(CStr(py(p - 1, 20)), j), UnitConvert(CStr(pz(p - 1, 20)), j)
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
    Sail$ = Sail$
    CMDialog1.Filter = "Panels DXF|*.DXF"
    CMDialog1.filename = Sail$ + ".DXF"
    CMDialog1.Action = 2
    Sail$ = CMDialog1.FileTitle
    Sail$ = Left$(Sail$, (Len(Sail$) - 4))
    fichier$ = CMDialog1.filename
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
    ' saving sail in VRML 1.0 file

Dim fichier$
'-----
CMDialog1.CancelError = True
                                          
On Error GoTo errhandvrml1
    '-----
    Sail$ = Sail$
    CMDialog1.Filter = "VRML Files|*.wrl"
    CMDialog1.filename = Sail$ + ".wrl"
    CMDialog1.Action = 2
    Sail$ = CMDialog1.FileTitle
    Sail$ = Left$(Sail$, (Len(Sail$) - 4))
    fichier$ = CMDialog1.filename
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
            Sail$ = Sail$
            CMDialog1.Filter = "Panels X-Y|*.XYZ"
            CMDialog1.filename = Sail$ + ".XYZ"
            CMDialog1.Action = 2
            Sail$ = CMDialog1.FileTitle
            Sail$ = Left$(Sail$, (Len(Sail$) - 4))
            fichier$ = CMDialog1.filename
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
    
    LBatten = LBattenScroll.value
    If SailStyle = 1 Then LYardScroll.value = LBatten
    infoDirty = True

    If Me.Visible And Not infoRead And Not infoRead Then form_activate

End Sub '-------------------------------------------------

Private Sub LoLeechScroll_Change()

    infoDirty = True

    If LoLeechScroll.value < LoLuffScroll.value Then
        LoLeechScroll.value = LoLuffScroll.value
    End If

    LoLeech = LoLeechScroll.value

    If Me.Visible And Not infoRead Then form_activate

End Sub ' ------------------------------------------------

Private Sub LoLuffScroll_Change()
        
        infoDirty = True
    Select Case SailStyle
    Case 0
        If LoLuffScroll.value > LoLeechScroll.value Then
            LoLuffScroll.value = LoLeechScroll.value
        End If
    Case 1          'Van Lorn
            If LoLuffScroll.value <> LoLeechScroll.value Then
            LoLeechScroll.value = LoLuffScroll.value
        End If
    End Select
    
        LoLuff = LoLuffScroll.value

        If Me.Visible And Not infoRead Then form_activate

End Sub  '------------------------------------------------

Private Sub LYardScroll_Change()
    
    infoDirty = True
    If LYardScroll.value < 0.5 * LBattenScroll.value Then
        LYardScroll.value = 0.5 * LBattenScroll.value
    
    ElseIf LYardScroll.value > 1.8 * LBattenScroll.value Then
        LYardScroll.value = 1.8 * LBattenScroll.value
    End If

    LYard = LYardScroll.value

    If Me.Visible And Not infoRead Then form_activate

End Sub    '----------------------------------------------

Private Sub mnuSailStyleStd_Click()
    SailStyle = 0
    mnuSailStyleVanLorn.Checked = False
    mnuSailStyleStd.Checked = True
    defaut
    Form_Resize
End Sub

Private Sub mnuSailStyleVanLorn_Click()
    SailStyle = 1
    mnuSailStyleVanLorn.Checked = True
    mnuSailStyleStd.Checked = False
    VanLornDefault
    Form_Resize
End Sub

Private Sub NBpanelScroll_Change()
    
    infoDirty = True
    Label1(13).BackColor = vbWindowBackground
    nBpanel = NBpanelScroll.value
    
    If Me.Visible And Not infoRead Then form_activate

End Sub ' NBpanelScroll_Change ----------------------------

Private Sub NHpanelScroll_Change()
    
    infoDirty = True
    nHpanel = NHpanelScroll.value
    
    If Me.Visible And Not infoRead Then form_activate

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
    Call CheckFlags
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

    infoDirty = True
    RPdepth = RPdepthScroll.value / 100

    If Me.Visible And Not infoRead Then form_activate

End Sub  ' -----------------------------------------------


Private Sub Save_Click()
fsave_data_Click

End Sub


Private Sub TwistScroll_Change()
    ' vrillage
    
    infoDirty = True
    twistScrl = TwistScroll.value
    Atwist = twistScrl / 57
    '--
    If Me.Visible And Not infoRead Then form_activate

End Sub '-------------------------------------------------

Private Sub UpLuffScroll_Change()
        
    infoDirty = True
    If UpLuffScroll.value > LoLeechScroll.value Then
        UpLuffScroll.value = LoLeechScroll.value
    End If

    UpLuff = UpLuffScroll.value
    
    If Me.Visible And Not infoRead Then form_activate

End Sub  '------------------------------------------------

Private Sub YardAScroll_Change()
    
    infoDirty = True
    If YardAScroll.value < FootAScroll.value Then
        YardAScroll.value = FootAScroll.value
    ElseIf YardAScroll.value < AHead * RAD And SailStyle = 0 Then
        YardAScroll.value = AHead * RAD + 0.5
    End If
    
    AYard = YardAScroll.value

    If Me.Visible And Not infoRead Then form_activate

End Sub  '------------------------------------------------

Private Sub ChangeUnitType()
Dim i%, j%

'**********************************************************************
'added conversion display for mm to imperial 31 October 2004
'**********************************************************************
j = UnitType


For i = 1 To 14
    
    Select Case i
        Case 1 To 5

             With Text1(i)
                .Left = Label1(i).Left + Label1(i).Width + 50
                .Top = Label1(i).Top
                .Height = Label1(i).Height / 2
                .Width = 1000
                .Text = UnitConvert(SCVars(i), j)
            End With
            
        Case 6 To 7

             With Text1(i)
                .Left = Label1(i).Left + Label1(i).Width + 50
                .Top = Label1(i).Top
                .Height = Label1(i).Height / 2
                .Width = 1000
                .Text = SCVars(i) & UnitCaption(i, j)
            End With


        Case 8

             With Text1(i)
                .Left = Label1(i).Left + Label1(i).Width + 50
                .Top = Label1(i).Top
                .Height = Label1(i).Height / 2
                .Width = 1000
                .Text = UnitConvertA(SCVars(i), j) & UnitCaption(i, j)
            End With
        
        Case 9 To 13

             With Text1(i)
                .Left = Label1(i).Left + Label1(i).Width + 50
                .Top = Label1(i).Top
                .Height = Label1(i).Height / 2
                .Width = 1000
                .Text = SCVars(i) & UnitCaption(i, j)
            End With
        
        Case 14
            
            If SailStyle = 1 Then
                With Text1(i)
                    .Left = Label1(i).Left + Label1(i).Width + 50
                    .Top = Label1(i).Top
                    .Height = Label1(i).Height / 2
                    .Width = 1000
                    .Visible = True
                    .Text = Format(LBatten / ((nBpanel + 1) * LoLuff), "#0.0##")
                End With
            Else
                Text1(i).Visible = False
            End If
            

    End Select

Next i


End Sub 'ChangeUnitType-----------------------------------------------

Private Sub SailStyleChange()
   
   If SailStyle = 1 Then
        Label1(1).Visible = False
        Text1(1).Visible = False
        UpLuffScroll.Visible = False
        
        If LoLeechScroll.value <> LoLuffScroll.value Then LoLeechScroll.value = LoLuffScroll.value
        LoLeechScroll.Enabled = False
        
        If LYardScroll.value <> LBattenScroll.value Then LYardScroll.value = LBattenScroll.value
        LYardScroll.Enabled = False
        
        NHpanelScroll.Enabled = False
        
   Else
        Label1(1).Visible = True
        Text1(1).Visible = True
        UpLuffScroll.Visible = True

        LoLeechScroll.value = LoLeech
        LoLeechScroll.Enabled = True

        LYardScroll.value = LYard
        LYardScroll.Enabled = True
        
        nHpanel = 2
        NHpanelScroll.value = nHpanel
        NHpanelScroll.Enabled = True
        


    End If
   
   
End Sub

