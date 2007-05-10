VERSION 5.00
Begin VB.Form Cloth1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Sailcut8 Cloth data"
   ClientHeight    =   5310
   ClientLeft      =   2355
   ClientTop       =   2070
   ClientWidth     =   7050
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
   Icon            =   "cloth1.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5310
   ScaleWidth      =   7050
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Done1 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   372
      Left            =   1560
      TabIndex        =   7
      Top             =   3240
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   15
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   14
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1(2)"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1(1)"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label2(2)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   5
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label2(1)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label2(0)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1(0)"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Cloth1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit  ' 17 October 2004
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



Private Sub chkWidth()
    
    If SeamW < 20 Then
        SeamW = 20
    ElseIf SeamW > 150 Then
        SeamW = 150
    End If

    If SeamT < 0 Then
        SeamT = 0
    ElseIf SeamT > 60 Then
        SeamT = 60
    End If

    Text1(0).Text = " N/A "
    Text1(1).Text = Str$(SeamW)
    Text1(2).Text = Str$(SeamT)

End Sub ' chkWidth ---------------------------------------

Private Sub Done1_Click()
    
    Cloth1.Hide

End Sub '-----------------------------------------------

Private Sub form_activate()
    ' cloth1 Activate

    Dim i%

    Text1(0).Text = " N/A "
    Text1(1).Text = Str$(SeamW)
    Text1(2).Text = Str$(SeamT)
    
    For i = 0 To 2
        Label2(i) = "  -  "
        Text2(i).Text = " N/A "
        Label3(i) = "mm"
    Next i

    Select Case langue
    Case 0 'français
        Cloth1.Caption = "Paramètres du tissus"
        Label1(0) = "Largeur du tissus"
        Label1(1) = "Largeur poche de lattes"
        Label1(2) = "Largeur retour de bords"
        'label2(0) = "Couleur 1"
        'label2(1) = "couleur 2"
        'label2(2) = "couleur 3"
        Done1.Caption = "Valider"

    Case 1 'anglais
        Cloth1.Caption = "Cloth parameters"
        Label1(0) = "Cloth width"
        Label1(1) = "Batten pocket width"
        Label1(2) = "Tabling width"
        'label2(0) = "Color 1"
        'label2(1) = "color 2"
        'label2(2) = "color 3"
        Done1.Caption = "Validation"
    
    End Select

End Sub  '------------------------------------------------

Private Sub Form_Deactivate()
    
    ClothW = Val(Text1(0).Text)
    SeamW = Val(Text1(1).Text)
    SeamT = Val(Text1(2).Text)

    chkWidth

End Sub '-------------------------------------------------

Private Sub Text1_KeyPress(Index As Integer, keyascii As Integer)

    If keyascii = 13 Then
        chkWidth
    End If

End Sub ' Text1_KeyPress ---------------------------------

