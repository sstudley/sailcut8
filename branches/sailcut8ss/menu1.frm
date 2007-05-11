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
   Begin VB.ListBox lstUnitType 
      Height          =   840
      Left            =   3960
      TabIndex        =   8
      Top             =   3120
      Width           =   1050
   End
   Begin VB.CommandButton cmdCredits 
      Appearance      =   0  'Flat
      Caption         =   " Credits"
      Height          =   364
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   1905
   End
   Begin VB.ListBox Language 
      Height          =   840
      Left            =   2808
      TabIndex        =   6
      Top             =   3150
      Width           =   1050
   End
   Begin VB.CommandButton cmdMailto 
      Appearance      =   0  'Flat
      Caption         =   "E-mail: robert.laine@sailcut.com"
      CausesValidation=   0   'False
      Height          =   364
      Left            =   240
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   2938
   End
   Begin VB.CommandButton Terminate 
      Appearance      =   0  'Flat
      Caption         =   "TERMINATE and EXIT"
      Height          =   492
      Left            =   2760
      TabIndex        =   2
      Top             =   2520
      Width           =   2532
   End
   Begin VB.CommandButton Create 
      Appearance      =   0  'Flat
      Caption         =   "CREATE A NEW SAIL"
      Height          =   492
      Left            =   2760
      TabIndex        =   1
      Top             =   1920
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
      Top             =   1920
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


Option Explicit
' 10 November 2004******************************************************
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


'******************************
'start declares for sending email
'******************************
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Const SW_SHOWNORMAL = 1
Const SW_MAXIMIZE = 3

Const SE_ERR_FNF = 2&
Const SE_ERR_PNF = 3&
Const SE_ERR_ACCESSDENIED = 5&
Const SE_ERR_OOM = 8&
Const SE_ERR_DLLNOTFOUND = 32&
Const SE_ERR_SHARE = 26&
Const SE_ERR_ASSOCINCOMPLETE = 27&
Const SE_ERR_DDETIMEOUT = 28&
Const SE_ERR_DDEFAIL = 29&
Const SE_ERR_DDEBUSY = 30&
Const SE_ERR_NOASSOC = 31&
Const ERROR_BAD_FORMAT = 11&
'******************************
'end declares for sending email
'******************************

Private Sub cmdCredits_Click()

    frmCredits.Show
    
End Sub

'********************************
'send email     5/10/2004
'*********************************

Private Sub cmdMailto_Click()


On Error GoTo err_mailto
    Dim lRet As Long, msg As String
'Mailto: robert.laine@sailcut.com,sstudley@verizon.net?subject=SailCut8"

    lRet = ShellExecute(Me.hwnd, "Open", _
    "Mailto: robert.laine@sailcut.com,sstudley@verizon.net?subject=SailCut8", _
    vbNullString, _
    vbNullString, _
    SW_SHOWNORMAL)

err_mailto:

          If lRet <= 32 Then
              'There was an error
              Select Case lRet
                  Case SE_ERR_FNF
                      msg = "File not found"
                  Case SE_ERR_PNF
                      msg = "Path not found"
                  Case SE_ERR_ACCESSDENIED
                      msg = "Access denied"
                  Case SE_ERR_OOM
                      msg = "Out of memory"
                  Case SE_ERR_DLLNOTFOUND
                      msg = "DLL not found"
                  Case SE_ERR_SHARE
                      msg = "A sharing violation occurred"
                  Case SE_ERR_ASSOCINCOMPLETE
                      msg = "Incomplete or invalid file association"
                  Case SE_ERR_DDETIMEOUT
                      msg = "DDE Time out"
                  Case SE_ERR_DDEFAIL
                      msg = "DDE transaction failed"
                  Case SE_ERR_DDEBUSY
                      msg = "DDE busy"
                  Case SE_ERR_NOASSOC
                      msg = "No association for file extension"
                  Case ERROR_BAD_FORMAT
                      msg = "Invalid EXE file or error in EXE image"
                  Case Else
                      msg = "Unknown error"
              End Select
              MsgBox msg, vbInformation, "Email Error"

          End If



    Exit Sub

End Sub


Private Sub Create_Click()

    MousePointer = 11  'change pointer to hourglass
    Create1.Show
    Me.Hide
    
    infoDirty = False

End Sub '--------------------------------------------------

Private Sub form_activate()
    
    'Create1.MousePointer = 0
    Menu1.MousePointer = 0

End Sub '--------------------------------------------------

Private Sub Form_Initialize()
Dim file$


file$ = AppIniPathGet

    mLang_LoadArray
    mCtlNames_LoadArray
    SailStyle = 0   ' to set std
'******************************************
'Added 9 Nov 2004 SBS
'Load INI file
'*******************************************
If Dir$(file$) <> "" Then
    Langue = KeyValueGet("App", "Language")
    UnitType = KeyValueGet("App", "UnitType")
    WinHeight = Val(KeyValueGet("Window", "Height"))
    WinWidth = Val(KeyValueGet("Window", "Width"))
    WinTop = Val(KeyValueGet("Window", "Top"))
    WinLeft = Val(KeyValueGet("Window", "Left"))

    
Else
    IniKeysWrite "App", "Language", "1"
    IniKeysWrite "App", "UnitType", 0
End If
    
    
    
    

End Sub

Private Sub Form_Load()

    timenow = Now
    
    Cls
    With Menu1
        .Visible = False
        .Width = 6000
        .Height = 6000
        .Top = (Screen.Height - Menu1.Height) / 2
        .Left = (Screen.Width - Menu1.Width) / 2
        .Caption = titre + "       Version :" + version
    End With
    
    With label1
        .Top = 200
        .Height = 300
        .Width = Menu1.Width * 0.9
        .Left = (Menu1.Width - label1.Width) / 2
    End With
    
    With Label2
        .Top = label1.Top + label1.Height + 100
        .Height = 1600
        .Width = label1.Width
        .Left = label1.Left
    End With
    
    '*********************************************
    ' cmdMailto added to send mail with default mail app May 10,2004
    '***********************************************
    With cmdMailto
        .Top = Label2.Top + Label2.Height
        .Height = 500
        .Width = label1.Width / 2
        .Left = label1.Left
        .BackColor = Label2.BackColor
    End With
    
    With cmdCredits
        .Top = cmdMailto.Top
        .Height = 500
        .Width = label1.Width / 2
        .Left = cmdMailto.Left + cmdMailto.Width
        .BackColor = Label2.BackColor
    End With
    
    
    Picture1.Top = Label2.Top + Label2.Height + 700
    Picture1.Left = Label2.Left
    
    With Create
        .Top = Picture1.Top
        .Height = 400
        .Left = Menu1.Width / 2 + 20
        .Width = 1300
    End With

    With Terminate
        .Top = Create.Top
        .Height = 400
        .Left = Create.Left + Create.Width + 150
        .Width = Create.Width
    End With
    

'******************************
'loads language variable strings from string array Lang()
'depending on Langue1.selected( index)
'can add a new language by adding a new .additem in Langue and
'a new string index in mdlLang module
'*****************************
    
    With Language
        .Left = Create.Left
        .Top = Create.Top + Create.Height + 100
        .Width = Create.Width
        .Height = Menu1.Height - .Top - 540
        'add languages
        .AddItem "français"
        .AddItem "English"
        .AddItem "norsk"
        .AddItem "Nederlands"
        .AddItem "Deutsch"
        .AddItem "Español"
        .AddItem "suomi"
        .Selected(Langue) = True
    End With
    
        With lstUnitType
        .Top = Language.Top
        .Height = Language.Height
        .Left = Terminate.Left
        .Width = Terminate.Width
        .AddItem "Metric - mm"
        .AddItem "Inches"
        .AddItem "Feet"
        .AddItem "Feet - Inches"
        .Selected(UnitType) = True
        
    End With
    
    intLangue = Language.ListCount
    'Langue.Height = (intLangue * 175) + 150
           
    'Langue = 1

    '---
    Call LOGO(75, vb3DShadow)
    Call LOGO(0, vbBlack)
    
    '------
    Menu1.Show
    '--- load default variables
    LoadVariables
    LoadLabelCaps

End Sub '-------------------------------------------------

Private Sub Form_Unload(Cancel As Integer)
    
    'IniKeysSet
    Unload Me
    End

End Sub '-------------------------------------------------

Private Sub LOGO(i As Integer, lngColor As Long)
' logo RL

With Picture1
    .Width = Menu1.Width / 2.35
    .Height = .Width
    .DrawWidth = 6
    .ForeColor = lngColor
End With

    ks = Picture1.Width / 1500
    
'draw rl

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
    
End Sub  '----------------------------------------------



Private Sub Language_Click()
Dim i As Integer


If Language.ListIndex > -1 Then
    i = Language.ListIndex
Else
    i = 1
End If

Langue = i
LoadLabelCaps

End Sub

Private Sub LoadLabelCaps()

Dim i, j As Integer
Dim strCap As String

i = Langue
For j = 1 To 5
    strCap = strCap & Lang(j, i) & Chr$(13)
Next j

    label1.Caption = Lang(0, i)
    Label2.Caption = strCap
    Create.Caption = Lang(6, i)
    Terminate.Caption = Lang(7, i)
    cmdCredits.Caption = Lang(42, i)
    cmdMailto.Caption = Lang(8, i)
    


End Sub


Private Sub Terminate_Click()

    Menu1.Hide
    Unload Me
    
End Sub ' Terminate_click --------------------------------


Private Sub LoadVariables()


    LBatten = 3000
    LYard = 3000
    LoLeech = 1000
    ClothW = 900
    SeamW = 25


    
    Sail = "NEW_JUNK"
    AFoot = 15
    AYard = 70
    nHpanel = 2
    nBpanel = 4
    RPdepth = 0.4
    

    
End Sub


Private Sub lstUnitType_Click()
Dim i As Integer


If lstUnitType.ListIndex > -1 Then
    i = lstUnitType.ListIndex
Else
    i = 0
End If

UnitType = i
End Sub
