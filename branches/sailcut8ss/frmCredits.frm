VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4470
   ClientLeft      =   2355
   ClientTop       =   1950
   ClientWidth     =   4305
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3094.666
   ScaleMode       =   0  'User
   ScaleWidth      =   4034.96
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1521
      TabIndex        =   0
      Top             =   3978
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   117
      TabIndex        =   1
      Top             =   3978
      Width           =   1245
   End
   Begin VB.Line Line1 
      X1              =   109.661
      X2              =   3184.854
      Y1              =   2669.582
      Y2              =   2669.582
   End
   Begin VB.Image imgIcon0 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   345
      Picture         =   "frmCredits.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Copyright"
      Height          =   169
      Left            =   702
      TabIndex        =   4
      Top             =   936
      Width           =   572
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Left            =   1053
      TabIndex        =   2
      Top             =   234
      Width           =   1729
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Version"
      Height          =   169
      Left            =   702
      TabIndex        =   3
      Top             =   702
      Width           =   455
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' 10 November 2004******************************************************
' Copyright (C) 2004 Steve Studley
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

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Dim strCredits(12, 1) As String
    Dim i As Integer
    Dim j%, k!
    Dim Fm As Form
    Dim ctlLbl As Label
    Dim ctlImg As Image
    Dim strLabelName$
    Dim strImageName$
    Dim strPicFullName$
    
    
    i = Langue
    
    Set Fm = frmCredits

    Fm.Caption = Lang(42, i)
    lblTitle.Caption = App.Title & " - " & Lang(42, Langue)
    
    With imgIcon0
        .Top = 50
        .Left = (Fm.Width - (imgIcon0.Width + lblTitle.Width + 50)) / 2
    End With
    
    With lblTitle
        .Top = imgIcon0.Top
        .Left = imgIcon0.Left + imgIcon0.Width + 50
        .Height = 350
    End With
    
    With lblVersion
        .Caption = "Version - " & version
        .Top = lblTitle.Top + lblTitle.Height + 50
        .Left = Fm.Width \ 2 - .Width \ 2 - 150
    End With
    
    With lblCopyright
        .Caption = copyright
        .Top = lblVersion.Top + lblVersion.Height + 50
        .Left = Fm.Width \ 2 - .Width \ 2 - 150
    End With



'create array for labels

strCredits(1, 0) = Lang(43, i)        ' Developers
strCredits(2, 0) = "  Robert Laine"
strCredits(2, 1) = " france.gif"
strCredits(3, 0) = "  Steve Studley"
strCredits(3, 1) = "usa.gif"
strCredits(4, 0) = " "
strCredits(5, 0) = Lang(44, i)      'Translations
strCredits(6, 0) = Lang(45, i) & " - Robert Laine"     'French
strCredits(6, 1) = "france.gif"
strCredits(7, 0) = Lang(46, i) & " - Robert Laine"     ' English
strCredits(7, 1) = "uk.gif"
strCredits(8, 0) = Lang(48, i) & " - Tony Mels"          'Dutch
strCredits(8, 1) = "holland.gif"
strCredits(9, 0) = Lang(49, i) & " - Leo Foltz"               'German
strCredits(9, 1) = "german.gif"
strCredits(10, 0) = Lang(51, i) & " - Terho Halme"       'Finnish
strCredits(10, 1) = "finland.gif"
strCredits(11, 0) = Lang(47, i) & " - Rolf Nilsen"          'Norwegian
strCredits(11, 1) = "norway.gif"
strCredits(12, 0) = Lang(50, i) & " - Joserra Mariño"   'Spanish
strCredits(12, 1) = "spain.gif"

'create label and images
k = lblCopyright.Top + lblCopyright.Height + 50


For j = 1 To UBound(strCredits)

    strLabelName = "lblCredits" & j
    strImageName = "imgIcon" & j

    

    Set ctlLbl = Controls.Add("VB.Label", strLabelName)
    
    With ctlLbl
        .Caption = strCredits(j, 0)
        .Tag = strCredits(j, 1)
        .BackColor = vbWindowBackground
        .Height = 500
        .AutoSize = True
        .Visible = True
    End With
    
    
If ctlLbl.Tag <> "" Then
    strPicFullName = App.Path & "\" & Trim(ctlLbl.Tag)

    Set ctlImg = Controls.Add("VB.Image", strImageName)
    With ctlImg
        .Picture = LoadPicture(strPicFullName)
        .Left = imgIcon0.Left
        .Top = k + (ctlLbl.Height + 110) * j
        .Visible = True
    End With
    
    With ctlLbl
        .Left = imgIcon0.Left + ctlImg.Width + 50
        .Top = ctlImg.Top + 50
    End With
                         
        
Else

      With ctlLbl
        .Left = imgIcon0.Left
        .Top = k + ((ctlLbl.Height + 50) * j)
    End With
        
End If
        
Next j
                        
With Line1
    .x1 = 10
    .y1 = ctlLbl.Top + ctlLbl.Height + 100
    .x2 = Fm.Width - 20
    .y2 = .y1
End With

With cmdSysInfo
    .Left = (Fm.Width - (cmdSysInfo.Width + cmdOK.Width + 200)) / 2 - 200
    .Top = ctlLbl.Top + ctlLbl.Height + 200
End With
With cmdOK
    .Left = cmdSysInfo.Left + cmdSysInfo.Width + 200
    .Top = cmdSysInfo.Top
End With

 
Fm.Height = cmdSysInfo.Top + cmdSysInfo.Height + 2400
   
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


