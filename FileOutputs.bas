Attribute VB_Name = "FileOutputs"


Option Explicit
' 4 May 2007******************************************************
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


Sub F_DXF(file$)
    Dim p%, i%, j%
    Dim fichier$
    Dim genre$
    Const fmt2$ = "0.00 "
    '-----------
    'change pointer to hourglass
    Create1.MousePointer = 11

    On Error GoTo errF_DXF
    
    j = UnitType
    If j = 3 Then j = 2     ' no mixed ft - inch


    For p = 1 To nBpanel + nHpanel
        fichier = Left$(file, Len(file) - 4) + Format$(p, "00") + ".DXF"

        If p > nBpanel Then
            genre = " Head panel "
        Else
            genre = " Lower panel "
        End If
        
        Open fichier$ For Output As #1
    
        DXFHeader genre + Format$(p, " #0 ")
        DXFSectionHeaderGeometry

        DXFPolyline 1, 7
        For i = 0 To 20
            DXFVertex 1, 7, UnitConvert(CStr(plx(p, i)), j), UnitConvert(CStr(ply(p, i)), j), 0
        Next i
        
            DXFVertex 1, 7, UnitConvert(CStr(pcx(p, 20)), j), UnitConvert(CStr(pcy(p, 20)), j), 0
        
        For i = 20 To 0 Step -1
            DXFVertex 1, 7, UnitConvert(CStr(pmx(p, i)), j), UnitConvert(CStr(pmy(p, i)), j), 0
        Next i
        
            DXFVertex 1, 7, UnitConvert(CStr(pcx(p, 0)), j), UnitConvert(CStr(pcy(p, 0)), j), 0
            DXFVertex 1, 7, UnitConvert(CStr(plx(p, 0)), j), UnitConvert(CStr(ply(p, 0)), j), 0

        DXFSequenceEnd
        DXFSectionEnd
        DXFEnd
        '-----
        Close #1
    Next p
    
    Create1.MousePointer = 0
    Exit Sub
    '------

errF_DXF:
    Close
    Create1.MousePointer = 0
    
    MsgBox Error(Err), 0, "F_DXF"
    
    Exit Sub

End Sub ' F_DXF ------------------------------------------

Sub F_Ecrire(fichier$)
    ' write data file
    Dim i%, p%
    '-------------------
    'change pointer to hourglass
        Create1.MousePointer = 11
    
    On Error GoTo errhandle1
    
    Open fichier$ For Output As #1
        Print #1, titre
        Print #1, Sail
        Print #1, genre1
        Print #1, genre2
        Print #1, Str$(UpLuffScrl)
        Print #1, Str$(LoLuffScrl)
        Print #1, Str$(LoLeechScrl)
        Print #1, Str$(LBattenScrl)
        Print #1, Str$(LyardScrl)
        Print #1, Str$(FootAScrl)
        Print #1, Str$(YardAScrl)
        
        Print #1, Str$(Mdepth(0))
        Print #1, Str$(RPdepth)
        Print #1, Str$(twistScrl)
        Print #1, Str$(nHpanel)
        Print #1, Str$(nBpanel)
        Print #1, Str$(ClothW)
        Print #1, Str$(SeamW)
        Print #1, Str$(SeamT)
        Print #1, "EOF "
    Close #1
    '-- change pointer back to default
        Create1.MousePointer = 0
    Exit Sub
    '------
errhandle1:
    MsgBox Error(Err), 0, "F_ecrire"
    '-- change pointer back to default
    Create1.MousePointer = 0
    Exit Sub

End Sub '---------------------------------- F_ecrire ---

Sub F_Lire(fichier$)
    ' read data file
    Dim i%, p%
    Dim aa#, bb#
    Dim a$, b$
    
    'change pointer to hourglass
    Create1.MousePointer = 11
    
    On Error GoTo errhandler
        
    Open fichier$ For Input As #1
        Input #1, a$
        Input #1, Sail
        Input #1, genre1
        Input #1, genre2
        Input #1, b
            UpLuffScrl = Val(b)
        Input #1, b
            LoLuffScrl = Val(b)
        Input #1, b
            LoLeechScrl = Val(b)
        Input #1, b
            LBattenScrl = Val(b)
        Input #1, b
            LyardScrl = Val(b)
        Input #1, b
            FootAScrl = Val(b)
        Input #1, b
            YardAScrl = Val(b)
        Input #1, b
            Mdepth(0) = Val(b)
        Input #1, b
            RPdepth = Val(b)
        Input #1, b
            twistScrl = Val(b)
        Input #1, b
            nHpanel = Val(b)
        Input #1, b
            nBpanel = Val(b)
        Input #1, b
            ClothW = Val(b)
        Input #1, b
            SeamW = Val(b)
        Input #1, b
            SeamT = Val(b)
        Input #1, b$
    Close #1
    
    '-- change pointer back to default
        Create1.MousePointer = 0
    Exit Sub
    
    '------
errhandler:
    Close
    Screen.MousePointer = 0
    
    MsgBox Error(Err), 0, "F_lire"
    '-- change pointer back to default
    Exit Sub

End Sub  ' F_lire -----------------------------------------

Sub F_VRML1(fichier$)
    'write a VRML file Author Robert Lainé
    ' 14 Nov 2004
    Dim p%, i%
    Dim tt$
    Dim t#, x#, y#, Z#
    Const fmt1$ = "0.0 "
    Const fmt2$ = "0.#0 "
    Const fmt3$ = "0.##0 "
      
    Dim npanel%
    npanel = 2 * (nBpanel + nHpanel)
    '---
    On Error GoTo errhandler_eVRML1

    Open fichier$ For Output As #1
    
    Print #1, "#VRML V1.0 ascii "
    '-----
    Print #1, "  DirectionalLight { direction .1 -.2 1 intensity .6 }"
    Print #1, "  DirectionalLight { direction -.2 .2  -1 intensity .4 }"
    Print #1, "  PerspectiveCamera { "
    Print #1, "    position "; Format$(XClew / 2000, fmt1); Format$(Ypeak / 2000, fmt1); Format$(Ypeak / 600, fmt1)
    Print #1, "    orientation 0 0 1 0 "
    Print #1, "  } " ' camera
    Print #1, "  ShapeHints { "
    'Print #1, "    vertexOrdering UNKNOWN_ORDERING "
    Print #1, "    shapeType      UNKNOWN_SHAPE_TYPE"
    Print #1, "  }"
    
    '-----
    Print #1, "Separator {   # world"
    '-----
    Print #1, "  Separator {  #  XYZ"
    Print #1, "    Material { diffuseColor 1 1 .5 } "
    Print #1, "    Coordinate3 { "
    Print #1, "      point [ "  'store all the points
    For p = 0 To npanel
        For i = 0 To 20
            Print #1, Space(10); Format$(px(p, i) / 1000, fmt3); Format$(py(p, i) / 1000, fmt3); Format$(pz(p, i) / 1000, fmt3); " , "
        Next i
    Next p
    Print #1, "         ] "
    Print #1, "    }  # coord"
        '-----
    Print #1, "    IndexedFaceSet { "
    Print #1, "      coordIndex [ "  'draw faces
    For p = 0 To npanel - 1
        For i = 0 To 19
            ' first triangle
            Print #1, Space(10); Str$(i + p * 21) + " , " + Str$(i + 1 + p * 21) + " , " + Str$(i + (p + 1) * 21) + " , -1 , "
            ' second triangle
            Print #1, Space(10); Str$(i + 1 + p * 21) + " , " + Str$(i + 1 + (p + 1) * 21) + " , " + Str$(i + (p + 1) * 21) + " , -1 , "
        Next i
    Next p
    Print #1, "      ] "
    Print #1, "    }  # FaceSet"
        '-----
    Print #1, "    IndexedLineSet { "
    Print #1, "      coordIndex [ " 'draw panels edge lines
    For p = 0 To npanel Step 2
        If p > 1 Then
            Print #1, Space(10); Str$(0 + (p - 2) * 21) + " , "
        End If
        
        For i = 0 To 20
            Print #1, Space(10); Str$(i + p * 21) + " , "
        Next i
        
        If p > 1 Then
            Print #1, Space(10); Str$(20 + (p - 2) * 21) + " , "
        End If
        
        Print #1, Space(10); "-1 , "
    Next p
    Print #1, "       ] "
    Print #1, "    }  # LineSet"
        '------
    Print #1, "  }  #  end XYZ"
    '-----
    Print #1, "} # end world"

    '----------------------- Title
    Print #1, "WWWAnchor { name ""http://www.sailcut.com/"" map NONE "
            t = Ypeak / 8000
    Print #1, "  FontStyle { size "; Format$(t, fmt3); " } "
    Print #1, "  Material { "
    Print #1, "     diffuseColor    .2 .5 .5 "
    Print #1, "     emissiveColor   .2 .5 .5 "
    Print #1, "  } "  ' end material
    '-------
    Print #1, "  Separator { " 'texte
            x = Xpeak / 2000
            y = 1.05 * Ypeak / 1000
            Z = -0.1
    Print #1, "     Transform { "
    Print #1, "        translation "; Format$(x, fmt3); Format$(y, fmt3); Format$(Z, fmt3)
    Print #1, "        rotation 1 0 0  -.5 "
    Print #1, "     } "  ' end transform

    Print #1, "     AsciiText { "
    Print #1, "        string    ""Sailcut8"""
    Print #1, "        justification  CENTER "
    Print #1, "     } "  ' end asciiText
    Print #1, "  } "  ' end text
    Print #1, "}  # end WWWanchor"
    Print #1, "  "
    '-----
    Close #1
    Exit Sub
    '-----
errhandler_eVRML1:
    Close #1
    tt$ = titre & " : SUB eVRML1"
    If Err = 75 Then
        MsgBox "Trying to write over a Read Only File or an invalid Path", 0, tt$
    Else
        MsgBox Error(Err), 0, tt$
    End If
    Exit Sub
    
 
End Sub ' F_VRML1 ----------------------------

Sub F_XY(fichier$)
    ' write file XY developed panels
    Dim p%, i%, j%, d%
    Dim unit$
    
    Dim genre$
    Const fmt2$ = "0.00 "
    j = UnitType
    'd = Precision for UnitConvert
    '-------------------
    'change pointer to hourglass
    Create1.MousePointer = 11

    On Error GoTo errhandle4

    Open fichier$ For Output As #1
    Width #1, 80

   Print #1, titre, " Panels development of " & Sail
   
   Select Case j
   Case 0
        unit = " mm"
        d = 2
    Case 1
        unit = " Inches"
        d = 3
    Case 2
        unit = " Feet"
        d = 4
    Case 3
        unit = " Feet - Inches"
        d = 3
    End Select
    
    Print #1, "Units are in" & unit
    
    For p = 1 To nBpanel + nHpanel
        If p > nBpanel Then
            genre = " Head panel "
        Else
            genre = " Lower panel "
        End If
        
        Print #1,
        Print #1, Format$(p, " #0 ") + Chr$(9) + genre + Chr$(9) + "lower edge"
        Print #1, "    X " + Chr$(9) + "    Y"

        For i = 0 To 20
            Print #1, UnitConvert(CStr(plx(p, i)), j, d, False) + Chr$(9) + UnitConvert(CStr(ply(p, i)), j, d, False)
        Next i
        
        Print #1,
        Print #1, Format$(p, " #0 ") + Chr$(9) + genre + Chr$(9) + "mid luff"
        Print #1, "    X " + Chr$(9) + "    Y"
        Print #1, UnitConvert(CStr(pcx(p, 0)), j, d, False) + Chr$(9) + UnitConvert(CStr(pcy(p, 0)), j, d, False)
        
        Print #1,
        Print #1, Format$(p, " #0 ") + Chr$(9) + genre + Chr$(9) + "upper edge"
        Print #1, "    X " + Chr$(9) + "    Y"

        For i = 0 To 20
            Print #1, UnitConvert(CStr(pmx(p, i)), j, d, False) + Chr$(9) + UnitConvert(CStr(pmy(p, i)), j, d, False)
        Next i
        
        Print #1,
        Print #1, Format$(p, " #0 ") + Chr$(9) + genre + Chr$(9) + "mid leech"
        Print #1, "    X " + Chr$(9) + "    Y"
        Print #1, UnitConvert(CStr(pcx(p, 20)), j, d, False) + Chr$(9) + UnitConvert(CStr(pcy(p, 20)), j, d, False)
        '-----
    Next p
    '-----
    Print #1,
    Print #1, "EOF "
    '-----
    Close #1
    
    Create1.MousePointer = 0
    Exit Sub
    '------

errhandle4:
    Close
    Create1.MousePointer = 0
    
    MsgBox Error(Err), 0, "F_XY"
    
    Exit Sub

End Sub  '------------------------------------------------


