Attribute VB_Name = "dxf1"

Option Explicit   
' 17 October 2004******************************************************
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
Sub DXFEnd()

    Print #1, "  0"
    Print #1, "EOF"
    Close #1

End Sub ' DXFEnd ------------------------------------------

Sub DXFHeader(txt$)
    'to be closed by DXFEnd
    Print #1, "999"
    Print #1, txt

End Sub ' DXFHeader ---------------------------------------

Sub DXFPolyline(layer%, couleur%)
        'to be followed by serie of DXFVertex and closed by DXFSequenceEnd
        Print #1, "  0"
        Print #1, "POLYLINE"
        Print #1, "  8"
        Print #1, Format(layer, "0")
        Print #1, " 62"
        Print #1, Format(couleur, "     0")
        Print #1, " 66"
        Print #1, "      1"
        Print #1, "  6"
        Print #1, "CONTINUOUS"
        Print #1, " 70"
        Print #1, "      8"

End Sub  ' DXFPolyline ------------------------------------

Sub DXFSectionEnd()

    Print #1, "  0"
    Print #1, "ENDSEC"

End Sub ' DXFSectionEnd -----------------------------------

Sub DXFSectionHeaderGeometry()
    'to be closed by DXFSectionEnd
    Print #1, "  0"
    Print #1, "SECTION"
    Print #1, "  2"
    Print #1, "ENTITIES"

End Sub ' DXFGeometryHeader -------------------------------

Sub DXFSequenceEnd()

    Print #1, 0
    Print #1, "SEQEND"

End Sub ' DXFSequenceEnd ----------------------------------

Sub DXFVertex(layer%, couleur%, X, Y, Z)
            
        Const fmt$ = "0.0000 "

        Print #1, "  0"
        Print #1, "VERTEX"
        Print #1, "  8"
        Print #1, Format(layer, "0")
        Print #1, " 62"
        Print #1, Format(couleur, "     0")
        Print #1, "  6"
        Print #1, "CONTINUOUS"
        Print #1, " 70"
        Print #1, "     32"
        Print #1, " 10"
        Print #1, Format$(X, fmt)
        Print #1, " 20"
        Print #1, Format$(Y, fmt)
        Print #1, " 30"
        Print #1, Format$(Z, fmt)

End Sub  ' DXFVertex --------------------------------------

