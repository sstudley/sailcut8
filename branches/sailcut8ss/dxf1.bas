Attribute VB_Name = "dxf1"
Option Explicit   ' 30 mars 2004

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

