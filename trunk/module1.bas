Attribute VB_Name = "MODULE1"
Option Explicit    '29 décembre 2002

Sub cadrer(x!(), y!(), pt%)
    ' cadrage d'un tableau de points
' x() tableau des x
' y() tableau des y
' pt  nombre de points 1 -> pt
' pt  0 = offset
Dim i%
'-----
    x(0) = 999999
    y(0) = 999999
    
    For i = 1 To pt
        If x(i) < x(0) Then
            x(0) = x(i)
        End If
        
        If y(i) < y(0) Then
            y(0) = y(i)
        End If
    Next i
    
    For i = 1 To pt
        x(i) = x(i) - x(0)
        y(i) = y(i) - y(0)
    Next i

End Sub '--------------------------------------------- cadrer --

Sub dePerp(x0, y0, x1, y1, d)
    ' calculate displacement d perpendicular to (x0,y0)-(x1,y1)
    ' x0, y0 origin of vector
    ' x1, y1 end of vector then return new point coordinates
    ' with displacement d at +PI/2 trigonometric
    Dim r#, xi#, yi#
    '------
    r = Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2)
    If r <> 0 Then
        xi = x1 - d / r * (y1 - y0)
        yi = y1 + d / r * (x1 - x0)
    Else
        xi = x0
        yi = y0 + d
    End If
    x1 = xi '-- return result of displacement of x1
    y1 = yi '-- return result of displacement of y1

End Sub ' dePerp ----------------------------------------

Sub F_DXF(file$)
    Dim p%, i%
    Dim fichier$
    Dim genre$
    Const fmt2$ = "0.00 "
    '-----------
    'change pointer to hourglass
        Create1.MousePointer = 11

    On Error GoTo errF_DXF


    For p = 1 To nBpanel + nHpanel
        fichier = Left$(file, Len(file) - 6) + Format$(p, "00") + ".DXF"

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
            DXFVertex 1, 7, plx(p, i), ply(p, i), 0
        Next i
        
            DXFVertex 1, 7, pcx(p, 20), pcy(p, 20), 0
        
        For i = 20 To 0 Step -1
            DXFVertex 1, 7, pmx(p, i), pmy(p, i), 0
        Next i
        
            DXFVertex 1, 7, pcx(p, 0), pcy(p, 0), 0
            DXFVertex 1, 7, plx(p, 0), ply(p, 0), 0

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
    ' fichier data
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
        
        Print #1, Str$(Mdepth)
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
    ' fichier data
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
            Mdepth = Val(b)
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
    Dim p%, i%, np%
    Dim r!, g!, b!, n!
    Dim x!, y!, z!, t!
    Const fmt1$ = "0.0 "
    Const fmt2$ = "0.#0 "
    Const fmt3$ = "0.##0 "

End Sub ' F_VRML1 ----------------------------

Sub F_XY(fichier$)
    ' fichier XY développement panneaux
    Dim p%, i%

    Dim genre$
    Const fmt2$ = "0.00 "
    '-------------------
    'change pointer to hourglass
        Create1.MousePointer = 11

    On Error GoTo errhandle4

    Open fichier$ For Output As #1

    Print #1, titre, " Panels development of" + Chr$(9) + Sail

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
            Print #1, Format$(plx(p, i), fmt2) + Chr$(9) + Format$(ply(p, i), fmt2)
        Next i
        
        Print #1,
        Print #1, Format$(p, " #0 ") + Chr$(9) + genre + Chr$(9) + "mid luff"
        Print #1, "    X " + Chr$(9) + "    Y"
        Print #1, Format$(pcx(p, 0), fmt2) + Chr$(9) + Format$(pcy(p, 0), fmt2)
        
        Print #1,
        Print #1, Format$(p, " #0 ") + Chr$(9) + genre + Chr$(9) + "upper edge"
        Print #1, "    X " + Chr$(9) + "    Y"

        For i = 0 To 20
            Print #1, Format$(pmx(p, i), fmt2) + Chr$(9) + Format$(pmy(p, i), fmt2)
        Next i
        
        Print #1,
        Print #1, Format$(p, " #0 ") + Chr$(9) + genre + Chr$(9) + "mid leech"
        Print #1, "    X " + Chr$(9) + "    Y"
        Print #1, Format$(pcx(p, 20), fmt2) + Chr$(9) + Format$(pcy(p, 20), fmt2)
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

Function kDepth(kd, y)
    ' depth function of height of profile
    
    kDepth = 1.2 * (1 - 0.2 * y - (1 - y) ^ (kd + 1))

End Function  ' kDepth ---------------------------------

Function kProfile(d, l)
    ' compute coef to achieve profile depth
    ' d= Shape Factor
    ' l= Leech Factor
    Dim b#, c#, dz#, z#, z1#, x#
    '-- d2z=-k*((1-x)^d-l*x)
    b = 1 / ((d + 2) * (d + 1))
    c = l / 6 - b
    
    x = 0.2
    z = 0
    z1 = 0.001
    
    ' search for maximum of depth
    While z1 >= z And x < 0.96
        x = x + 0.02
        z = z1
        z1 = -(1 - x) ^ (d + 2) / ((d + 2) * (d + 1)) - l / 6 * x ^ 3 + c * x + b
    Wend
    
    kProfile = 1 / z

End Function ' kProfile ----------------------------------

