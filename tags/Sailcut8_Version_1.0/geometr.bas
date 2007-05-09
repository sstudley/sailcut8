Attribute VB_Name = "GEOMETR"
Option Explicit  ' 30 aout 2000

Sub CubicP(pos, a, b, c)
    ' équation de la forme  y= a + bx^2 + cx^3
    ' input  p = position du maximum (centré si p=.5)
    ' renvoi les coefficients a,b,c.
    ' x=-p  => y=0
    ' x=0   => y=1
    ' x=1-p => y=0
    '--- vérification des limites de la position du maximum
    If pos < 0.3 Then
        pos = 0.3
    ElseIf pos > 0.7 Then
        pos = 0.7
    End If
    '--- calcul des coefficients  a, b, c
    a = 1
    b = (3 * pos - 1 - 3 * pos ^ 2) / (pos ^ 2 * (1 - 2 * pos + pos ^ 2))
    c = (1 - 2 * pos) / (pos ^ 4 - 2 * pos ^ 3 + pos ^ 2)

End Sub ' CubicP -----------------------------------------

Function Direction2D(x0#, y0#, x1#, y1#) As Double
' renvoi la direction en radians du vecteur (x0,y0)=>(x1,y1)
' x0, y0 = origine
' x1, y1 = extrémité
    '------
    If (x1 - x0) = 0 Then
        If y1 > y0 Then
            Direction2D = 1.570796326795
        ElseIf y1 = y0 Then
            Direction2D = 0
        Else
            Direction2D = 4.712388980385
        End If
    
    ElseIf x1 > x0 Then
        If y1 >= y0 Then
            Direction2D = Atn((y1 - y0) / (x1 - x0))
        Else
            Direction2D = 6.28318530718 + Atn((y1 - y0) / (x1 - x0))
        End If
    
    Else
        Direction2D = 3.14159265359 + Atn((y1 - y0) / (x1 - x0))
        
    End If

End Function  ' Sub Direction2D --------------------------

Sub dPerp(x1#, y1#, x2#, y2#, d#)
    ' calculate displacement d perpendicular to (x1,y1)-(x2,y2)
    ' x1,y1 origin of vector
    ' x2,y2 end of vector to be moved= return new point coordinates
    ' d displacement at +90 deg trigonometric
    Dim r#, xi#, yi#
    '------
    r = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)

    If r <> 0 Then
        xi = x2 - d / r * (y2 - y1)
        yi = y2 + d / r * (x2 - x1)
    Else
        xi = x1
        yi = y1 + d
    End If
    x2 = xi '-- return result of displacement of x2
    y2 = yi '-- return result of displacement of y2

End Sub ' dPerp ------------------------------------------

Function Hauteur2D(x1#, y1#, x2#, y2#, x3#, y3#)
    'hauteur du point (x1,y1) à la droite (x2,y2)-(x3,y3)
    Dim a#, b#, c#
    Dim aa#, bb#, cc#
    
    a = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    b = Sqr((x2 - x3) ^ 2 + (y2 - y3) ^ 2)
    c = Sqr((x3 - x1) ^ 2 + (y3 - y1) ^ 2)
    
    Triangle a, b, c, aa, bb, cc
    
    If a >= c Then
        Hauteur2D = a * Cos(cc)
    Else
        Hauteur2D = c * Cos(aa)
    End If
    
End Function 'Hauteur2D -----------------------------------

Function Hauteur3D(x1#, y1#, z1#, x2#, y2#, z2#, x3#, y3#, z3#)
    'hauteur du point (x1,y1,z1) à la droite (x2,y2,z2)-(x3,y3,z3)
    Dim a#, b#, c#
    Dim aa#, bb#, cc#
    
    a = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2 + (z1 - z2) ^ 2)
    b = Sqr((x2 - x3) ^ 2 + (y2 - y3) ^ 2 + (z2 - z3) ^ 2)
    c = Sqr((x3 - x1) ^ 2 + (y3 - y1) ^ 2 + (z3 - z1) ^ 2)
    
    Triangle a, b, c, aa, bb, cc
    
    If a >= c Then
        Hauteur3D = a * Cos(cc)
    Else
        Hauteur3D = c * Cos(aa)
    End If

End Function 'Hauteur3D -----------------------------------

Sub Intersection(x1#, y1#, x2#, y2#, x3#, y3#, x4#, y4#, x#, y#)
    ' Renvoi le point x,y correspondant à l'intersection
    ' de la droite passant par (x1,y1)-(x2,y2)    y=a1*X+b1
    ' avec la droite passant par (x3,y3)-(x4,y4)  y=a3*X+b3

    Dim a1#, b1#, a3#, b3# ' As Double

    If x1 <> x2 Then 'première droite inclinée
        a1 = (y2 - y1) / (x2 - x1)
        b1 = y1 - x1 * a1
    
        If x3 <> x4 Then
            a3 = (y4 - y3) / (x4 - x3)
            b3 = y3 - x3 * a3
            If a1 <> a3 Then 'deuxième droite inclinée
                x = (b3 - b1) / (a1 - a3)
                y = a1 * x + b1
              Else 'parrallèles =>indétermination
                x = (x1 + x2 + x3 + x4) / 4
                y = (y1 + y2 + y3 + y4) / 4
            End If
    
          Else 'deuxième droite verticale
            x = x3
            y = y1 + (x3 - x1) * a1
        End If
        '-----
      Else 'première droite verticale
        If x3 <> x4 Then 'deuxième droite inclinée
            a3 = (y4 - y3) / (x4 - x3)
            b3 = y3 - x3 * a3
            x = x1
            y = y3 + (x1 - x3) * a3
    
          Else 'deuxième droite verticale => indétermination
                x = (x1 + x3) / 2
                y = (y1 + y2 + y3 + y4) / 4
        End If
        '-----
    End If

End Sub ' intersection ------------------------------------

Function profileP(pos, x)
    ' return depth value =f(x) parabolic
    ' for a profile with depth max position =pos
    ' depth is normalised to 1
    Dim y#
    
    If x < 0 Or x > 1 Then
        profileP = 0
    ElseIf x < pos Then
        profileP = 1 - (1 - (x / pos)) ^ 2
    Else
        profileP = 1 - ((x - pos) / (1 - pos)) ^ 2
    End If
    
End Function ' -----------------------

Sub Translation(x0, y0, d, alfa, x1, y1)
    ' x0, y0 coordonnées de l'origine
    ' d module du vecteur de translation
    ' alfa = direction du vecteur de translation en radians
    ' x1, y1 point résultat
    
    x1 = x0 + d * Cos(alfa)
    y1 = y0 + d * Sin(alfa)

End Sub ' translation -------------------------------------

Sub TranslationPerp(x0, y0, x1, y1, d, x2, y2)
    ' déplacement d perpendiculaire à (x0,y0)-(x1,y1)
    ' x0, y0 origine du vecteur
    ' x1, y1 extrémité du vecteur
    ' d déplacement perpendiculaire = +PI/2 sens trigonométrique directe
    ' x2, y2 résultat de la translation de x1, y1
    
    Dim r#, xi#, yi#
    
    r = Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2)
    If r <> 0 Then
        xi = x1 - d / r * (y1 - y0)
        yi = y1 + d / r * (x1 - x0)
    Else
        xi = x0
        yi = y0 + d
    End If
    '-----
    x2 = xi
    y2 = yi

End Sub  ' Sub translationPerp ----------------------------

Function TranslatX(x0#, y0#, d#, alfa#)  'as double
    ' renvoi l'absisse X du point translaté
    ' x0, y0 coordonnées de l'origine
    ' d déplacement
    ' alfa = direction de translation en radians
    
    TranslatX = x0 + d * Cos(alfa)
    'TranslatY = y0 + d * Sin(alfa)

End Function  ' TranslatX ---------------------------------

Function TranslatY(x0#, y0#, d#, alfa#)  'As Double
    ' renvoi l'ordonnée Y du point translaté
    ' x0, y0 coordonnées de l'origine
    ' d déplacement
    ' alfa = direction de translation en radians
    
    'TranslatX = x0 + d * Cos(alfa)
    TranslatY = y0 + d * Sin(alfa)

End Function ' TranslatY ----------------------------------

Sub Triangle(a, b, c, alfa, beta, gama)
    ' calcul les angles d'un triangle à partir des 3 cotés a,b,c.
    Dim per#  'demi périmètre
    
    If (a + b + c) < 0.000001 Then
        alfa = PI / 3
        beta = PI / 3
        gama = PI / 3
    Else
        per = (a + b + c) / 2
        If (per - a) < 0.0000001 * per Then
            alfa = PI
            beta = 0
            gama = 0
        ElseIf (per - b) < 0.0000001 * per Then
            beta = PI
            alfa = 0
            gama = 0
        ElseIf (per - c) < 0.0000001 * per Then
            gama = PI
            alfa = 0
            beta = 0
        Else
            alfa = 2 * Atn(Sqr((per - b) * (per - c) / (per * (per - a))))
            beta = 2 * Atn(Sqr((per - a) * (per - c) / (per * (per - b))))
            gama = 2 * Atn(Sqr((per - a) * (per - b) / (per * (per - c))))
        End If
    End If
    
End Sub 'Triangle -----------------------------------------

