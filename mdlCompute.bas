Attribute VB_Name = "mdlCompute"
Option Explicit
' 20 December 2004******************************************************
' Copyright (C) 1997-2004 Robert Lainé and Steve Studley
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

Public Sub compute()
    Dim n%, i%
    Dim a!, b!, c As Single
    Dim alfa!, beta!, gama!, delta As Single
    Dim dir1!, dir2 As Single
    Dim xi As Single
    Dim h!, dl!, l As Single
    
    ' reset flags
    AngleDir = True
    AngleMaxOK = 0
    YardReset = False
    UpLeechGood = True
    
    ' set all point arrays to 0
    Erase px, py, pz
    
    '-----
    n = 0
    px(0, 0) = 0
    py(0, 0) = 0
    pz(0, 0) = 0

    '----- baton foot
    For i = 1 To 20
        px(0, i) = px(0, 0) + LBatten * i / 20 * Cos(AFoot / RAD)
        py(0, i) = py(0, 0) + LBatten * i / 20 * Sin(AFoot / RAD)
        pz(0, i) = 0
    Next i
    
    '----- batten lower panels

    For n = 1 To nBpanel
        alfa = Atn((py(2 * (n - 1), 20) - py(2 * (n - 1), 0)) / (px(2 * (n - 1), 20) - px(2 * (n - 1), 0)))
        
        px(2 * n, 0) = 0
        py(2 * n, 0) = py(2 * (n - 1), 0) + LoLuff
        pz(2 * n, 0) = 0
        
        beta = Atn((LoLeech - LoLuff * Cos(alfa)) / LBatten)
        '----- check angle

        If (alfa + beta) >= 1.3 Then 'angle max
            beta = 1.3 - alfa
            AngleMaxOK = 1
        End If

        If (alfa + beta) >= AYard / RAD Then
            beta = AYard / RAD - alfa
            AngleMaxOK = 2
        End If
        '-----
        For i = 1 To 20
            px(2 * n, i) = px(2 * n, 0) + LBatten * i / 20 * Cos(alfa + beta)
            py(2 * n, i) = py(2 * n, 0) + LBatten * i / 20 * Sin(alfa + beta)
            pz(2 * n, i) = 0
        Next i

        AHead = alfa + beta
        '----- check pour discontinuity in leech of lower panels
        If n > 2 Then
          dir1 = directionXY(px(2 * (n - 1), 20), py(2 * (n - 1), 20), px(2 * n, 20), py(2 * n, 20))
          dir2 = directionXY(px(2 * (n - 2), 20), py(2 * (n - 2), 20), px(2 * (n - 1), 20), py(2 * (n - 1), 20))
          
          If dir1 <= dir2 Then
            AngleDir = False
          End If
          
        End If
        '-----
    Next n

    '----- intermediates lines in lower panels with profile
    'Call CubicP(RPdepth, a, b, c)  ' profile
    For n = 1 To nBpanel
        For i = 0 To 20
          px(2 * n - 1, i) = (px(2 * n - 2, i) + px(2 * n, i)) / 2
          py(2 * n - 1, i) = (py(2 * n - 2, i) + py(2 * n, i)) / 2
          pz(2 * n - 1, i) = (pz(2 * n - 2, i) + pz(2 * n, i)) / 2
          xi = (i / 20)
          If n < nBpanel Then
            pz(2 * n - 1, i) = pz(2 * n - 1, i) + Mdepth(0) * LBatten * profileP(RPdepth, xi)
            Else ' half depth
            pz(2 * n - 1, i) = pz(2 * n - 1, i) + Mdepth(0) / 2 * LBatten * profileP(RPdepth, xi)
          End If
        Next i
    Next n
'************************************************************************************************************
    '----- batten of head panels
    Select Case SailStyle
        Case 0
            ComputeUpper n
        Case 1
            ComputeUpperVanLoan
    End Select
    

    Xpeak = px(2 * (nBpanel + nHpanel), 20)
    Ypeak = py(2 * (nBpanel + nHpanel), 20)
    Zpeak = pz(2 * (nBpanel + nHpanel), 20)

    '----- intermediate lines in head panels without profile
    For n = (nBpanel + 1) To (nBpanel + nHpanel)
        For i = 0 To 20
            px(2 * n - 1, i) = (px(2 * n - 2, i) + px(2 * n, i)) / 2
            py(2 * n - 1, i) = (py(2 * n - 2, i) + py(2 * n, i)) / 2
            pz(2 * n - 1, i) = (pz(2 * n - 2, i) + pz(2 * n, i)) / 2
        Next i
    Next n

    '----- area
    ComputeArea alfa, beta, gama
    
    '----- adding twist
    For n = 0 To 2 * (nBpanel + nHpanel)
        h = (py(n, 20) - py(0, 20)) / (Ypeak - py(0, 20) + 0.001)
        For i = 0 To 20
            Rot2D 0, 0, px(n, i), pz(n, i), h * Atwist
        Next i
    Next n
    
    '-----
End Sub ' compute ----------------------------------------


Public Sub ComputeUpper(n%)
   Dim i%
    Dim a!, b!, c As Single
    Dim alfa!, beta!, gama!, delta As Single
    Dim dir1!, dir2 As Single
    Dim xi As Single
    Dim h!, dl!, l As Single
    
    alfa = Atn((py(2 * (n - 1), 20) - py(2 * (n - 1), 0)) / (px(2 * (n - 1), 20) - px(2 * (n - 1), 0)))
    delta = (AYard / RAD - alfa) / nHpanel

    If beta < 0 Then
        beta = 0
        AYard = alfa * RAD
        YardReset = True
    End If

    dl = (LYard - LBatten)


    For n = (nBpanel + 1) To (nBpanel + nHpanel)

        alfa = alfa + delta
        l = LBatten + dl * ((n - nBpanel) / nHpanel) ^ 1.5

            px(2 * n, 0) = 0
            py(2 * n, 0) = py(2 * (n - 1), 0) + UpLuff
            pz(2 * n, 0) = 0


        For i = 1 To 20
            px(2 * n, i) = px(2 * n, 0) + l * i / 20 * Cos(alfa)
            py(2 * n, i) = py(2 * n, 0) + l * i / 20 * Sin(alfa)
            pz(2 * n, i) = 0
        Next i

        '----- check for discontinuity in leech of head panels
        If nBpanel > 1 Then
            dir1 = directionXY(px(2 * (n - 1), 20), py(2 * (n - 1), 20), px(2 * n, 20), py(2 * n, 20))
            dir2 = directionXY(px(2 * (n - 2), 20), py(2 * (n - 2), 20), px(2 * (n - 1), 20), py(2 * (n - 1), 20))

            If dir1 <= dir2 Then
                UpLeechGood = False
            End If

        End If
        '-----
    Next n


End Sub

Public Sub ComputeUpperVanLoan()
    Dim i%, n%
    Dim a0!, a1!, b0!, b1!, c0!, c1!, c2!
    Dim alfa!, beta!, gama!, delta!, theta!
    Dim x0!, x1!, x2!, x3!, y0!, y1!, y2!, y3!
   

    n = nBpanel + nHpanel          ' Yard first
    
       px(2 * n, 0) = 0
       py(2 * n, 0) = (nBpanel + 1) * LoLuff
       pz(2 * n, 0) = 0
       
       For i = 1 To 20
           px(2 * n, i) = px(2 * n, 0) + LBatten * i / 20 * Cos(AYard / RAD)
           py(2 * n, i) = py(2 * n, 0) + LBatten * i / 20 * Sin(AYard / RAD)
           pz(2 * n, i) = 0
       Next i
            
    n = nBpanel + 1         ' mid upper batten

    x1 = px(0, 20)          ' Boom endpoint x coor
    x0 = px((nBpanel + nHpanel) * 2, 20)                ' Yard endpoint x coor
    y1 = py(0, 20)          ' Boom endpoint y coor
    y0 = py((nBpanel + nHpanel) * 2, 20)               ' Yard endpoint y coor
    
    c0 = Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2)         ' length VL line from clew to yard peak
    theta = Atn((x1 - x0) / c0)                          ' angle of VL line from leech (vertical)
    beta = (PI / 2) - (AFoot / RAD)                 ' angle of upper lower batten to leech (vertical)
    gama = PI - (theta + beta)                         '  angle of VL line to upper lower batten
    c1 = (nBpanel * LoLuff) / Sin(gama) * Sin(beta)         ' length of VL line from clew to upper lower batten
    c2 = (c0 - c1) / 2                                     ' length from yard peak to intersection of upper batten
    x2 = (c2 / Sin(PI / 2) * Sin(theta)) + x0        ' intersection x coor
    y2 = y0 - (c2 / Sin(PI / 2) * Sin((PI / 2) - theta))     ' intersection y coor
    
    x3 = 0
    px(2 * n, 0) = x3                                          ' upper batten origin x coor
    y3 = (nBpanel + 0.75) * LoLuff
    py(2 * n, 0) = y3                                           ' upper batten origin y coor
    pz(2 * n, 0) = 0
    
    delta = Atn((y2 - y3) / x2)

    For i = 1 To 20
        px(2 * n, i) = px(2 * n, 0) + LBatten * i / 20 * Cos(delta)
        py(2 * n, i) = py(2 * n, 0) + LBatten * i / 20 * Sin(delta)
        pz(2 * n, i) = 0
    Next i
   

End Sub



Public Sub ComputeArea(alfa!, beta!, gama!)
    Dim a!, b!, c!
    Dim n%, i%
    
    Surface = 0
    For n = 1 To 2 * (nBpanel + nHpanel)
        For i = 1 To 20
            a = distance2D(px(n - 1, i - 1), py(n - 1, i - 1), px(n, i - 1), py(n, i - 1))
            b = distance2D(px(n - 1, i - 1), py(n - 1, i - 1), px(n, i), py(n, i))
            c = distance2D(px(n, i - 1), py(n, i - 1), px(n, i), py(n, i))
            Call Triangle(a, b, c, alfa, beta, gama)
            Surface = Surface + Sin(gama) * (a * b) / 2
            '-----
            a = distance2D(px(n - 1, i), py(n - 1, i), px(n, i), py(n, i))
            c = distance2D(px(n - 1, i - 1), py(n - 1, i - 1), px(n - 1, i), py(n - 1, i))
            Call Triangle(a, b, c, alfa, beta, gama)
            Surface = Surface + Sin(gama) * (a * b) / 2
        Next i
    Next n
    Surface = Surface / 1000000

End Sub
