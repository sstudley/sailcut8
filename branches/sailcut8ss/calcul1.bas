Attribute VB_Name = "calcul1"

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

Sub deplacePerp(x0, y0, x1, y1, d)
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
    '-----
    x1 = xi
    y1 = yi

End Sub  '---------------------------------------- dePerp --

Function directionXY(x0, y0, x1, y1)
    ' return direction of vector (x0,y0)-(x1,y1) in radians
    ' x0, y0 = origin of vector
    ' x1, y1 = eextrmity of vector
    '------
    If (x1 - x0) = 0 Then
        If y1 > y0 Then
            directionXY = 1.570796326795
        ElseIf y1 = y0 Then
            directionXY = 0
        Else
            directionXY = 4.712388980385
        End If
    
    ElseIf x1 > x0 Then
        If y1 >= y0 Then
            directionXY = Atn((y1 - y0) / (x1 - x0))
        Else
            directionXY = 6.28318530718 + Atn((y1 - y0) / (x1 - x0))
        End If
    
    Else
        directionXY = 3.14159265359 + Atn((y1 - y0) / (x1 - x0))
        
    End If

End Function  '--------------------------- direction ----

Function distance2D(x0!, y0!, x1!, y1!)
    ' return distance between 2 2D points
    distance2D = Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2)

End Function   '------------------------------------------

Function distance3D(x0!, y0!, z0!, x1!, y1!, z1!)
    ' return distance between 2 3D points

    distance3D = Sqr((x1 - x0) ^ 2 + (y1 - y0) ^ 2 + (z1 - z0) ^ 2)

End Function  '--------------------------------------------


Sub intermed(x10!, y10!, x20!, y20!, x30!, y30!, x12!, y12!, x23!, y23!)
    'point 1 = x10,y10
    'point 2 = x20,y20
    'point 3 = x30,y30
    'return intermediates points
    'point 12 is between 1 et 2
    'point 23 is between 2 et 3
    Dim dir13!, dir31!, dir21!, dir23!, dir0!
    Dim l13!, l21!, l23!
    Dim xi!, yi!
    '-----
    dir13 = directionXY(x10, y10, x30, y30)
    dir31 = directionXY(x30, y30, x10, y10)
    dir21 = directionXY(x20, y20, x10, y10)
    dir23 = directionXY(x20, y20, x30, y30)

    l13 = Sqr(Abs(x30 - x10) ^ 2 + Abs(y30 - y10) ^ 2)
    l21 = Sqr(Abs(x10 - x20) ^ 2 + Abs(y10 - y20) ^ 2)
    l23 = Sqr(Abs(x30 - x20) ^ 2 + Abs(y30 - y20) ^ 2)

    'point inter 2-3
    xi = x30 + l23 * Cos(dir13)
    yi = y30 + l23 * Sin(dir13)
    dir0 = directionXY(x20, y20, xi, yi)
    x23 = x20 + l23 / 2 * Cos(dir0)
    y23 = y20 + l23 / 2 * Sin(dir0)

    'point inter 2-1
    xi = x10 + l21 * Cos(dir31)
    yi = y10 + l21 * Sin(dir31)
    dir0 = directionXY(x20, y20, xi, yi)
    x12 = x20 + l21 / 2 * Cos(dir0)
    y12 = y20 + l21 / 2 * Sin(dir0)

End Sub '--------------------------------------------------

Sub Rot2D(xc, yc, X, Y, alfa)
' xc,yc coordinates of center of rotation
' alfa angle of rotation
' X,Y coordinates of point before and then after rotation
Dim r#, a#
'------
   r = Sqr((X - xc) ^ 2 + (Y - yc) ^ 2)
   If (X - xc) = 0 Then
        a = 1.570796326795 * Sgn(Y - yc)
    Else
        a = Atn((Y - yc) / (X - xc))
   End If

   If (X - xc) < 0 Then a = a + 3.14159265359
   
   X = xc + r * Cos(a + alfa)
   Y = yc + r * Sin(a + alfa)

End Sub '--- Rot2D ---------------------------------------

Sub transBorc(x1!, y1!, x2!, y2!, x3!, y3!, d!, x4!, y4!)
    'translation of corner x2,y2 into x4,y4
    'by displacement d projected in direction point 3->2
    Dim a12! 'direction 1-2
    Dim a23! 'direction 2-3
    Dim e!
    '---------
    a12 = directionXY(x1, y1, x2, y2)
    a23 = directionXY(x2, y2, x3, y3)
    If (a23 - a12) <> 0 Then
        e = d / Sin(a23 - a12)
        x4 = x2 - e * Cos(a23)
        y4 = y2 - e * Sin(a23)
    
    Else
        x4 = x2 + d * Cos(a12)
        y4 = y2 + d * Sin(a12)
    
    End If

End Sub  '--------------------- transBorc ----

Sub transBord(x1!, y1!, x2!, y2!, x3!, y3!, d!, x4!, y4!)
    'translation of  corner x2,y2 into x4,y4
    'by displacement d projected in direction 1->2
    Dim a12! 'direction 1-2
    Dim a23! 'direction 2-3
    Dim e!
    '----------
    a12 = directionXY(x1, y1, x2, y2)
    a23 = directionXY(x2, y2, x3, y3)
    If (a23 - a12) <> 0 Then
        e = d / Sin(a23 - a12)
        x4 = x2 + e * Cos(a12)
        y4 = y2 + e * Sin(a12)
    
    Else
        x4 = x2 + d * Cos(a12)
        y4 = y2 + d * Sin(a12)
    
    End If

End Sub '-------- transBord -------------------------------

Sub transCoin(x1!, y1!, x2!, y2!, x3!, y3!, d!, x4!, y4!)
    'translation of corner x2,y2 into x4,y4
    'by displacement d in direction of external bissector
    'of angle 1-2-3
    Dim a12!, a23! 'directions 1->2 and  2->3
    Dim e!
    '-----
    a12 = directionXY(x1, y1, x2, y2)
    a23 = directionXY(x2, y2, x3, y3)
    e = d * Tan((a23 - a12) / 2)
    x4 = x2 + d * Cos(a12 - 1.570796326795) + e * Cos(a12)
    y4 = y2 + d * Sin(a12 - 1.570796326795) + e * Sin(a12)
    
End Sub  '------------------- transCoin ------------------

