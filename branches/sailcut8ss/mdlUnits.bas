Attribute VB_Name = "mdlUnits"
Option Explicit
' 10 November 2004
' Copyright (C) 2004 Steve Studley************************************************
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
'******************************************************************************

Public Function UnitConvert(strUnit As String, UnitType%, Optional Precision% = 0, Optional UnitsShow As Boolean = True) As String
'**********************************************************************
' data is stored as mm, conversion is only for interface display
'***********************************************************************
Dim dblUnit As Double
Dim Foot$, Inch$
Dim intLen As Integer
Dim Convert$
Dim fmt$
Dim i%
Dim unit$

fmt = "######0"



If strUnit = "" Then
    dblUnit = 0
Else
   dblUnit = CDbl(strUnit)
End If

Select Case UnitType
    Case 0  'metric - normal mm
    
        Convert = CStr(dblUnit) ', "#####0 mm") std Precision = 0, std Unit = " mm"
        unit = " mm"
            
        If Precision <> 0 Then
           fmt = fmt & "."
           For i = 1 To Precision
               fmt = fmt & "0"
           Next i
        End If
        
        If UnitsShow Then fmt = fmt & unit
        UnitConvert = Format(Convert, fmt)

    Case 1  'imperial - convert mm to inches
    
        Convert = CStr(dblUnit / 25.4) ', "#####0.00") & Chr(34) std Precision = 2   std Unit = Chr(34)
        unit = Chr(34)
        
        If Precision = 0 Then Precision = 2

            If Precision <> 0 Then
               fmt = fmt & "."
               For i = 1 To Precision
                   fmt = fmt & "0"
               Next i
            End If

            
            If UnitsShow Then
                UnitConvert = Format(Convert, fmt) & unit
            Else
                UnitConvert = Format(Convert, fmt)
            End If
            
        
        
    Case 2  'imperial - convert mm to decimal feet
    
        Convert = CStr(dblUnit / 304.8) ', "#####0.000'") std Precision = 3   std Unit = Chr(39)
        unit = Chr(39)
        
        If Precision = 0 Then Precision = 3

            If Precision <> 0 Then
               fmt = fmt & "."
               For i = 1 To Precision
                   fmt = fmt & "0"
               Next i
            End If

            
            If UnitsShow Then
                UnitConvert = Format(Convert, fmt) & unit
            Else
                UnitConvert = Format(Convert, fmt)
            End If
        

    Case 3  'imperial - convert mm to feet and inches
    
        'get feet part without any decimal part
        Foot$ = Format(dblUnit / 304.8, "#####0.00000")
        intLen = Len(Foot$) - 6
        
        'convert decimal part of feet to inches and decimal inches
        Inch$ = CStr(CDbl(Right(Foot$, 6) * 12)) ', "#0.###") std Precision = 3   std Unit = Chr(39) & " - " & Chr(34)
    
        Foot$ = Left(Foot$, intLen)

        If Precision = 0 Then Precision = 3
        
            If Precision <> 0 Then
               fmt = fmt & "."
               For i = 1 To Precision
                   fmt = fmt & "0"
               Next i
            End If

            
        If UnitsShow Then
            UnitConvert = Foot$ & Chr(39) & " - " & Format(Inch$, fmt) & Chr(34)
        Else
            UnitConvert = Foot$ & " - " & Format(Inch$, fmt)
        End If
   
End Select
    


End Function


Public Function UnitConvertA(strUnit As String, UnitType As Integer, Optional Precision% = 0, Optional UnitsShow As Boolean = True) As String
'**********************************************************************
'For Surface Area
' data is stored as m^2, conversion is only for interface display
'***********************************************************************
Dim dblUnit As Double


If strUnit = "" Then
    dblUnit = 0
Else
   dblUnit = CDbl(strUnit)
End If

Select Case UnitType
    Case 0  'metric - normal meter^2
        UnitConvertA = Format(dblUnit, "#####0.00")
'    Case 1  'imperial - convert m^2 to inch^2
'        UnitConvertA = Format(dblUnit * 645.16, "########0")
    Case 1 To 3    'imperial - convert m^2 to  feet^2
        UnitConvertA = Format(dblUnit / 0.0929, "#####0.0")

   
End Select

End Function

