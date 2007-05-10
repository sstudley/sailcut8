Attribute VB_Name = "mdlMRU"
Option Explicit
' 11 November 2004******************************************************
' Copyright (C) 1997-2004 Steve Studley
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
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-
'********************************************************************************


Public Sub AddFileMRU(filename$)
Dim MRUindex$, MRUvalue$, MRUtime$


    If Dir$(filename) <> "" Then
    
        MRUtime = Now
 'to do- need to check for dupes
 '  need number of files in MRU
 '  then write mru files according to date
 '  files need to be written with date as part of the key
 '  write each file to its own key: mrUindex and update as you go?
 '  no more files than MaxMRU
 'THEN add files to Files Menu with seperator if mru > 0
 
        IniKeysWrite "MRU", MRUindex, MRUvalue
        
    Else
        Exit Sub
    End If
    
End Sub


'
' List all of the keys in a particular section
'
Public Sub ListMRU()
    Dim characters As Long
    Dim KeyList$, i%
    Dim filename$

    filename$ = AppIniPathGet
    
    KeyList$ = String$(128, 0)

    characters = GetPrivateProfileStringKeys("MRU", 0, "", KeyList$, 127, filename$)
    
    ' Load sections into array
    Dim NullOffset%
    Do
        NullOffset% = InStr(KeyList$, Chr$(0))
        If NullOffset% > 1 Then
                i = 1
            MRUList$(i) = Mid$(KeyList$, 1, NullOffset% - 1)
            KeyList$ = Mid$(KeyList$, NullOffset% + 1)
            i = i + 1
        End If
    Loop While NullOffset% > 1
    
End Sub
