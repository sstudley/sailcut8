Attribute VB_Name = "mdlINI"
Option Explicit

'10 November 2004
' Copyright 2004, Steve Studley************************************************
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


' Added 9 Nov 2004 - SBS
' This first line is the declaration from win32api.txt
' Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function GetPrivateProfileStringKeys& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
' This first line is the declaration from win32api.txt
' Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Long
Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)
Declare Function WritePrivateProfileStringToDeleteKey& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String)
Declare Function WritePrivateProfileStringToDeleteSection& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lplFileName As String)

  
Public Function VBGetPrivateProfileString(section$, key$) As String        ' Added 9 Nov 2004 - SBS

    Dim KeyValue$
    Dim characters As Long
    Dim filename$

    filename$ = AppIniPathGet

    
    KeyValue$ = String$(128, 0)
    
    characters = GetPrivateProfileStringByKeyName(section$, key$, "", KeyValue$, 127, filename$)

    If characters > 1 Then
        KeyValue$ = Left$(KeyValue$, characters)
    End If
    
    VBGetPrivateProfileString = KeyValue$

End Function

Public Function KeyValueGet$(section$, key$)        ' Added 9 Nov 2004 - SBS

 
    
    ' Retrieve the list of keys in the section
    KeyValueGet$ = VBGetPrivateProfileString(section$, key$)
    

End Function
Public Sub IniKeysWrite(section$, key$, value$)     ' Added 9 Nov 2004 - SBS
Dim filename$
Dim success%

filename$ = AppIniPathGet

 success% = WritePrivateProfileStringByKeyName(section$, key$, value$, filename$)
    
    If success% = 0 Then
        Dim msg$
        msg$ = "Write to INI file failed - this is typically caused by a write protected INI file"
        MsgBox msg$
        Exit Sub
    End If
End Sub

Public Sub IniKeysSet()     'Added 10 Nov 2004 -SBS

' Write keys to INI file

    IniKeysWrite "App", "Language", CStr(Langue)
    IniKeysWrite "App", "UnitType", CStr(UnitType)
    IniKeysWrite "Window", "Width", CStr(Create1.Width)
    IniKeysWrite "Window", "Height", CStr(Create1.Height)
    IniKeysWrite "Window", "Top", CStr(Create1.Top)
    IniKeysWrite "Window", "Left", CStr(Create1.Left)
    
End Sub





Public Function AppIniPathGet$()
Dim filename$

    filename$ = App.Path
    If Right$(filename$, 1) <> "\" Then filename$ = filename$ & "\"
    filename$ = filename$ & App.EXEName & ".ini"
    
    AppIniPathGet$ = filename$
    

End Function
