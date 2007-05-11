Attribute VB_Name = "GLOBAL"
 
Option Explicit
' 10 May 2007******************************************************
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

'-------------------------- definition of constants
Global Const titre$ = "Sailcut 8"
Global Const copyright$ = "Copyright Robert Lainé  and Steve Studley 1997-2007"
Global Const version$ = " 10 May 2007"
Global Const PI = 3.141592653589
Global Const RAD = 57.29577951308
Global Const timeout As Double = 35800.1
Global timenow As Double   '35731 = 28oct97
Global Langue As Integer  '0=français 1=english 3=norwegian
'-------------------------- definition of variables
Option Base 0
Global Sail As String                    '--- name of sail
Global genre1$, genre2$, genre3$     '--- sail charactéristiques
Global LoLuff!, UpLuff  As Single         '--- hauteur coté guindant
Global LoLeech As Single                 '--- hauteur coté chute
Global LBatten!, LYard As Single
Global Xtack!, Ytack!, Ztack As Single       '--- tack coordonnées point d'amure
Global XClew!, YClew!, ZClew As Single       '--- clew coordonnées point d'écoute
Global Xhead!, Yhead!, Zhead As Single       '--- head coordonnées tête
Global Xpeak!, Ypeak!, Zpeak As Single       '--- peak of yard coordonnées
Global Xmluff!, Ymluff!, Zmluff As Single    '--- mid luff coordonnées milieu luff
Global Xtluff!, Ytluff!, Ztluff As Single    '--- luff top quarter coordonnées top1/4 luff
Global Xbluff!, Ybluff!, Zbluff As Single    '--- luff bottom quarter coordonnées bottom1/4 luff
Global Xmleech!, Ymleech!, Zmleech As Single '--- mid leech coordonnées milieu leech
Global Xtleech!, Ytleech!, Ztleech As Single '--- leech top quarter coordonnées top1/4 leech
Global Xbleech!, Ybleech!, Zbleech As Single '--- leech bottom quarter coordonnées bottom1/4 leech

Global AHead As Single         ' in RADIAN angle tête section basse
Global AFoot%, AYard As Integer  ' in DEGREE
Global Mdepth(0) As Single       ' Mid depth array for individual panel chords
'Global Mdepth As Single        ' Mid depth
Global RPdepth As Single       ' depth position
Global Atwist As Single        ' sail twist
Global Surface As Single       ' sail area

Global ClothW As Integer        '-- largeur utile du tissu
Global SeamW As Integer         '-- largeur recouvrement couture radiales
Global SeamT As Integer         '-- largeur recouvrement couture horizontales
Global dCloth(40) As Single   '-- excés de largeur d'un panneau

Global nHpanel As Integer       '-- number of head panels nombre de panneaux tête
Global nBpanel As Integer       '-- delta nbre de panneaux bas à couture2

Global xo!, yo As Single       ' origin (tack) coordonnées origine(amure)
Global ks As Single            ' screen scale factor - facteur échelle écran
Global kp As Single            ' printer scale factor - facteur echelle imprimante

Global px(0 To 64, 22) As Single  ' maximum 64 panels of 22 points each
Global py(0 To 64, 22) As Single
Global pz(0 To 64, 22) As Single

Global pcx(0 To 64, 22) As Single ' developped centre line
Global pcy(0 To 64, 22) As Single
Global plx(0 To 64, 22) As Single ' developped lower edge
Global ply(0 To 64, 22) As Single
Global pmx(0 To 64, 22) As Single ' developped upper edge
Global pmy(0 To 64, 22) As Single

Global optImpt(5) As Integer
Global optSave(5) As Integer
Global infoRead As Boolean

'----- scroll bars position
Global LoLuffScrl%, UpLuffScrl%, LBattenScrl%, FootAScrl%
Global LyardScrl%, YardAScrl%, LoLeechScrl%
Global twistScrl%
Global Lang(60, 6) As String
Global SCVars(30) As String
Global UnitType%
Global intLangue%
Global UnitCaption(25, 3) As String
Global CtlNames(13) As String
Global infoDirty As Boolean
Global WinTop!
Global WinLeft!
Global WinHeight!
Global WinWidth!
Global MRUList$()
Global MaxMRU%
Global AngleMaxOK%
Global AngleDir As Boolean
Global YardReset As Boolean
Global UpLeechGood As Boolean
Global SailStyle%
