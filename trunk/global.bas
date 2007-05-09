Attribute VB_Name = "GLOBAL"
Option Explicit
'-------------------------- d�finition des constantes
Global Const titre$ = "Sailcut 8"
Global Const copyright$ = "Copyright Robert Lain� 1997-2004"
Global Const version$ = " 30 March 2004"
Global Const PI = 3.141592653589
Global Const RAD = 57.29577951308
Global Const timeout# = 35800.1
Global timenow#   '35731 = 28oct97
Global langue%  '0=fran�ais 1=anglais
'-------------------------- d�finition des variables
Option Base 0
Global Sail$                    '--- nom de la voile
Global genre1$, genre2$, genre3$ '--- charact�ristiques
Global LoLuff!, UpLuff!         '--- hauteur cot� guindant
Global LoLeech!                 '--- hauteur cot� chute
Global LBatten!, LYard!
Global Xtack!, Ytack!, Ztack!   '--- coordonn�es point d'amure
Global XClew!, YClew!, ZClew!   '--- coordonn�es point d'�coute
Global Xhead!, Yhead!, Zhead!   '--- coordonn�es t�te
Global Xpeak!, Ypeak!, Zpeak!   '--- coordonn�es yard peak
Global Xmluff!, Ymluff!, Zmluff!   '--- coordonn�es milieu luff
Global Xtluff!, Ytluff!, Ztluff!   '--- coordonn�es top1/4 luff
Global Xbluff!, Ybluff!, Zbluff!   '--- coordonn�es bottom1/4 luff
Global Xmleech!, Ymleech!, Zmleech! '--- coordonn�es milieu leech
Global Xtleech!, Ytleech!, Ztleech! '--- coordonn�es top1/4 leech
Global Xbleech!, Ybleech!, Zbleech! '--- coordonn�es bottom1/4 leech

Global AHead!   ' RADIAN angle t�te section basse
Global AFoot%, AYard% 'deg
Global Mdepth!                  ' Mid depth
Global RPdepth!                 ' position creux
Global Atwist!                  ' angle de vrillage voile
Global Surface!                 ' sail area

Global ClothW%          '-- largeur utile du tissu
Global SeamW%           '-- largeur recouvrement couture radiales
Global SeamT%           '-- largeur recouvrement couture horizontales
Global dCloth!(40)      '-- exc�s de largeur d'un panneau

Global nHpanel%          '-- nombre de panneaux t�te
Global nBpanel%          '-- delta nbre de panneaux bas � couture2

Global xo!, yo!         '-- coordonn�es origine(amure)
Global ks!              '-- facteur �chelle �cran
Global kp!              '-- facteur echelle imprimante

Global px!(0 To 64, 22) ' 64 panneaux maxi 22 pts each
Global py!(0 To 64, 22)
Global pz!(0 To 64, 22)

Global pcx!(0 To 64, 22) ' developped centre line
Global pcy!(0 To 64, 22)
Global plx!(0 To 64, 22) ' developped lower edge
Global ply!(0 To 64, 22)
Global pmx!(0 To 64, 22) ' developped upper edge
Global pmy!(0 To 64, 22)

Global optImpt%(5)
Global optSave%(5)

Global LoLuffScrl%, UpLuffScrl%, LBattenScrl%, FootAScrl%
Global LyardScrl%, YardAScrl%, LoLeechScrl%
Global twistScrl%

