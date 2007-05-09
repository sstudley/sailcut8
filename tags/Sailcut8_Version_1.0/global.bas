Attribute VB_Name = "GLOBAL"
Option Explicit
'-------------------------- définition des constantes
Global Const titre$ = "Sailcut 8"
Global Const copyright$ = "Copyright Robert Lainé 1997-2004"
Global Const version$ = " 30 March 2004"
Global Const PI = 3.141592653589
Global Const RAD = 57.29577951308
Global Const timeout# = 35800.1
Global timenow#   '35731 = 28oct97
Global langue%  '0=français 1=anglais
'-------------------------- définition des variables
Option Base 0
Global Sail$                    '--- nom de la voile
Global genre1$, genre2$, genre3$ '--- charactéristiques
Global LoLuff!, UpLuff!         '--- hauteur coté guindant
Global LoLeech!                 '--- hauteur coté chute
Global LBatten!, LYard!
Global Xtack!, Ytack!, Ztack!   '--- coordonnées point d'amure
Global XClew!, YClew!, ZClew!   '--- coordonnées point d'écoute
Global Xhead!, Yhead!, Zhead!   '--- coordonnées tête
Global Xpeak!, Ypeak!, Zpeak!   '--- coordonnées yard peak
Global Xmluff!, Ymluff!, Zmluff!   '--- coordonnées milieu luff
Global Xtluff!, Ytluff!, Ztluff!   '--- coordonnées top1/4 luff
Global Xbluff!, Ybluff!, Zbluff!   '--- coordonnées bottom1/4 luff
Global Xmleech!, Ymleech!, Zmleech! '--- coordonnées milieu leech
Global Xtleech!, Ytleech!, Ztleech! '--- coordonnées top1/4 leech
Global Xbleech!, Ybleech!, Zbleech! '--- coordonnées bottom1/4 leech

Global AHead!   ' RADIAN angle tête section basse
Global AFoot%, AYard% 'deg
Global Mdepth!                  ' Mid depth
Global RPdepth!                 ' position creux
Global Atwist!                  ' angle de vrillage voile
Global Surface!                 ' sail area

Global ClothW%          '-- largeur utile du tissu
Global SeamW%           '-- largeur recouvrement couture radiales
Global SeamT%           '-- largeur recouvrement couture horizontales
Global dCloth!(40)      '-- excés de largeur d'un panneau

Global nHpanel%          '-- nombre de panneaux tête
Global nBpanel%          '-- delta nbre de panneaux bas à couture2

Global xo!, yo!         '-- coordonnées origine(amure)
Global ks!              '-- facteur échelle écran
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

