Attribute VB_Name = "mdlLang"

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

'**************************
'french = 0
'english = 1
'norwegian = 2
'dutch = 3
'german = 4
'spanish = 5
'finnish = 6
'*************************


Public Sub mLang_LoadArray()
    'menu1 french
    Lang(0, 0) = "Bienvenue aux maîtres voiliers"
    Lang(1, 0) = "Ce programme pour concevoir des voiles de jonques est mis à disposition pour évaluation et sans aucune garantie sous"
    Lang(2, 0) = "GNU General Public Licence version 2, tel que publié par la Free Software Foundation."
    Lang(3, 0) = ""
    Lang(4, 0) = "Copyright et autres droits réservés"
    Lang(5, 0) = "Sailcut est une marque déposée par Robert Lainé"
    Lang(6, 0) = "Débuter"
    Lang(7, 0) = "Fin"
    Lang(8, 0) = "Ecrivez à: robert.laine@sailcut.com"
        'Create1 french
    Lang(9, 0) = "  Géométrie : modifier la voile - "
    Lang(10, 0) = "  Géométrie d'une nouvelle voile"
    Lang(11, 0) = "Largeur coté guindant panneaux de tête"
    Lang(12, 0) = "Largeur coté guindant panneaux bas"
    Lang(13, 0) = "Largeur coté chute panneaux bas"
    Lang(14, 0) = "Longueur latte"
    Lang(15, 0) = "Longueur Yard"
    Lang(16, 0) = "Angle bordure"
    Lang(17, 0) = "Angle yard"
    Lang(18, 0) = "Surface"
    Lang(19, 0) = "Creux"
    Lang(20, 0) = "Position du creux"
    Lang(21, 0) = "Vrillage"
    Lang(22, 0) = "Panneaux tête"
    Lang(23, 0) = "Panneaux bas "
    Lang(24, 0) = "Nom fichier voile"
    Lang(25, 0) = "Développement"
    Lang(26, 0) = "Quitter"
    Lang(27, 0) = "&Fichier"
    Lang(28, 0) = "&Nouveau"
    Lang(29, 0) = "&Ouvrir fichier"
    Lang(30, 0) = "Panneaux DXF"
    Lang(31, 0) = "Im&Primer"
    Lang(32, 0) = "Données"
    Lang(33, 0) = "Panneaux XY"
    Lang(34, 0) = "&Sauver"
    Lang(35, 0) = "Données"
    Lang(36, 0) = "Voile DXF"
    Lang(37, 0) = "VRML 1.0"
    Lang(38, 0) = "Quitter"
    Lang(39, 0) = "&Tissus"
    Lang(40, 0) = "Ouvrir"
    Lang(41, 0) = "Sauver"
    Lang(42, 0) = "Crédits"
    Lang(43, 0) = "Développeurs"
    Lang(44, 0) = "Traductions"
    Lang(45, 0) = "Français"
    Lang(46, 0) = "Anglais"
    Lang(47, 0) = "Norvégien"
    Lang(48, 0) = "Néerlandais"
    Lang(49, 0) = "Allemand"
    Lang(50, 0) = "Espagnol"
    Lang(51, 0) = "Finlandais"
    Lang(52, 0) = "Paramètres du tissus"
    Lang(53, 0) = "Largeur du tissus"
    Lang(54, 0) = "Largeur poche de lattes"
    Lang(55, 0) = "Largeur retour de bords"
    Lang(56, 0) = "Couleur 1"
    Lang(57, 0) = "couleur 2"
    Lang(58, 0) = "couleur 3"
    Lang(59, 0) = "Valider"
             
    

    
    'menu1 english
    
    Lang(0, 1) = "Welcome to all sailmakers"
    Lang(1, 1) = "This software for designing junk sails is made available for evaluation purpose and without any warranty under"
    Lang(2, 1) = "GNU General Public Licence version 2, as published by the Free Software Foundation."
    Lang(3, 1) = ""
    Lang(4, 1) = "Copyright and other rights reserved"
    Lang(5, 1) = "Sailcut is a registered trademark by Robert Lainé"
    Lang(6, 1) = "Begin"
    Lang(7, 1) = "End"
    Lang(8, 1) = "E-mail: robert.laine@sailcut.com"
    Lang(9, 1) = "  Geometry : modifying the sail - "
    Lang(10, 1) = "  Geometry of a new sail"
    Lang(11, 1) = "Luff width of upper panels"
    Lang(12, 1) = "Luff width of lower panels"
    Lang(13, 1) = "Leech width of lower panels"
    Lang(14, 1) = "Batten length"
    Lang(15, 1) = "Yard Length"
    Lang(16, 1) = "Foot angle"
    Lang(17, 1) = "Yard angle"
    Lang(18, 1) = "Area"
    Lang(19, 1) = "Depth"
    Lang(20, 1) = "Depth position"
    Lang(21, 1) = "Twist"
    Lang(22, 1) = "Head panels"
    Lang(23, 1) = "Bottom panels"
    Lang(24, 1) = "Sail file name"
    Lang(25, 1) = "Development"
    Lang(26, 1) = "Quit"
    Lang(27, 1) = "&File"
    Lang(28, 1) = "&New"
    Lang(29, 1) = "&Open file"
    Lang(30, 1) = "Panels DXF"
    Lang(31, 1) = "&Print"
    Lang(32, 1) = "Data"
    Lang(33, 1) = "Panels XY"
    Lang(34, 1) = "&Save"
    Lang(35, 1) = "Data"
    Lang(36, 1) = "Sail DXF"
    Lang(37, 1) = "VRML 1.0"
    Lang(38, 1) = "&Quit"
    Lang(39, 1) = "Clo&Th"
    Lang(40, 1) = "Open"
    Lang(41, 1) = "Save"
    Lang(42, 1) = "Credits"
    Lang(43, 1) = "Developers"
    Lang(44, 1) = "Translations"
    Lang(45, 1) = "French"
    Lang(46, 1) = "English"
    Lang(47, 1) = "Norwegian"
    Lang(48, 1) = "Dutch"
    Lang(49, 1) = "German"
    Lang(50, 1) = "Spanish"
    Lang(51, 1) = "Finnish"
    Lang(52, 1) = "Cloth parameters"
    Lang(53, 1) = "Cloth width"
    Lang(54, 1) = "Batten pocket width"
    Lang(55, 1) = "Tabling width"
    Lang(56, 1) = "Color 1"
    Lang(57, 1) = "Color 2"
    Lang(58, 1) = "Color 3"
    Lang(59, 1) = "Validation"
    
    'Norwegian Rolf Nilsen <rolf@moldenett.no>
    'menu1
    Lang(0, 2) = "Seilmakere, velkommen"
    Lang(1, 2) = "Denne programvaren for å designe junke riggede seil er kun tilgjengelig for evalueringsformål, og kommer uten garanti for funksjonalitet eller resultat"
    Lang(2, 2) = "Programvaren gjøres tilgjengelig under GNU General Public Licence versjon 2, som publisert av Free Software Foundation"
    Lang(3, 2) = ""
    Lang(4, 2) = "Kopirettigheter og andre rettigheter reservert"
    Lang(5, 2) = "Sailcut er et registrert varemerke ved Robert Lainé"
    Lang(6, 2) = "Start"
    Lang(7, 2) = "Slutt"
    Lang(8, 2) = "Send e-post: robert.laine@sailcut.com"
      'create1
    Lang(9, 2) = "  Geometri: modifisering av seil - "
    Lang(10, 2) = "  Geometri for et nytt seil"
    Lang(11, 2) = "Forlik bredde på øvre paneler"
    Lang(12, 2) = "Forlik bredde på neder paneler"
    Lang(13, 2) = "Akterlikslengde på nedre paneler"
    Lang(14, 2) = "Lengde på spile"
    Lang(15, 2) = "Skjøtelengde"
    Lang(16, 2) = "Vinkel på fot"
    Lang(17, 2) = "Skjøtevinkel"
    Lang(18, 2) = "Areal"
    Lang(19, 2) = "Dybde"
    Lang(20, 2) = "Dybdens posisjon"
    Lang(21, 2) = "Tvist"
    Lang(22, 2) = "Topp paneler"
    Lang(23, 2) = "Bunn paneler"
    Lang(24, 2) = "Filnavn for seil"
    Lang(25, 2) = "Utvikling av panel"
    Lang(26, 2) = "Slutt"
    Lang(27, 2) = "&Fil"
    Lang(28, 2) = "&Ny"
    Lang(29, 2) = "Åpne fil"
    Lang(30, 2) = "Paneler DXF"
    Lang(31, 2) = "Utskrift"
    Lang(32, 2) = "Data"
    Lang(33, 2) = "Paneler XY"
    Lang(34, 2) = "Lagre"
    Lang(35, 2) = "Data"
    Lang(36, 2) = "Seil DXF"
    Lang(37, 2) = "VRML 1.0"
    Lang(38, 2) = "Slutt"
    Lang(39, 2) = "Materiale"
    Lang(40, 2) = "Åpne"
    Lang(41, 2) = "Lagre"
    Lang(42, 2) = "Bekreftelse"
    Lang(43, 2) = "Developers "
    Lang(44, 2) = "Translations"
    Lang(45, 2) = "French"
    Lang(46, 2) = "English"
    Lang(47, 2) = "Norwegian"
    Lang(48, 2) = "Dutch"
    Lang(49, 2) = "German"
    Lang(50, 2) = "Spanish"
    Lang(51, 2) = "Finnish"
    Lang(52, 2) = "Cloth parameters"
    Lang(53, 2) = "Cloth width"
    Lang(54, 2) = "Batten pocket width"
    Lang(55, 2) = "Tabling width"
    Lang(56, 2) = "Color 1"
    Lang(57, 2) = "Color 2"
    Lang(58, 2) = "Color 3"
    Lang(59, 2) = "Validation"
    
    
    
    'Dutch Tony Mels   <Tony@Hobie17.com>
    'menu1
    Lang(0, 3) = "Alle zeilmakers gegroet"
    Lang(1, 3) = "Deze software voor het ontwerpen van jonk zeilen is beschikbaar gesteld voor evaluatie doeleinden en zonder garantie onder"
    Lang(2, 3) = "de GNU General Public License versie 2, zoals gepubliceerd door de Free Software Foundation."
    Lang(3, 3) = ""
    Lang(4, 3) = "Handelsmerk en andere rechten voorbehouden"
    Lang(5, 3) = "Sailcut is een geregistreerd handelsmerk van Robert Lainé"
    Lang(6, 3) = "Begin"
    Lang(7, 3) = "Einde"
    Lang(8, 3) = "Stuur e-mail: robert.laine@sailcut.com"
    'Create1
    Lang(9, 3) = "  Geometrie : aanpassen van het zeil - "
    Lang(10, 3) = "  Geometrie van een nieuw zeil"
    Lang(11, 3) = "Voorlijk breedte van bovenste panelen"
    Lang(12, 3) = "Voorlijk breedte van onderste panelen"
    Lang(13, 3) = "Achterlijk breedte van onderste panelen"
    Lang(14, 3) = "Zeillat lengte"
    Lang(15, 3) = "Ra lengte"
    Lang(16, 3) = "Onderlijk hoek"
    Lang(17, 3) = "Ra hoek"
    Lang(18, 3) = "Oppervlakte"
    Lang(19, 3) = "Diepte"
    Lang(20, 3) = "Diepte positie"
    Lang(21, 3) = "Twist"
    Lang(22, 3) = "Voor panelen"
    Lang(23, 3) = "Onderste panelen"
    Lang(24, 3) = "Zeil bestandsnaam"
    Lang(25, 3) = "Ontwikkeling"
    Lang(26, 3) = "Afsluiten"
    Lang(27, 3) = "Bestand"
    Lang(28, 3) = "&Nieuw"
    Lang(29, 3) = "&Open bestand"
    Lang(30, 3) = "Panelen DXF"
    Lang(31, 3) = "Afdrukken"
    Lang(32, 3) = "Data"
    Lang(33, 3) = "Panelen XY"
    Lang(34, 3) = "Op&slaan"
    Lang(35, 3) = "Data"
    Lang(36, 3) = "Zeil DXF"
    Lang(37, 3) = "VRML 1.0"
    Lang(38, 3) = "Afsluiten"
    Lang(39, 3) = "Doek"
    Lang(40, 3) = "Open"
    Lang(41, 3) = "Opslaan"
    Lang(42, 3) = "Erkenning"
    Lang(43, 3) = "Developers"
    Lang(44, 3) = "Translations"
    Lang(45, 3) = "French"
    Lang(46, 3) = "English"
    Lang(47, 3) = "Norwegian"
    Lang(48, 3) = "Dutch"
    Lang(49, 3) = "German"
    Lang(50, 3) = "Spanish"
    Lang(51, 3) = "Finnish"
    Lang(52, 3) = "Cloth parameters"
    Lang(53, 3) = "Cloth width"
    Lang(54, 3) = "Batten pocket width"
    Lang(55, 3) = "Tabling width"
    Lang(56, 3) = "Color 1"
    Lang(57, 3) = "Color 2"
    Lang(58, 3) = "Color 3"
    Lang(59, 3) = "Validation"

    
    'German Leo Foltz <leo@leow.de>
    'menu1
     Lang(0, 4) = "Hallo Segelmacher"
    Lang(1, 4) = "Dieses Programm zum Entwerfen von Dschunkensegeln dient ausschliesslich Demonstrationszwecken und wird ohne jegliche Garantie zur Verfügung gestellt unter"
    Lang(2, 4) = "GNU General Public Licence version 2, wie veröffentlicht von der Free Software Foundation."
    Lang(3, 4) = ""
    Lang(4, 4) = "Alle Rechte vorbehalten"
    Lang(5, 4) = "Sailcut ist eingetragenes Warenzeichen von Robert Lainé"
    Lang(6, 4) = "Anfang"
    Lang(7, 4) = "Ende"
    Lang(8, 4) = "E-mail: robert.laine@sailcut.com"
    'create1
    Lang(9, 4) = "  Geometrie: Segel verändern - "
    Lang(10, 4) = "  Geometrie: Neues Segel"
    Lang(11, 4) = "Breite obere Paneele am Vorliek"
    Lang(12, 4) = "Breite untere Paneele am Vorliek"
    Lang(13, 4) = "Breite untere Paneele am Achterliek"
    Lang(14, 4) = "Länge Segellatten"
    Lang(15, 4) = "Länge Rah(Gaffel)"
    Lang(16, 4) = "Winkel am Hals"
    Lang(17, 4) = "Winkel der Rah(Gaffel)"
    Lang(18, 4) = "Fläche"
    Lang(19, 4) = "Tiefe max."
    Lang(20, 4) = "Position max. Tiefe"
    Lang(21, 4) = "Verwindung"
    Lang(22, 4) = "Obere Paneele"
    Lang(23, 4) = "Untere Paneele"
    Lang(24, 4) = "Segel-Dateiname"
    Lang(25, 4) = "Abwicklung"
    Lang(26, 4) = "Beenden"
    Lang(27, 4) = "Datei"
    Lang(28, 4) = "Neu"
    Lang(29, 4) = "Datei öffnen"
    Lang(30, 4) = "Paneele DXF"
    Lang(31, 4) = "Drucken"
    Lang(32, 4) = "Data"
    Lang(33, 4) = "Paneele XY"
    Lang(34, 4) = "Speichern"
    Lang(35, 4) = "Data"
    Lang(36, 4) = "Segel DXF"
    Lang(37, 4) = "VRML 1.0"
    Lang(38, 4) = "Beenden"
    Lang(39, 4) = "Segeltuch"
    Lang(40, 4) = "Öffnen"
    Lang(41, 4) = "Speichern"
    Lang(42, 4) = "Danksagung"
    Lang(43, 4) = "Entwickler"
    Lang(44, 4) = "Uebersetzungen"
    Lang(45, 4) = "Franzoesisch"
    Lang(46, 4) = "Englisch"
    Lang(47, 4) = "Norwegisch"
    Lang(48, 4) = "Niederlaendisch"
    Lang(49, 4) = "Deutsch"
    Lang(50, 4) = "Spanisch"
    Lang(51, 4) = "Finnisch"
    Lang(52, 4) = "Segeltuch-Parameter"
    Lang(53, 4) = "Segeltuch Breite"
    Lang(54, 4) = "Lattentaschen Breite"
    Lang(55, 4) = "Doppelung Breite"
    Lang(56, 4) = "Farbe 1"
    Lang(57, 4) = "Farbe 2"
    Lang(58, 4) = "Farbe 3"
    Lang(59, 4) = "Richtigkeit"
    
    'Spanish Joserra Mariño    <joserracat@wanadoo.es>
    'menu1
    Lang(0, 5) = "Bienvenida a todos los constructores de velas"
    Lang(1, 5) = "Este programa para diseñar velas chatarra está dis"
    Lang(2, 5) = "GNU General Public Licence version 2, como publicó"
    Lang(3, 5) = ""
    Lang(4, 5) = "Copyright y otros derechos reservados"
    Lang(5, 5) = "Sailcut es una marca registrada de Robert Lainé"
    Lang(6, 5) = "Inicio"
    Lang(7, 5) = "Fin"
    Lang(8, 5) = "E-mail: robert.laine@sailcut.com"
    'create1
    Lang(9, 5) = "  Geometria : modificando la vela - "
    Lang(10, 5) = "  Geometria de una nueva vela"
    Lang(11, 5) = "Ancho en el grátil de los paneles superiores"
    Lang(12, 5) = "Ancho en el grátil de los paneles inferiores"
    Lang(13, 5) = "Ancho en la baluma de los paneles inferiores"
    Lang(14, 5) = "Longitud del sable"
    Lang(15, 5) = "Longitud del Pico"
    Lang(16, 5) = "Angulo de pujamen"
    Lang(17, 5) = "Angulo del Pico"
    Lang(18, 5) = "Area"
    Lang(19, 5) = "profundidad o bolsa"
    Lang(20, 5) = "Posición de la bolsa"
    Lang(21, 5) = "Twist"
    Lang(22, 5) = "Paneles superiores"
    Lang(23, 5) = "Paneles inferiores"
    Lang(24, 5) = "Nombre de archivo de la vela"
    Lang(25, 5) = "Desarrollo"
    Lang(26, 5) = "Salir"
    Lang(27, 5) = "&Archivo"
    Lang(28, 5) = "&Nuevo"
    Lang(29, 5) = "&Abrir"
    Lang(30, 5) = "DXF de Paneles"
    Lang(31, 5) = "&Imprimir"
    Lang(32, 5) = "Datos"
    Lang(33, 5) = "Paneles XY"
    Lang(34, 5) = "Guardar"
    Lang(35, 5) = "Data"
    Lang(36, 5) = "DXF de la vela"
    Lang(37, 5) = "VRML 1.0"
    Lang(38, 5) = "&Salir"
    Lang(39, 5) = "&Tejido"
    Lang(40, 5) = "Abrir"
    Lang(41, 5) = "Guardar"
    Lang(42, 5) = "Agradecimientos"
    Lang(43, 5) = "Desarrolladores"
    Lang(44, 5) = "Traducciones"
    Lang(45, 5) = "Frances"
    Lang(46, 5) = "Ingles"
    Lang(47, 5) = "Noruego"
    Lang(48, 5) = "Holandés"
    Lang(49, 5) = "Aleman"
    Lang(50, 5) = "Español"
    Lang(51, 5) = "Finlandes"
    Lang(52, 5) = "Parametros de tejido"
    Lang(53, 5) = "Anchura del tejido"
    Lang(54, 5) = "Anchura del bolsillo del sable"
    Lang(55, 5) = "Anchura del tableado"
    Lang(56, 5) = "Color 1"
    Lang(57, 5) = "Color 2"
    Lang(58, 5) = "Color 3"
    Lang(59, 5) = "Validacion"
    
    'Finnish Terho Halme <terho.halme@luukku.com>
    'menu1
    Lang(0, 6) = "Tervetuloa kaikki purjeentekijät"
    Lang(1, 6) = "Tämä ohjelma on julkaistu kokeilukäyttöön dzonkkipurjeiden suunnittelua varten ilman takuita"
    Lang(2, 6) = "de GNU General Public License versie 2, zoals gepubliceerd door de Free Software Foundation."
    Lang(3, 6) = ""
    Lang(4, 6) = "Kaikki oikeudet pidätetään"
    Lang(5, 6) = "Sailcut on Robert Lainé:n rekisteröimä tavaramerkki"
    Lang(6, 6) = "Aloita"
    Lang(7, 6) = "Lopeta"
    Lang(8, 6) = "Sähköposti: robert.laine@sailcut.com"
    'create1
    Lang(9, 6) = "  Geometria : muokkaa purjetta - "
    Lang(10, 6) = "  Uuden purjeen geometria"
    Lang(11, 6) = "Yläpaneelien takaliesma"
    Lang(12, 6) = "Alapaneelien etuliesma"
    Lang(13, 6) = "Alapaneelien takaliesma"
    Lang(14, 6) = "Latan pituus"
    Lang(15, 6) = "Kahvelipuomin pituus"
    Lang(16, 6) = "Puomin kulma"
    Lang(17, 6) = "Kahvelipuomin kulma"
    Lang(18, 6) = "Pinta-ala"
    Lang(19, 6) = "Pussin syvyys"
    Lang(20, 6) = "Pussin paikka"
    Lang(21, 6) = "Kierto"
    Lang(22, 6) = "Yläpaneelit"
    Lang(23, 6) = "Alapaneelit"
    Lang(24, 6) = "Purjeen tiedostonimi"
    Lang(25, 6) = "Levitä"
    Lang(26, 6) = "Ulos"
    Lang(27, 6) = "Tiedosto"
    Lang(28, 6) = "Uusi"
    Lang(29, 6) = "Avaa tiedosto"
    Lang(30, 6) = "Paneelien DXF"
    Lang(31, 6) = "Tulosta"
    Lang(32, 6) = "Tiedot"
    Lang(33, 6) = "Paneelien XY"
    Lang(34, 6) = "Tallenna"
    Lang(35, 6) = "Tiedot"
    Lang(36, 6) = "Purjeen DXF"
    Lang(37, 6) = "VRML 1.0"
    Lang(38, 6) = "Ulos"
    Lang(39, 6) = "Kangas"
    Lang(40, 6) = "Avaa"
    Lang(41, 6) = "Tallenna"
    Lang(42, 6) = "Tekijäluettelo"
    Lang(43, 6) = "Kehittäjät"
    Lang(44, 6) = "Käännökset"
    Lang(45, 6) = "ranska"
    Lang(46, 6) = "englanti"
    Lang(47, 6) = "norja"
    Lang(48, 6) = "hollanti"
    Lang(49, 6) = "saksa"
    Lang(50, 6) = "espanja"
    Lang(51, 6) = "suomi"
    Lang(52, 6) = "Kankaan parametrit"
    Lang(53, 6) = "Kankaan leveys"
    Lang(54, 6) = "Lattataskun leveys"
    Lang(55, 6) = "Reunavahvikkeen leveys"
    Lang(56, 6) = "Väri 1"
    Lang(57, 6) = "Väri 1"
    Lang(58, 6) = "Väri 1"
    Lang(59, 6) = "Vahvistus"
    
    
End Sub
                 
                 
                 

Public Sub mSCVars_LoadArray()
   
   SCVars(1) = Str(Int(UpLuff))
   SCVars(2) = Str(Int(LoLuff))
   SCVars(3) = Str(Int(LoLeech))
   SCVars(4) = Str(Int(LBatten))
   SCVars(5) = Str(Int(LYard))
   SCVars(6) = Str(Int(AFoot))
   SCVars(7) = Str(Int(AYard))
   SCVars(8) = Format$(Surface, " 0.#0")
   SCVars(9) = Format$(Mdepth(0) * 100, " #0")
   SCVars(10) = Format$(RPdepth * 100, " #0")
   SCVars(11) = Str$(twistScrl)
   SCVars(12) = Str(nHpanel)
   SCVars(13) = Str(nBpanel)
   SCVars(14) = Sail$
End Sub


Public Sub mUnits_LoadArray()

'metric
    UnitCaption(1, 0) = " mm"
    UnitCaption(2, 0) = " mm"
    UnitCaption(3, 0) = " mm"
    UnitCaption(4, 0) = " mm"
    UnitCaption(5, 0) = " mm"
    UnitCaption(6, 0) = Chr(176)
    UnitCaption(7, 0) = Chr(176)
    UnitCaption(8, 0) = "m" & Chr(178)
    UnitCaption(9, 0) = " %"
    UnitCaption(10, 0) = " %"
    UnitCaption(11, 0) = Chr(176)
    UnitCaption(12, 0) = " panel"
    UnitCaption(13, 0) = " panel"
    UnitCaption(14, 0) = ".sc8"
    
'inches
    UnitCaption(1, 1) = ""
    UnitCaption(2, 1) = ""
    UnitCaption(3, 1) = ""
    UnitCaption(4, 1) = ""
    UnitCaption(5, 1) = ""
    UnitCaption(6, 1) = Chr(176)
    UnitCaption(7, 1) = Chr(176)
    UnitCaption(8, 1) = " ft" & Chr(178)
    UnitCaption(9, 1) = " %"
    UnitCaption(10, 1) = " %"
    UnitCaption(11, 1) = Chr(176)
    UnitCaption(12, 1) = " panel"
    UnitCaption(13, 1) = " panel"
    UnitCaption(14, 1) = ".sc8"
'feet
    UnitCaption(1, 2) = "  "
    UnitCaption(2, 2) = "  "
    UnitCaption(3, 2) = "  "
    UnitCaption(4, 2) = "  "
    UnitCaption(5, 2) = "  "
    UnitCaption(6, 2) = Chr(176)
    UnitCaption(7, 2) = Chr(176)
    UnitCaption(8, 2) = " ft" & Chr(178)
    UnitCaption(9, 2) = " %"
    UnitCaption(10, 2) = " %"
    UnitCaption(11, 2) = Chr(176)
    UnitCaption(12, 2) = " panel"
    UnitCaption(13, 2) = " panel"
    UnitCaption(14, 2) = ".sc8"
'feet-inches
    UnitCaption(1, 3) = ""
    UnitCaption(2, 3) = ""
    UnitCaption(3, 3) = ""
    UnitCaption(4, 3) = ""
    UnitCaption(5, 3) = ""
    UnitCaption(6, 3) = Chr(176)
    UnitCaption(7, 3) = Chr(176)
    UnitCaption(8, 3) = " ft" & Chr(178)
    UnitCaption(9, 3) = " %"
    UnitCaption(10, 3) = " %"
    UnitCaption(11, 3) = Chr(176)
    UnitCaption(12, 3) = " panel"
    UnitCaption(13, 3) = " panel"
    UnitCaption(14, 3) = ".sc8"
End Sub


Public Sub mCtlNames_LoadArray()

CtlNames(1) = "UpLuffScroll"
CtlNames(2) = "LoLuffScroll"
CtlNames(3) = "LoLeechScroll"
CtlNames(4) = "LBattenScroll"
CtlNames(5) = "LYardScroll"
CtlNames(6) = "FootAScroll"
CtlNames(7) = "YardAScroll"
CtlNames(8) = ""
CtlNames(9) = "DepthScroll"
CtlNames(10) = "RPdepthScroll"
CtlNames(11) = "TwistScroll"
CtlNames(12) = "NHpanelScroll"
CtlNames(13) = "NBpanelScroll"
End Sub


