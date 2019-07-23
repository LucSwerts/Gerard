Attribute VB_Name = "m___info"
' *******************************************************************************
' * ##### ##### #####  ###  ##### ####                      #####   #####   #   *
' * #     #     #   # #   # #   # #   #                         #       #  ##   *
' * # ### ###   ##### ##### ##### #   #   #####   #   #     #####   ##### # #   *
' * #   # #     #  #  #   # #  #  #   #            # #      #       #       #   *
' * ##### ##### #   # #   # #   # ####              #   #   ##### # ##### ##### *
' *******************************************************************************

'-----------------------------------------------------------------------------------------------
' Project       GERARD - Gericht Efficient Rechercheren met ANPR tegen Rondtrekkende Daders
' Doel          zoeken naar opvallende profielen van targets inzake WIB door rondtrekkende daders
'               gebruik makend van ANPR-gegevens
'               ANPR-gegevens importeren en exploiteren
'               schema met alle unieke nummers, overzicht per camera
'               Zoom-functie met Joker en Combi-functies
'               Tandem zoekt samenhorende voertuigen zonder voorafgaande kennis van Targets
' Module        m___Info
' Auteur        Luc Swerts - FGP Limburg
' Copyright ©   Luc Swerts - FGP Limburg
' Inhoud        Info
' References    Geen

' Revisies overzicht:
' Versie        Datum(jjjj-mm-dd)       Beschrijving
'-----------------------------------------------------------------------------------------------
' 01.00         2017-04-11              Eerste release
' 01.30         2018-02-28              DatumTijd splitsen voorzien in Lint
' 01.40         2018-09-12              nExtraTijd optioneel door cfgPlusToepassen
'                                       update frmInfo
' 01.50         2018-12-13              import van XLS
' 02.00         2018-12-30              vergelijking met Collection
' 02.10         2019-01-01              vergelijking van Collection met Collection
'                                       volgorde omgekeerd: zoek deel in geheel
'                                       volledige run 28 clusters van 43'02" => 1'37" = x26 sneller
' 02.11         2019-01-02              opkuis van resterende lege cellen in Schema
' 02.20         2019-02-28              sorteerfuncties ingebouwd, diverse sleutels, op en neer
' 02.21         2019-03-02              importscherm verbeterd, knop voor eigen map, wissel Excel/CSV afgewerkt
'                                       Werkbalk definitief afgevoerd
' 02.22         2019-03-03              Thema's ingevoerd
' 02.30         2019-03-05              Aanzet Tandem
' 02.33         2019-03-07              frmTandem
' 02.38         2019-03-11              centrale logging op netwerk
' 02.40         2019-05-08              algoritmes voor Schema
' 02.41         2019-05-09              Schema opbouw verbeterd, test met dossier Gaatje
'                                       24161 rijen van 2'15" naar 0'13" = snelheid x 10!
' 02.42         2019-05-11              mDC opgenomen
' 02.47         2019-05-17              frmImport: map kiezen ging 2x open, opgelost (.show x2)
' 02.48         2019-05-18              refactoring mSupport
'                                       frmInventaris toegevoegd, inventaris over alle sets heen
' 02.49         2019-05-21              inhoudstafel toegevoegd
' 02.54         2019-06-06              Tandem herwerken, testen op snelheid
'                                       gebruik van caption aangepast (Application & ActiveWindow)
' 02.55         2019-06-08              Tandem: TintAndShade, databalkjes voor interval en aantal voertuigen
' 02.56         2019-06-09              opmaak van tijd hersteld (zie v. 02.45) van char naar float
'                                       Tandem werkte niet meer, kon geen tijdsverschil berekenen
' 02.58         2019-06-11              Accentueren toegevoegd in Tandem
' 02.59         2019-06-11              frmPuzzel toegevoegd
' 02.62         2019-06-13              Refactoring frmPuzzel, sorteerroutines, ...
' 02.64         2019-06-18              Refactoring en boilerplate code
' 02.66         2019-07-02              handleiding schrijven en minimale Refactoring
' 02.67         2019-07-03              indeling Ribbon herschikt

' TODO: Tandem opbouwen op basis van alle Sets of alleen die in Schema
' TODO: gebruik techniek van Zoom
' TODO: VerzamelSets => IsDataSheet

'-----------------------------------------------------------------------------------------------

' ----------------------------------------------------------------------------------------------
' Succes en veel plezier met het gebruik van GERARD!!!                                         '
' Suggesties en opmerkingen zijn altijd welkom :-)                                             '
' ' => luc.swerts@police.belgium.eu                                                            '
' ----------------------------------------------------------------------------------------------
'
'-----------------------------------------------------------------------------------------------
' Einde module - m___Info                                                                      '
'-----------------------------------------------------------------------------------------------
