#This File is in Unicode format.  Do not edit in an ASCII editor.

<#
.SYNOPSIS
	Creates a complete inventory of a Microsoft Active Directory Forest using Microsoft Word.
.DESCRIPTION
	Creates a complete inventory of a Microsoft Active Directory Forest using Microsoft Word and PowerShell.
	Creates a Word document named after the Active Directory Forest.
	Document includes a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

	While most of the script can be run with a non admin account, there are some features 
	that will not or may not work without domain admin rights.  The Hardware and Services 
	parameters require domain admin privileges.  The count of all users may not be accurate
	if the user running the script doesn't have the necessary permissions on all user
	objects.  In that case, there may be user accounts classified as "unknown".
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
	If either registry key does not exist and this parameter is not specified, the report will
	not contain a Company Name on the cover page.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	(default cover pages in Word en-US)
	Valid input is:
		Alphabet (Word 2007/2010. Works)
		Annual (Word 2007/2010. Doesn't really work well for this report)
		Austere (Word 2007/2010. Works)
		Austin (Word 2010/2013. Doesn't work in 2013, mostly works in 2007/2010 but Subtitle/Subject & Author fields need to me moved after title box is moved up)
		Banded (Word 2013. Works)
		Conservative (Word 2007/2010. Works)
		Contrast (Word 2007/2010. Works)
		Cubicles (Word 2007/2010. Works)
		Exposure (Word 2007/2010. Works if you like looking sideways)
		Facet (Word 2013. Works)
		Filigree (Word 2013. Works)
		Grid (Word 2010/2013.Works in 2010)
		Integral (Word 2013. Works)
		Ion (Dark) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Ion (Light) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Mod (Word 2007/2010. Works)
		Motion (Word 2007/2010/2013. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2007/2010. Works)
		Puzzle (Word 2007/2010. Top date doesn't fit, box needs to be manually resized or font changed to 14 point)
		Retrospect (Word 2013. Works)
		Semaphore (Word 2013. Works)
		Sideline (Word 2007/2010/2013. Doesn't work in 2013, works in 2007/2010)
		Slice (Dark) (Word 2013. Doesn't work)
		Slice (Light) (Word 2013. Doesn't work)
		Stacks (Word 2007/2010. Works)
		Tiles (Word 2007/2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2007/2010. Works)
		ViewMaster (Word 2013. Works)
		Whisp (Word 2013. Works)
	Default value is Motion.
	This parameter has an alias of CP.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
.PARAMETER Hardware
	Use WMI to gather hardware information on: Computer System, Disks, Processor and Network Interface Cards
	Will be used on Domain Controllers only.
	This parameters requires the script be run from an elevated PowerShell session 
	using an account with permission to retrieve hardware information (i.e. Domain Admin).
	Selecting this parameter will add to both the time it takes to run the script and size of the report.
	This parameter is disabled by default.
.PARAMETER ADForest
	Specifies an Active Directory forest object by providing one of the following attribute values. 
	The identifier in parentheses is the LDAP display name for the attribute.

	Fully qualified domain name
		Example: corp.contoso.com
	GUID (objectGUID)
		Example: 599c3d2e-f72d-4d20-8a88-030d99495f20
	DNS host name
		Example: dnsServer.corp.contoso.com
	NetBIOS name
		Example: corp
		
	This parameter is required.
.PARAMETER ComputerName
	Specifies which domain controller to use to run the script against.
	If ADForest is a trusted forest then ComputerName is required to detect the existence of ADForest.
	ComputerName can be entered as the NetBIOS name, FQDN, localhost or IP Address.
	If entered as localhost, the actual computer name is determined and used.
	If entered as an IP address, an attempt is made to determine and use the actual computer name.
.PARAMETER Services
	Gather information on all services running on domain controllers.
	Servers that are configured to automatically start but are not running will be colored in red.
	Will be used on Domain Controllers only.
	This parameters requires the script be run from an elevated PowerShell session
	using an account with permission to retrieve service information (i.e. Domain Admin).
	Selecting this parameter will add to both the time it takes to run the script and size of the report.
	This parameter is disabled by default.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory.ps1 -ADForest company.tld
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Motion for the Cover Page format.
	Administrator for the User Name.
	company.tld for the AD Forest
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory.ps1 -PDF -ADForest corp.carlwebster.com
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Motion for the Cover Page format.
	Administrator for the User Name.
	corp.carlwebster.com for the AD Forest
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory.ps1 -hardware
	
	Will use all default values and add additional information for each domain controller about its hardware.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Motion for the Cover Page format.
	Administrator for the User Name.
	The user will be prompted for ADForest.
.EXAMPLE
	PS C:\PSScript .\ADDS_Inventory.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster" -ComputerName ADDC01

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
		The user will be prompted for ADForest.
		Domain Controller named ADDC01 for the ComputerName.
.EXAMPLE
	PS C:\PSScript .\ADDS_Inventory.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
		The user will be prompted for ADForest.
		The computer running the script for the ComputerName.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word or PDF document.
.NOTES
	NAME: ADDS_Inventory.ps1
	VERSION: 1.00
	AUTHOR: Carl Webster
	LASTEDIT: May 31, 2014
#>


#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding( SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "" ) ]

Param(
	[parameter(
	Position = 0, 
	Mandatory=$False )
	] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(
	Position = 1, 
	Mandatory=$False )
	] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(
	Position = 2, 
	Mandatory=$False )
	] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(
	Position = 3, 
	Mandatory=$False )
	] 
	[Switch]$PDF=$False,

	[parameter(
	Position = 4, 
	Mandatory=$False )
	] 
	[Switch]$Hardware=$False,

	[parameter(
	Position = 5, 
	Mandatory=$True )
	] 
	[string]$ADForest="", 

	[parameter(
	Position = 6, 
	Mandatory=$False )
	] 
	[string]$ComputerName="",
	
	[parameter(
	Position = 7, 
	Mandatory=$False )
	] 
	[Switch]$Services=$False)
	

#force -verbose on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
#since the Microsoft AD cmdlets do not honor -EA 0, set $ErrorActionPreference
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

If($PDF -eq $Null)
{
	$PDF = $False
}
If($Hardware -eq $Null)
{
	$Hardware = $False
}
If($Services -eq $Null)
{
	$Services = $False
}

#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on April 10, 2014

Set-StrictMode -Version 2

#the following values were attained from 
#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
[int]$wdAlignPageNumberRight = 2
[long]$wdColorGray15 = 14277081
[long]$wdColorGray05 = 15987699 
[int]$wdMove = 0
[int]$wdSeekMainDocument = 0
[int]$wdSeekPrimaryFooter = 4
[int]$wdStory = 6
[int]$wdColorRed = 255
[int]$wdColorBlack = 0
[int]$wdWord2007 = 12
[int]$wdWord2010 = 14
[int]$wdWord2013 = 15
[int]$wdSaveFormatPDF = 17
#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
[int]$wdAlignParagraphLeft = 0
[int]$wdAlignParagraphCenter = 1
[int]$wdAlignParagraphRight = 2
#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
[int]$wdCellAlignVerticalTop = 0
[int]$wdCellAlignVerticalCenter = 1
[int]$wdCellAlignVerticalBottom = 2
#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
[int]$wdAutoFitFixed = 0
[int]$wdAutoFitContent = 1
[int]$wdAutoFitWindow = 2
#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
[int]$wdAdjustNone = 0
[int]$wdAdjustProportional = 1
[int]$wdAdjustFirstColumn = 2
[int]$wdAdjustSameWidth = 3

[int]$PointsPerTabStop = 36
[int]$Indent0TabStops = 0 * $PointsPerTabStop
[int]$Indent1TabStops = 1 * $PointsPerTabStop
[int]$Indent2TabStops = 2 * $PointsPerTabStop
[int]$Indent3TabStops = 3 * $PointsPerTabStop
[int]$Indent4TabStops = 4 * $PointsPerTabStop

# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
[int]$wdStyleHeading1 = -2
[int]$wdStyleHeading2 = -3
[int]$wdStyleHeading3 = -4
[int]$wdStyleHeading4 = -5
[int]$wdStyleNoSpacing = -158
[int]$wdTableGrid = -155

#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
[int]$wdLineStyleNone = 0
[int]$wdLineStyleSingle = 1

[int]$wdHeadingFormatTrue = -1
[int]$wdHeadingFormatFalse = 0 

[string]$RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

$hash = @{}

# DE and FR translations for Word 2010 by Vladimir Radojevic
# Vladimir.Radojevic@Commerzreal.com

# DA translations for Word 2010 by Thomas Daugaard
# Citrix Infrastructure Specialist at edgemo A/S

# CA translations by Javier Sanchez 
# CEO & Founder 101 Consulting

#ca - Catalan
#da - Danish
#de - German
#en - English
#es - Spanish
#fi - Finnish
#fr - French
#nb - Norwegian
#nl - Dutch
#pt - Portuguese
#sv - Swedish

Switch ($PSCulture.Substring(0,3))
{
	'ca-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Taula automÃ¡tica 2';
			}
		}

	'da-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Automatisk tabel 2';
			}
		}

	'de-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Automatische Tabelle 2';
			}
		}

	'en-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents'  = 'Automatic Table 2';
			}
		}

	'es-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Tabla automÃ¡tica 2';
			}
		}

	'fi-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Automaattinen taulukko 2';
			}
		}

	'fr-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Sommaire Automatique 2';
			}
		}

	'nb-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Automatisk tabell 2';
			}
		}

	'nl-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Automatische inhoudsopgave 2';
			}
		}

	'pt-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'SumÃ¡rio AutomÃ¡tico 2';
			}
		}

	'sv-'	{
			$hash.($($PSCulture)) = @{
				'Word_TableOfContents' = 'Automatisk innehÃ¥llsfÃ¶rteckning2';
			}
		}

	Default	{$hash.('en-US') = @{
				'Word_TableOfContents'  = 'Automatic Table 2';
			}
		}
}

$myHash = $hash.$PSCulture

If($myHash -eq $Null)
{
	$myHash = $hash.('en-US')
}

$myHash.Word_NoSpacing = $wdStyleNoSpacing
$myHash.Word_Heading1 = $wdStyleheading1
$myHash.Word_Heading2 = $wdStyleheading2
$myHash.Word_Heading3 = $wdStyleheading3
$myHash.Word_Heading4 = $wdStyleheading4
$myHash.Word_TableGrid = $wdTableGrid

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP)
	
	$xArray = ""
	
	Switch ($PSCulture.Substring(0,3))
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "IÃ³ (clar)", "IÃ³ (fosc)", "LÃ­nia lateral",
					"Moviment", "QuadrÃ­cula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "SemÃ for", "VisualitzaciÃ³", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "DiplomÃ tic", "ExposiciÃ³",
					"LÃ­nia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "QuadrÃ­cula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Anual", "Conservador", "Contrast",
					"Cubicles", "DiplomÃ tic", "En mosaic", "ExposiciÃ³", "LÃ­nia lateral",
					"Mod", "Moviment", "Piles", "Sobri", "Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevÃ¦gElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mÃ¸rk)", "Ion (mÃ¸rk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevÃ¦gElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "GÃ¥de",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"NÃ¥lestribet", "Ã…rlig", "Avispapir", "Tradionel")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Ã…rlig", "BevÃ¦gElse", "Eksponering",
					"Enkel", "Firkanter", "Fliser", "GÃ¥de", "Kontrast",
					"Mod", "NÃ¥lestribet", "Overskrid", "Sidelinje", "Stakke",
					"Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "RÃ¼ckblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "JÃ¤hrlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Bewegung", "Durchscheinend", "Herausgestellt",
					"JÃ¤hrlich", "Kacheln", "Kontrast", "Kubistisch", "Modern",
					"Nadelstreifen", "Puzzle", "Randlinie", "Raster", "Schlicht", "Stapel",
					"Traditionell")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast",
					"Cubicles", "Exposure", "Mod", "Motion", "Pinstripes", "Puzzle",
					"Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "SemÃ¡foro", "Retrospectiva", "CuadrÃ­cula",
					"Movimiento", "Cortar (oscuro)", "LÃ­nea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "CuadrÃ­cula", "CubÃ­culos", "ExposiciÃ³n", "LÃ­nea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periÃ³dico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Conservador",
					"Contraste", "CubÃ­culos", "ExposiciÃ³n", "LÃ­nea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Pilas", "Puzzle",
					"Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "VaihtuvavÃ¤rinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Aakkoset", "Alttius", "Kontrasti", "Kuvakkeet ja tiedot",
					"Liike" , "Liituraita" , "Mod" , "Palapeli", "Perinteinen", "Pinot",
					"Sivussa", "TyÃ¶pisteet", "Vuosittainen", "Yksinkertainen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("ViewMaster", "Secteur (foncÃ©)", "SÃ©maphore",
					"RÃ©trospective", "Ion (foncÃ©)", "Ion (clair)", "IntÃ©grale",
					"Filigrane", "Facette", "Secteur (clair)", "Ã€ bandes", "Austin",
					"Guide", "Whisp", "Lignes latÃ©rales", "Quadrillage")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("MosaÃ¯ques", "Ligne latÃ©rale", "Annuel", "Perspective",
					"Contraste", "Emplacements de bureau", "Moderne", "Blocs empilÃ©s",
					"Rayures fines", "AustÃ¨re", "Transcendant", "Classique", "Quadrillage",
					"Exposition", "Alphabet", "Mots croisÃ©s", "Papier journal", "Austin", "Guide")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Annuel", "AustÃ¨re", "Blocs empilÃ©s", "Blocs superposÃ©s",
					"Classique", "Contraste", "Exposition", "Guide", "Ligne latÃ©rale", "Moderne",
					"MosaÃ¯ques", "Mots croisÃ©s", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mÃ¸rk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mÃ¸rk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Ã…rlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Ã…rlig", "Avlukker", "BevegElse", "Engasjement",
					"Enkel", "Fliser", "Konservativ", "Kontrast", "Mod", "Puslespill",
					"Sidelinje", "Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Bescheiden", "Beweging",
					"Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks", "Krijtstreep",
					"Mod", "Puzzel", "Stapels", "Tegels", "Terzijde", "Werkplekken")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("AnimaÃ§Ã£o", "Austin", "Em Tiras", "ExibiÃ§Ã£o Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana",
					"Grade", "Integral", "Ãon (Claro)", "Ãon (Escuro)", "Linha Lateral",
					"Retrospectiva", "SemÃ¡foro")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "AnimaÃ§Ã£o", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "ExposiÃ§Ã£o", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeÃ§a", "Transcend")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabeto", "AnimaÃ§Ã£o", "Anual", "Austero", "Baias", "Conservador",
					"Contraste", "ExposiÃ§Ã£o", "Ladrilhos", "Linha Lateral", "Listras", "Mod",
					"Pilhas", "Quebra-cabeÃ§a", "Transcendente")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mÃ¶rkt)", "Knippe", "RutnÃ¤t", "RÃ¶rElse", "Sektor (ljus)", "Sektor (mÃ¶rk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Ã…terblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("AlfabetmÃ¶nster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "RutnÃ¤t",
					"RÃ¶rElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Ã…rligt",
					"Ã–vergÃ¥ende")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("AlfabetmÃ¶nster", "Ã…rligt", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Ã–vergÃ¥ende", "Plattor", "Pussel", "RÃ¶rElse",
					"Sidlinje", "Sobert", "Staplat")
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid", "Integral",
						"Ion (Dark)", "Ion (Light)", "Motion", "Retrospect", "Semaphore",
						"Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster", "Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast",
						"Cubicles", "Exposure", "Mod", "Motion", "Pinstripes", "Puzzle",
						"Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function GetComputerWMIInfo
{
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com
	# modified 1-May-2014 to work in trusted AD Forests and using different domain admin credentials	

	#Get Computer info
	Write-Verbose "$(Get-Date): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date): `t`t`tHardware information"
	WriteWordLine 3 0 "Computer Information"
	WriteWordLine 0 1 "General Computer"
	[bool]$GotComputerItems = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_computersystem
	}
	
	Catch
	{
		$Results = $Null
	}
	
	If($? -and $Results -ne $Null)
	{
		$ComputerItems = $Results | Select Manufacturer, Model, Domain, @{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}
		$Results = $Null

		ForEach($Item in $ComputerItems)
		{
			WriteWordLine 0 2 "Manufacturer`t: " $Item.manufacturer
			WriteWordLine 0 2 "Model`t`t: " $Item.model
			WriteWordLine 0 2 "Domain`t`t: " $Item.domain
			WriteWordLine 0 2 "Total Ram`t: $($Item.totalphysicalram) GB"
			WriteWordLine 0 2 ""
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		WriteWordLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
		WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
		WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
		WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results returned for Computer information"
		WriteWordLine 0 2 "No results returned for Computer information" "" $Null 0 $False $True
	}
	
	#Get Disk info
	Write-Verbose "$(Get-Date): `t`t`tDrive information"
	WriteWordLine 0 1 "Drive(s)"
	[bool]$GotDrives = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName Win32_LogicalDisk
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Results -ne $Null)
	{
		$drives = $Results | Select caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
		ForEach($drive in $drives)
		{
			If($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				WriteWordLine 0 2 "Caption`t`t: " $drive.caption
				WriteWordLine 0 2 "Size`t`t: $($drive.drivesize) GB"
				If(![String]::IsNullOrEmpty($drive.filesystem))
				{
					WriteWordLine 0 2 "File System`t: " $drive.filesystem
				}
				WriteWordLine 0 2 "Free Space`t: $($drive.drivefreespace) GB"
				If(![String]::IsNullOrEmpty($drive.volumename))
				{
					WriteWordLine 0 2 "Volume Name`t: " $drive.volumename
				}
				If(![String]::IsNullOrEmpty($drive.volumedirty))
				{
					WriteWordLine 0 2 "Volume is Dirty`t: " -nonewline
					If($drive.volumedirty)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
				}
				If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
				{
					WriteWordLine 0 2 "Volume Serial #`t: " $drive.volumeserialnumber
				}
				WriteWordLine 0 2 "Drive Type`t: " -nonewline
				Switch ($drive.drivetype)
				{
					0	{WriteWordLine 0 0 "Unknown"}
					1	{WriteWordLine 0 0 "No Root Directory"}
					2	{WriteWordLine 0 0 "Removable Disk"}
					3	{WriteWordLine 0 0 "Local Disk"}
					4	{WriteWordLine 0 0 "Network Drive"}
					5	{WriteWordLine 0 0 "Compact Disc"}
					6	{WriteWordLine 0 0 "RAM Disk"}
					Default {WriteWordLine 0 0 "Unknown"}
				}
				WriteWordLine 0 2 ""
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		WriteWordLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
		WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
		WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
		WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results returned for Drive information"
		WriteWordLine 0 2 "No results returned for Drive information" "" $Null 0 $False $True
	}
	

	#Get CPU's and stepping
	Write-Verbose "$(Get-Date): `t`t`tProcessor information"
	WriteWordLine 0 1 "Processor(s)"
	[bool]$GotProcessors = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_Processor
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Results -ne $Null)
	{
		$Processors = $Results | Select availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
		ForEach($processor in $processors)
		{
			WriteWordLine 0 2 "Name`t`t`t: " $processor.name
			WriteWordLine 0 2 "Description`t`t: " $processor.description
			WriteWordLine 0 2 "Max Clock Speed`t: $($processor.maxclockspeed) MHz"
			If($processor.l2cachesize -gt 0)
			{
				WriteWordLine 0 2 "L2 Cache Size`t`t: $($processor.l2cachesize) KB"
			}
			If($processor.l3cachesize -gt 0)
			{
				WriteWordLine 0 2 "L3 Cache Size`t`t: $($processor.l3cachesize) KB"
			}
			If($processor.numberofcores -gt 0)
			{
				WriteWordLine 0 2 "# of Cores`t`t: " $processor.numberofcores
			}
			If($processor.numberoflogicalprocessors -gt 0)
			{
				WriteWordLine 0 2 "# of Logical Procs`t: " $processor.numberoflogicalprocessors
			}
			WriteWordLine 0 2 "Availability`t`t: " -nonewline
			Switch ($processor.availability)
			{
				1	{WriteWordLine 0 0 "Other"}
				2	{WriteWordLine 0 0 "Unknown"}
				3	{WriteWordLine 0 0 "Running or Full Power"}
				4	{WriteWordLine 0 0 "Warning"}
				5	{WriteWordLine 0 0 "In Test"}
				6	{WriteWordLine 0 0 "Not Applicable"}
				7	{WriteWordLine 0 0 "Power Off"}
				8	{WriteWordLine 0 0 "Off Line"}
				9	{WriteWordLine 0 0 "Off Duty"}
				10	{WriteWordLine 0 0 "Degraded"}
				11	{WriteWordLine 0 0 "Not Installed"}
				12	{WriteWordLine 0 0 "Install Error"}
				13	{WriteWordLine 0 0 "Power Save - Unknown"}
				14	{WriteWordLine 0 0 "Power Save - Low Power Mode"}
				15	{WriteWordLine 0 0 "Power Save - Standby"}
				16	{WriteWordLine 0 0 "Power Cycle"}
				17	{WriteWordLine 0 0 "Power Save - Warning"}
				Default	{WriteWordLine 0 0 "Unknown"}
			}
			WriteWordLine 0 2 ""
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		WriteWordLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
		WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
		WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
		WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results returned for Processor information"
		WriteWordLine 0 2 "No results returned for Processor information" "" $Null 0 $False $True
	}

	#Get Nics
	Write-Verbose "$(Get-Date): `t`t`tNIC information"
	WriteWordLine 0 1 "Network Interface(s)"
	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration
	}
	
	Catch
	{
		$Results
	}

	If($? -and $Results -ne $Null)
	{
		$Nics = $Results | Where {$_.ipaddress -ne $Null}
		$Results = $Null

		If($Nics -eq $Null ) 
		{ 
			$GotNics = $False 
		} 
		Else 
		{ 
			$GotNics = !($Nics.__PROPERTY_COUNT -eq 0) 
		} 
	
		If($GotNics)
		{
			ForEach($nic in $nics)
			{
				Try
				{
					$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | Where {$_.index -eq $nic.index}
				}
				
				Catch 
				{
					$ThisNic = $Null
				}
				
				If($? -and $ThisNic -ne $Null)
				{
					If($ThisNic.Name -eq $nic.description)
					{
						WriteWordLine 0 2 "Name`t`t`t: " $ThisNic.Name
					}
					Else
					{
						WriteWordLine 0 2 "Name`t`t`t: " $ThisNic.Name
						WriteWordLine 0 2 "Description`t`t: " $nic.description
					}
					WriteWordLine 0 2 "Connection ID`t`t: " $ThisNic.NetConnectionID
					WriteWordLine 0 2 "Manufacturer`t`t: " $ThisNic.manufacturer
					WriteWordLine 0 2 "Availability`t`t: " -nonewline
					Switch ($ThisNic.availability)
					{
						1	{WriteWordLine 0 0 "Other"}
						2	{WriteWordLine 0 0 "Unknown"}
						3	{WriteWordLine 0 0 "Running or Full Power"}
						4	{WriteWordLine 0 0 "Warning"}
						5	{WriteWordLine 0 0 "In Test"}
						6	{WriteWordLine 0 0 "Not Applicable"}
						7	{WriteWordLine 0 0 "Power Off"}
						8	{WriteWordLine 0 0 "Off Line"}
						9	{WriteWordLine 0 0 "Off Duty"}
						10	{WriteWordLine 0 0 "Degraded"}
						11	{WriteWordLine 0 0 "Not Installed"}
						12	{WriteWordLine 0 0 "Install Error"}
						13	{WriteWordLine 0 0 "Power Save - Unknown"}
						14	{WriteWordLine 0 0 "Power Save - Low Power Mode"}
						15	{WriteWordLine 0 0 "Power Save - Standby"}
						16	{WriteWordLine 0 0 "Power Cycle"}
						17	{WriteWordLine 0 0 "Power Save - Warning"}
						Default	{WriteWordLine 0 0 "Unknown"}
					}
					WriteWordLine 0 2 "Physical Address`t: " $nic.macaddress
					WriteWordLine 0 2 "IP Address`t`t: " $nic.ipaddress
					WriteWordLine 0 2 "Default Gateway`t: " $nic.Defaultipgateway
					WriteWordLine 0 2 "Subnet Mask`t`t: " $nic.ipsubnet
					If($nic.dhcpenabled)
					{
						$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
						$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
						WriteWordLine 0 2 "DHCP Enabled`t`t: " $nic.dhcpenabled
						WriteWordLine 0 2 "DHCP Lease Obtained`t: " $dhcpleaseobtaineddate
						WriteWordLine 0 2 "DHCP Lease Expires`t: " $dhcpleaseexpiresdate
						WriteWordLine 0 2 "DHCP Server`t`t:" $nic.dhcpserver
					}
					If(![String]::IsNullOrEmpty($nic.dnsdomain))
					{
						WriteWordLine 0 2 "DNS Domain`t`t: " $nic.dnsdomain
					}
					If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
					{
						[int]$x = 1
						WriteWordLine 0 2 "DNS Search Suffixes`t:" -nonewline
						$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
						ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
						{
							If($x -eq 1)
							{
								$x = 2
								WriteWordLine 0 0 " $($DNSDomain)"
							}
							Else
							{
								WriteWordLine 0 5 " $($DNSDomain)"
							}
						}
					}
					WriteWordLine 0 2 "DNS WINS Enabled`t: " -nonewline
					If($nic.dnsenabledforwinsresolution)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
					{
						[int]$x = 1
						WriteWordLine 0 2 "DNS Servers`t`t:" -nonewline
						$nicdnsserversearchorder = $nic.dnsserversearchorder
						ForEach($DNSServer in $nicdnsserversearchorder)
						{
							If($x -eq 1)
							{
								$x = 2
								WriteWordLine 0 0 " $($DNSServer)"
							}
							Else
							{
								WriteWordLine 0 5 " $($DNSServer)"
							}
						}
					}
					WriteWordLine 0 2 "NetBIOS Setting`t`t: " -nonewline
					Switch ($nic.TcpipNetbiosOptions)
					{
						0	{WriteWordLine 0 0 "Use NetBIOS setting from DHCP Server"}
						1	{WriteWordLine 0 0 "Enable NetBIOS"}
						2	{WriteWordLine 0 0 "Disable NetBIOS"}
						Default	{WriteWordLine 0 0 "Unknown"}
					}
					WriteWordLine 0 2 "WINS:"
					WriteWordLine 0 3 "Enabled LMHosts`t: " -nonewline
					If($nic.winsenablelmhostslookup)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
					{
						WriteWordLine 0 3 "Host Lookup File`t: " $nic.winshostlookupfile
					}
					If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
					{
						WriteWordLine 0 3 "Primary Server`t`t: " $nic.winsprimaryserver
					}
					If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
					{
						WriteWordLine 0 3 "Secondary Server`t: " $nic.winssecondaryserver
					}
					If(![String]::IsNullOrEmpty($nic.winsscopeid))
					{
						WriteWordLine 0 3 "Scope ID`t`t: " $nic.winsscopeid
					}
				}
				ElseIf(!$?)
				{
					Write-Warning "$(Get-Date): Error retrieving NIC information"
					Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					WriteWordLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
					WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
					WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
					WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
					WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
				}
				Else
				{
					Write-Verbose "$(Get-Date): No results returned for NIC information"
					WriteWordLine 0 2 "No results returned for NIC information" "" $Null 0 $False $True
				}
			}
		}	
	}
	ElseIf(!$?)
	{
		Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
		Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		WriteWordLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
		WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
		WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
		WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
		WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results returned for NIC configuration information"
		WriteWordLine 0 2 "No results returned for NIC configuration information" "" $Null 0 $False $True
	}
	
	WriteWordLine 0 0 ""
	$Results = $Null
	$ComputerItems = $Null
	$Drives = $Null
	$Processors = $Null
	$Nics = $Null
}

Function GetComputerServices 
{
	Param([string]$RemoteComputerName)
	
	#Get Computer services info
	Write-Verbose "$(Get-Date): `t`tProcessing Computer services information"
	WriteWordLine 3 0 "Services"

	Try
	{
		$DCServices = Get-Service -ComputerName $RemoteComputerName | Sort DisplayName
	}
	
	Catch
	{
		$DCServices = $Null
	}
	
	If($? -and $DCServices -ne $Null)
	{
		[int]$Columns = 3
		If($DCServices -is [array])
		{
			[int]$Rows = $DCServices.count + 1
			[int]$NumServices = $DCServices.count
		}
		Else
		{
			[int]$Rows = 2
			[int]$NumServices = 1
		}
		Write-Verbose "$(Get-Date): `t`t $NumServices Services found"
		WriteWordLine 0 1 "Services ($NumServices Services found)"
		$TableRange = $doc.Application.Selection.Range
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.Style = $myHash.Word_TableGrid
		$Table.rows.first.headingformat = $wdHeadingFormatTrue
		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
		[int]$xRow = 1
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Display Name"
		$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,2).Range.Font.Bold = $True
		$Table.Cell($xRow,2).Range.Text = "Status"
		$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,3).Range.Font.Bold = $True
		$Table.Cell($xRow,3).Range.Text = "Startup Type"
		ForEach($Service in $DCServices)
		{
			$xRow++
			$Table.Cell($xRow,1).Range.Text = $Service.DisplayName
			
			Try
			{
				$StartupType = (Get-WMIObject Win32_Service -Filter "Name='$($Service.Name)'" -ComputerName $RemoteComputerName).StartMode
			}
			
			Catch
			{
				$StartupType = "Not Found"
			}
			
			If($Service.Status -eq "Stopped" -and $StartupType -eq "Auto")
			{
				$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorRed
				$Table.Cell($xRow,2).Range.Font.Bold  = $True
				$Table.Cell($xRow,2).Range.Font.Color = $WDColorBlack
			}
			$Table.Cell($xRow,2).Range.Text = $Service.Status
			$Table.Cell($xRow,3).Range.Text = $StartupType
		}

		$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
		$Table.AutoFitBehavior($wdAutoFitContent)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null
	}
	ElseIf(!$?)
	{
		Write-Warning "No services were retrieved."
		WriteWordLine 0 0 "Warning: No Services were retrieved" "" $Null 0 $False $True
		WriteWordLine 0 1 "If this is a trusted Forest, you may need to rerun the" "" $Null 0 $False $True
		WriteWordLine 0 1 "script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
	}
	Else
	{
		Write-Warning "Services retrieval was successful but no services were returned."
		WriteWordLine 0 0 "Services retrieval was successful but no services were returned." "" $Null 0 $False $True
	}
	$DCServices = $Null
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		Exit
	}
}

Function CheckWord2007SaveAsPDFInstalled
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Installer\Products\000021090B0090400000000000F01FEC) -eq $False)
	{
		Write-Host "`n`n`t`tWord 2007 is detected and the option to SaveAs PDF was selected but the Word 2007 SaveAs PDF add-in is not installed."
		Write-Host "`n`n`t`tThe add-in can be downloaded from http://www.microsoft.com/en-us/download/details.aspx?id=9943"
		Write-Host "`n`n`t`tInstall the SaveAs PDF add-in and rerun the script."
		Return $False
	}
	Return $True
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}
	
Function Check-LoadedModule
#Function created by Jeff Wouters
#@JeffWouters on Twitter
#modified by Michael B. Smith to handle when the module doesn't exist on server
#modified by @andyjmorgan
#bug fixed by @schose
#This Function handles all three scenarios:
#
# 1. Module is already imported into current session
# 2. Module is not already imported into current session, it does exists on the server and is imported
# 3. Module does not exist on the server

{
	Param([parameter(Mandatory = $True)][alias("Module")][string]$ModuleName)
	#$LoadedModules = Get-Module | Select Name
	#following line changed at the recommendation of @andyjmorgan
	$LoadedModules = Get-Module | % { $_.Name.ToString() }
	#bug reported on 21-JAN-2013 by @schose 
	#the following line did not work if the citrix.grouppolicy.commands.psm1 module
	#was manually loaded from a non Default folder
	#$ModuleFound = (!$LoadedModules -like "*$ModuleName*")
	$ModuleFound = ($LoadedModules -like "*$ModuleName*")
	If(!$ModuleFound) 
	{
		$module = Import-Module -Name $ModuleName -PassThru -EA 0
		If($module -and $?)
		{
			# module imported properly
			Return $True
		}
		Else
		{
			# module import failed
			Return $False
		}
	}
	Else
	{
		#module already imported into current session
		Return $True
	}
}

Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Selection.Style = $myHash.Word_NoSpacing}
		1 {$Selection.Style = $myHash.Word_Heading1}
		2 {$Selection.Style = $myHash.Word_Heading2}
		3 {$Selection.Style = $myHash.Word_Heading3}
		4 {$Selection.Style = $myHash.Word_Heading4}
		Default {$Selection.Style = $myHash.Word_NoSpacing}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Selection.TypeParagraph()
	}
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop = $properties | ForEach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
		If($propname -eq $Name) 
		{
			Return $_
		}
	} #ForEach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
}

Function AbortScript
{
	$Word.quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Function BuildMultiColumnTable
{
	Param([Array]$xArray, [String]$xType)
	
	#divide by 0 bug reported 9-Apr-2014 by Lee Dehmer 
	#if security group name or OU name was longer than 60 characters it caused a divide by 0 error
	
	#added a second parameter to the function so the verbose message would say whether 
	#the function is processing servers, security groups or OUs.
	
	If(-not ($xArray -is [Array]))
	{
		$xArray = (,$xArray)
	}
	[int]$MaxLength = 0
	[int]$TmpLength = 0
	#remove 60 as a hard-coded value
	#60 is the max width the table can be when indented 36 points
	[int]$MaxTableWidth = 60
	ForEach($xName in $xArray)
	{
		$TmpLength = $xName.Length
		If($TmpLength -gt $MaxLength)
		{
			$MaxLength = $TmpLength
		}
	}
	$TableRange = $doc.Application.Selection.Range
	#removed hard-coded value of 60 and replace with MaxTableWidth variable
	[int]$Columns = [Math]::Floor($MaxTableWidth / $MaxLength)
	If($xArray.count -lt $Columns)
	{
		[int]$Rows = 1
		#not enough array items to fill columns so use array count
		$MaxCells  = $xArray.Count
		#reset column count so there are no empty columns
		$Columns   = $xArray.Count 
	}
	ElseIf($Columns -eq 0)
	{
		#divide by 0 bug if this condition is not handled
		#number was larger than $MaxTableWidth so there can only be one column
		#with one cell per row
		[int]$Rows = $xArray.count
		$Columns   = 1
		$MaxCells  = 1
	}
	Else
	{
		[int]$Rows = [Math]::Floor( ( $xArray.count + $Columns - 1 ) / $Columns)
		#more array items than columns so don't go past last column
		$MaxCells  = $Columns
	}
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$Table.Style = $myHash.Word_TableGrid
	$Table.Borders.InsideLineStyle = $wdLineStyleSingle
	$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
	[int]$xRow = 1
	[int]$ArrayItem = 0
	While($xRow -le $Rows)
	{
		For($xCell=1; $xCell -le $MaxCells; $xCell++)
		{
			$Table.Cell($xRow,$xCell).Range.Text = $xArray[$ArrayItem]
			$ArrayItem++
		}
		$xRow++
	}
	$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
	$Table.AutoFitBehavior($wdAutoFitContent)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
	$TableRange = $Null
	$Table = $Null
	$xArray = $Null
}

Function UserIsaDomainAdmin
{
	#function adapted from sample code provided by Thomas Vuylsteke
	$IsDA = $False
	$name = $env:username
	Write-Verbose "$(Get-Date): TokenGroups - Checking groups for $name"

	$root = [ADSI]""
	$filter = "(sAMAccountName=$name)"
	$props = @("distinguishedName")
	$Searcher = new-Object System.DirectoryServices.DirectorySearcher($root,$filter,$props)
	$account = $Searcher.FindOne().properties.distinguishedname

	$user = [ADSI]"LDAP://$Account"
	$user.GetInfoEx(@("tokengroups"),0)
	$groups = $user.Get("tokengroups")

	$domainAdminsSID = New-Object System.Security.Principal.SecurityIdentifier (((Get-ADDomain -Server $ADForest).DomainSid).Value+"-512") 

	ForEach($group in $groups)
	{     
		$ID = New-Object System.Security.Principal.SecurityIdentifier($group,0)       
		If($ID.CompareTo($domainAdminsSID) -eq 0)
		{
			$IsDA = $True
		}     
	}

	$root = $Null
	$filter = $Null
	$props = $Null
	$Searcher = $Null
	$account = $Null
	$user = $Null
	$groups = $Null
	$domainAdminsSID = $Null
	Return $IsDA
}

Function GetComputerCountByOS
{
	Param([string]$xDomain)

	<#
	  This function will count the number of Windows workstations, Windows servers and
	  non-Windows computers and list them by Operating System.

	  Note that for servers we filter out Cluster Name Objects (CNOs) and
	  Virtual Computer Objects (VCOs) by checking the objects serviceprincipalname
	  property for a value of MSClusterVirtualServer. The CNO is the cluster
	  name, whereas a VCO is the client access point for the clustered role.
	  These are not actual computers, so we exlude them to assist with
	  accuracy.

	  Function Name: GetComputerCountByOS
	  Release: 1.0
	  Written by Jeremy@jhouseconsulting.com 20th May 2012
	#>

	#function optimized by Michael B. Smith
	
	Write-Verbose "$(Get-Date): `t`tGathering computer misc data"
	$Computers = @()
	$UnknownComputers = @()
	
	$Results = Get-ADComputer -Filter * -Properties Name,Operatingsystem,servicePrincipalName,DistinguishedName -Server $Domain
	
	If($? -and $Results -ne $Null)
	{
	
		Write-Verbose "$(Get-Date): `t`t`tGetting server OS counts"
		$Computers += $Results | `
			Where-Object {($_.Operatingsystem -like '*server*') -AND !($_.serviceprincipalname -like '*MSClusterVirtualServer*')} | `
			Sort-Object Name
		
		Write-Verbose "$(Get-Date): `t`t`tGetting workstation OS counts"
		$Computers += $Results | `
			Where-Object {($_.Operatingsystem -like '*windows*') -AND !($_.Operatingsystem -like '*server*')} | `
			Sort-Object Name
		
		Write-Verbose "$(Get-Date): `t`t`tGetting unknown OS counts"
		$UnknownComputers += $Results | `
			Where-Object {!($_.Operatingsystem -like '*windows*') -AND !($_.serviceprincipalname -like '*MSClusterVirtualServer*')} | `
			Sort-Object Name
		
		$Computers += $UnknownComputers
		$UnknownComputers = $UnknownComputers | Sort DistinguishedName
		
		$Computers = $Computers | Group-Object operatingsystem | Sort-Object Count -Descending

		Write-Verbose "$(Get-Date): `t`tBuild table for OS counts"
		WriteWordLine 3 0 "Windows Computer Operating Systems"
		$TableRange   = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($Computers -is [array])
		{
			[int]$Rows = $Computers.Count
		}
		Else
		{
			[int]$Rows = 1
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.Style = $myHash.Word_TableGrid
		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
		
		[int]$xRow = 0
		
		ForEach($Computer in $Computers)
		{
			$xRow++
			[string]$CountStr = "{0,7:N0}" -f $Computer.Count
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			If([String]::IsNullOrEmpty($Computer.Name))
			{
				$Table.Cell($xRow,1).Range.Text = "<No OS name>"
			}
			Else
			{
				$Table.Cell($xRow,1).Range.Text = $Computer.Name
			}
			$Table.Cell($xRow,2).Range.ParagraphFormat.Alignment = $wdCellAlignVerticalTop
			$Table.Cell($xRow,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
			$Table.Cell($xRow,2).Range.Text = $CountStr
		}
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
		$Table.AutoFitBehavior($wdAutoFitContent)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null
		
		If($UnknownComputers -ne $Null)
		{
			Write-Verbose "$(Get-Date): `t`tBuild table for unknown computers"
			WriteWordLine 3 0 "Non-Windows Computer Operating Systems"
			$TableRange   = $doc.Application.Selection.Range
			[int]$Columns = 2
			If($UnknownComputers -is [array])
			{
				[int]$Rows = $UnknownComputers.Count + 1
			}
			Else
			{
				[int]$Rows = 2
			}
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$Table.AutoFitBehavior($wdAutoFitFixed)
			$Table.Style = $myHash.Word_TableGrid
			$Table.rows.first.headingformat = $wdHeadingFormatTrue
			$Table.Borders.InsideLineStyle = $wdLineStyleSingle
			$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
			$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(1,1).Range.Font.Bold = $True
			$Table.Cell(1,1).Range.Text = "Distinguished Name"
			$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(1,2).Range.Font.Bold = $True
			$Table.Cell(1,2).Range.Text = "Operating System"
			
			[int]$xRow = 1
			
			ForEach($Computer in $UnknownComputers)
			{
				$xRow++
				[string]$CountStr = "{0,7:N0}" -f $Computer.Count
				$Table.Cell($xRow,1).Range.Text = $Computer.DistinguishedName
				If([String]::IsNullOrEmpty($Computer.OperatingSystem))
				{
					$Table.Cell($xRow,2).Range.Text = "<No OS name>"
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $Computer.OperatingSystem
				}
			}
			
			#set column widths
			$xcols = $table.columns

			ForEach($xcol in $xcols)
			{
			    switch ($xcol.Index)
			    {
				  1 {$xcol.width = 400}
				  2 {$xcol.width = 100}
			    }
			}

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
			$Table.AutoFitBehavior($wdAutoFitFixed)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "Error retrieving computer data for domain $($xDomain)"
		WriteWordLine 0 0 "Error retrieving computer data for domain $($xDomain)" "" $Null 0 $False $True
	}
	Else
	{
		WriteWordLine 0 0 "No computer data was retrieved for domain $($xDomain)" "" $Null 0 $False $True
	}
	$Computers = $Null
	$UnknownComputers = $Null
	$Results = $Null
}

#Script begins

$script:startTime = Get-Date

CheckWordPreReq

#If hardware inventory or services are requested, make sure user is running the script with Domain Admin rights
$DARights = $False
If($Hardware -or $Services)
{
	Write-Verbose "$(Get-Date): `tTesting to see if $($env:username) has Domain Admin rights"
	If($Hardware -and -not $Services)
	{
		Write-Verbose "$(Get-Date): Hardware inventory requested"
	}
	ElseIf($Services -and -not $Hardware)
	{
		Write-Verbose "$(Get-Date): Services requested"
	}
	ElseIf($Hardware -and $Services)
	{
		Write-Verbose "$(Get-Date): Hardware inventory and Services requested"
	}

	If(UserIsaDomainAdmin)
	{
		#user has Domain Admin rights
		Write-Verbose "$(Get-Date): $($env:username) has Domain Admin rights in the $($ADForest) Forest"
		$DARights = $True
	}
	Else
	{
		#user does not have Domain Admin rights
		If($Hardware -and -not $Services)
		{
			#don't abort script, set $hardware to false
			Write-Warning "`n`n`t`tHardware inventory was requested but $($WindowsIdentity.Name) does not have Domain Admin rights."
			Write-Warning "`n`n`t`tHardware inventory option will be turned off."
			$Hardware = $False
		}
		ElseIf($Services -and -not $Hardware)
		{
			#don't abort script, set $services to false
			Write-Warning "`n`n`t`tServices were requested but $($WindowsIdentity.Name) does not have Domain Admin rights."
			Write-Warning "`n`n`t`tServices option will be turned off."
			$Services = $False
		}
		ElseIf($Hardware -and $Services)
		{
			#don't abort script, set $hardware and $services to false
			Write-Warning "`n`n`t`tHardware inventory and Services were requested but $($WindowsIdentity.Name) does not have Domain Admin rights."
			Write-Warning "`n`n`t`tHardware inventory and Services options will be turned off."
			$Hardware = $False
			$Services = $False
		}
	}
}

#make sure ActiveDirectory module is loaded
If(!(Check-LoadedModule "ActiveDirectory"))
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Error "`n`n`t`tThe ActiveDirectory module could not be loaded.`nScript cannot continue.`n`n"
	Exit
}

If(![String]::IsNullOrEmpty($ComputerName)) 
{
	#get server name
	#first test to make sure the server is reachable
	Write-Verbose "$(Get-Date): Testing to see if $($ComputerName) is online and reachable"
	If(Test-Connection -ComputerName $ComputerName -quiet)
	{
		Write-Verbose "$(Get-Date): Server $($ComputerName) is online."
		Write-Verbose "$(Get-Date): `tTesting to see if it is a Domain Controller."
		#the server may be online but is it really a domain controller?

		#is the ComputerName in the current domain
		$Results = Get-ADDomainController $ComputerName
		
		If(!$?)
		{
			#try using the Forest name
			$Results = Get-ADDomainController $ComputerName -Server $ADForest
			If(!$?)
			{
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "`n`n`t`t$($ComputerName) is not a domain controller for $($ADForest).`n`t`tScript cannot continue.`n`n"
				Exit
			}
		}
		$Results = $Null
	}
	Else
	{
		Write-Verbose "$(Get-Date): Computer $($ComputerName) is offline"
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tComputer $($ComputerName) is offline.`nScript cannot continue.`n`n"
		Exit
	}
}

#if computer name is localhost, get actual server name
If($ComputerName -eq "localhost")
{
	$ComputerName = $env:ComputerName
	Write-Verbose "$(Get-Date): Computer name has been renamed from localhost to $($ComputerName)"
}

#if computer name is an IP address, get host name from DNS
#http://blogs.technet.com/b/gary/archive/2009/08/29/resolve-ip-addresses-to-hostname-using-powershell.aspx
#help from Michael B. Smith
$ip = $ComputerName -as [System.Net.IpAddress]
If($ip)
{
	$Result = [System.Net.Dns]::gethostentry($ip)
	
	If($? -and $Result -ne $Null)
	{
		$ComputerName = $Result.HostName
		Write-Verbose "$(Get-Date): Computer name has been renamed from $($ip) to $($ComputerName)"
	}
	Else
	{
		Write-Warning "Unable to resolve $($ComputerName) to a hostname"
	}
}
Else
{
	#server is online but for some reason $ComputerName cannot be converted to a System.Net.IpAddress
}

#get forest information so output filename can be generated
Write-Verbose "$(Get-Date): Testing to see if $($ADForest) is a valid forest name"
If([String]::IsNullOrEmpty($ComputerName))
{
	$Forest = Get-ADForest -Identity $ADForest
	
	If(!$?)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tCould not find a forest identified by: $($ADForest).`nScript cannot continue.`n`n"
		Exit
	}
}
Else
{
	$Forest = Get-ADForest -Identity $ADForest -Server $ComputerName

	If(!$?)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tCould not find a forest with the name of $($ADForest).`n`n`t`tScript cannot continue.`n`n`t`tIs $($ComputerName) running Active Directory Web Services?"
		Exit
	}
}
Write-Verbose "$(Get-Date): $($ADForest) is a valid forest name"
#store root domain so it only has to be accessed once
[string]$ForestRootDomain = $Forest.RootDomain

[string]$ForestName = $Forest.Name
[string]$Title      = "Inventory Report for the $($ForestName) Forest"
[string]$filename1  = "$($pwd.path)\$($ForestName).docx"
If($PDF)
{
	[string]$filename2 = "$($pwd.path)\$($ForestName).pdf"
}

Write-Verbose "$(Get-Date): Setting up Word"

# Setup word for output
Write-Verbose "$(Get-Date): Create Word comObject.  If you are not running Word 2007, ignore the next message."
$Word = New-Object -comobject "Word.Application" -EA 0

If(!$? -or $Word -eq $Null)
{
	Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
	$ErrorActionPreference = $SaveEAPreference
	Write-Error "`n`n`t`tThe Word object could not be created.  You may need to repair your Word installation.`n`n`t`tScript cannot continue.`n`n"
	Exit
}

[int]$WordVersion = [int] $Word.Version
If($WordVersion -eq $wdWord2013)
{
	$WordProduct = "Word 2013"
}
ElseIf($WordVersion -eq $wdWord2010)
{
	$WordProduct = "Word 2010"
}
ElseIf($WordVersion -eq $wdWord2007)
{
	$WordProduct = "Word 2007"
}
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
	AbortScript
}

If($PDF -and $WordVersion -eq $wdWord2007)
{
	Write-Verbose "$(Get-Date): Verify the Word 2007 Save As PDF add-in is installed"
	If(CheckWord2007SaveAsPDFInstalled)
	{
		Write-Verbose "$(Get-Date): The Word 2007 Save As PDF add-in is installed"
	}
	Else
	{
		AbortScript
	}
}

Write-Verbose "$(Get-Date): Validate company name"
#only validate CompanyName if the field is blank
If([String]::IsNullOrEmpty($CompanyName))
{
	$CompanyName = ValidateCompanyName
	If([String]::IsNullOrEmpty($CompanyName))
	{
		Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
		Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
		Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
	}
}

Write-Verbose "$(Get-Date): Check Default Cover Page for language specific version"
[bool]$CPChanged = $False
Switch ($PSCulture.Substring(0,3))
{
	'ca-'	{
			If($CoverPage -eq "Sideline")
			{
				$CoverPage = "LÃ­nia lateral"
				$CPChanged = $True
			}
		}

	'da-'	{
			If($CoverPage -eq "Sideline")
			{
				$CoverPage = "Sidelinje"
				$CPChanged = $True
			}
		}

	'de-'	{
			If($CoverPage -eq "Sideline")
			{
				$CoverPage = "Randlinie"
				$CPChanged = $True
			}
		}

	'es-'	{
			If($CoverPage -eq "Sideline")
			{
				$CoverPage = "LÃ­nea lateral"
				$CPChanged = $True
			}
		}

	'fi-'	{
			If($CoverPage -eq "Sideline")
			{
				$CoverPage = "Sivussa"
				$CPChanged = $True
			}
		}

	'fr-'	{
			If($CoverPage -eq "Sideline")
			{
				If($WordVersion -eq $wdWord2013)
				{
					$CoverPage = "Lignes latÃ©rales"
					$CPChanged = $True
				}
				Else
				{
					$CoverPage = "Ligne latÃ©rale"
					$CPChanged = $True
				}
			}
		}

	'nb-'	{
			If($CoverPage -eq "Sideline")
			{
				$CoverPage = "Sidelinje"
				$CPChanged = $True
			}
		}

	'nl-'	{
			If($CoverPage -eq "Sideline")
			{
				$CoverPage = "Terzijde"
				$CPChanged = $True
			}
		}

	'pt-'	{
			If($CoverPage -eq "Sideline")
			{
				$CoverPage = "Linha Lateral"
				$CPChanged = $True
			}
		}

	'sv-'	{
			If($CoverPage -eq "Sideline")
			{
				$CoverPage = "Sidlinje"
				$CPChanged = $True
			}
		}
}

If($CPChanged)
{
	Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
}

Write-Verbose "$(Get-Date): Validate cover page"
[bool]$ValidCP = ValidateCoverPage $WordVersion $CoverPage
If(!$ValidCP)
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Error "`n`n`t`tFor $WordProduct, $CoverPage is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
	AbortScript
}

Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): Company Name : $CompanyName"
Write-Verbose "$(Get-Date): Cover Page   : $CoverPage"
Write-Verbose "$(Get-Date): User Name    : $UserName"
Write-Verbose "$(Get-Date): Save As PDF  : $PDF"
Write-Verbose "$(Get-Date): HW Inventory : $Hardware"
Write-Verbose "$(Get-Date): Services     : $Services"
Write-Verbose "$(Get-Date): Forest Name  : $ADForest"
Write-Verbose "$(Get-Date): Title        : $Title"
Write-Verbose "$(Get-Date): Filename1    : $filename1"
If($PDF)
{
	Write-Verbose "$(Get-Date): Filename2    : $filename2"
}
Write-Verbose "$(Get-Date): OS Detected  : $RunningOS"
Write-Verbose "$(Get-Date): PSUICulture  : $PSUICulture"
Write-Verbose "$(Get-Date): PSCulture    : $PSCulture"
Write-Verbose "$(Get-Date): Word version : $WordProduct"
Write-Verbose "$(Get-Date): Word language: $($Word.Language)"
Write-Verbose "$(Get-Date): PoSH version : $($Host.Version)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): Script start : $($Script:StartTime)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): "

$Word.Visible = $False

#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
#using Jeff's Demo-WordReport.ps1 file for examples
Write-Verbose "$(Get-Date): Load Word Templates"

[bool]$CoverPagesExist = $False
[bool]$BuildingBlocksExist = $False

$word.Templates.LoadBuildingBlocks()
If($WordVersion -eq $wdWord2007)
{
	$BuildingBlocks = $word.Templates | Where {$_.name -eq "Building Blocks.dotx"}
}
Else
{
	#word 2010/2013
	$BuildingBlocks = $word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}
}

Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
$part = $Null

If($BuildingBlocks -ne $Null)
{
	$BuildingBlocksExist = $True

	Try 
	{
		$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
	}

	Catch
	{
		$part = $Null
	}

	If($part -ne $Null)
	{
		$CoverPagesExist = $True
	}
}

If(!$CoverPagesExist)
{
	Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
	Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
	Write-Warning "This report will not have a Cover Page."
}

Write-Verbose "$(Get-Date): Create empty word doc"
$Doc = $Word.Documents.Add()
If($Doc -eq $Null)
{
	Write-Verbose "$(Get-Date): "
	$ErrorActionPreference = $SaveEAPreference
	Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
	AbortScript
}

$Selection = $Word.Selection
If($Selection -eq $Null)
{
	Write-Verbose "$(Get-Date): "
	$ErrorActionPreference = $SaveEAPreference
	Write-Error "`n`n`t`tAn unknown error happened selecting the entire Word document for default formatting options.`n`n`t`tScript cannot continue.`n`n"
	AbortScript
}

#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
#36 = .50"
$Word.ActiveDocument.DefaultTabStop = 36

#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
Write-Verbose "$(Get-Date): Disable grammar and spell checking"
#bug reported 1-Apr-2014 by Tim Mangan
#save current options first before turning them off
$CurrentGrammarOption = $Word.Options.CheckGrammarAsYouType
$CurrentSpellingOption = $Word.Options.CheckSpellingAsYouType
$Word.Options.CheckGrammarAsYouType = $False
$Word.Options.CheckSpellingAsYouType = $False

If($BuildingBlocksExist)
{
	#insert new page, getting ready for table of contents
	Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
	$part.Insert($selection.Range,$True) | Out-Null
	$selection.InsertNewPage()

	#table of contents
	Write-Verbose "$(Get-Date): Table of Contents - $($myHash.Word_TableOfContents)"
	$toc = $BuildingBlocks.BuildingBlockEntries.Item($myHash.Word_TableOfContents)
	If($toc -eq $Null)
	{
		Write-Verbose "$(Get-Date): "
		Write-Verbose "$(Get-Date): Table of Content - $($myHash.Word_TableOfContents) could not be retrieved."
		Write-Warning "This report will not have a Table of Contents."
	}
	Else
	{
		$toc.insert($selection.Range,$True) | Out-Null
	}
}
Else
{
	Write-Verbose "$(Get-Date): Table of Contents are not installed."
	Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
}

#set the footer
Write-Verbose "$(Get-Date): Set the footer"
[string]$footertext = "Report created by $username"

#get the footer
Write-Verbose "$(Get-Date): Get the footer and format font"
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
#get the footer and format font
$footers = $doc.Sections.Last.Footers
ForEach ($footer in $footers) 
{
	If($footer.exists) 
	{
		$footer.range.Font.name = "Calibri"
		$footer.range.Font.size = 8
		$footer.range.Font.Italic = $True
		$footer.range.Font.Bold = $True
	}
} #end ForEach
Write-Verbose "$(Get-Date): Footer text"
$selection.HeaderFooter.Range.Text = $footerText

#add page numbering
Write-Verbose "$(Get-Date): Add page numbering"
$selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

#return focus to main document
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

#move to the end of the current document
Write-Verbose "$(Get-Date):"
$selection.EndKey($wdStory,$wdMove) | Out-Null
#end of Jeff Hicks 
[gc]::collect() 
######################START OF BUILDING REPORT

#Forest information

#set naming context
$ConfigNC = (Get-ADRootDSE -Server $ADForest).ConfigurationNamingContext

Write-Verbose "$(Get-Date): Writing forest data"

$selection.InsertNewPage()
WriteWordLine 1 0 "Forest Information"

Switch ($Forest.ForestMode)
{
	"Windows2000Forest"        {$ForestMode = "Windows 2000"}
	"Windows2003InterimForest" {$ForestMode = "Windows Server 2003 interim"}
	"Windows2003Forest"        {$ForestMode = "Windows Server 2003"}
	"Windows2008Forest"        {$ForestMode = "Windows Server 2008"}
	"Windows2008R2Forest"      {$ForestMode = "Windows Server 2008 R2"}
	"Windows2012Forest"        {$ForestMode = "Windows Server 2012"}
	"Windows2012R2Forest"      {$ForestMode = "Windows Server 2012 R2"}
	"UnknownForest"            {$ForestMode = "Unknown Forest Mode"}
	Default                    {$ForestMode = "Unable to determine Forest Mode: $($Forest.ForestMode)"}
}

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 2
[int]$Rows = 12
[int]$xRow = 1
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$Table.AutoFitBehavior($wdAutoFitFixed)
$Table.Style = $myHash.Word_TableGrid
$Table.Borders.InsideLineStyle = $wdLineStyleSingle
$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "Forest mode"
$Table.Cell($xRow,2).Range.Text = $ForestMode

$xRow++
$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "Forest name"
$Table.Cell($xRow,2).Range.Text = $Forest.Name

$xRow++
$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "Root domain"
$Table.Cell($xRow,2).Range.Text = $ForestRootDomain

$xRow++
$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "Domain naming master"
$Table.Cell($xRow,2).Range.Text = $Forest.DomainNamingMaster

$xRow++
$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "Schema master"
$Table.Cell($xRow,2).Range.Text = $Forest.SchemaMaster

$xRow++
$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "Partitions container"
$Table.Cell($xRow,2).Range.Text = $Forest.PartitionsContainer

Write-Verbose "$(Get-Date): `tApplication partitions"
$xRow++
$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "Application partitions"

$AppPartitions = $Forest.ApplicationPartitions | Sort
If($AppPartitions -eq $Null)
{
	$Table.Cell($xRow,2).Range.Text = "<None>"
}
Else
{
	$tmp = @()
	ForEach($AppPartition in $AppPartitions)
	{
		$tmp += "$($AppPartition)`r"
	}
	$Table.Cell($xRow,2).Range.Text = $tmp
}
$AppPartitions = $Null
$tmp = $Null

Write-Verbose "$(Get-Date): `tCross forest references"
$xRow++
$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "Cross forest references"

$CrossForestReferences = $Forest.CrossForestReferences | Sort
If($CrossForestReferences -eq $Null)
{
	$Table.Cell($xRow,2).Range.Text = "<None>"
}
Else
{
	$tmp = @()
	ForEach($CrossForestReference in $CrossForestReferences)
	{
		$tmp += "$($CrossForestReference)`r"
	}
	$Table.Cell($xRow,2).Range.Text = $tmp
}
$CrossForestReferences = $Null
$tmp = $Null

Write-Verbose "$(Get-Date): `tSPN suffixes"
$xRow++
$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "SPN suffixes"
$SPNSuffixes = $Forest.SPNSuffixes | Sort
If($SPNSuffixes -eq $Null)
{
	$Table.Cell($xRow,2).Range.Text = "<None>"
}
Else
{
	$tmp = @()
	ForEach($SPNSuffix in $SPNSuffixes)
	{
		$tmp += "$($SPNSuffix)`r"
	}
	$Table.Cell($xRow,2).Range.Text = $tmp
}
$SPNSuffixes = $Null
$tmp = $Null

Write-Verbose "$(Get-Date): `tUPN suffixes"
$xRow++
$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "UPN Suffixes"
$UPNSuffixes = $Forest.UPNSuffixes | Sort
If($UPNSuffixes -eq $Null)
{
	$Table.Cell($xRow,2).Range.Text = "<None>"
}
Else
{
	$tmp = @()
	ForEach($UPNSuffix in $UPNSuffixes)
	{
		$tmp += "$($UPNSuffix)`r"
	}
	$Table.Cell($xRow,2).Range.Text = $tmp
}
$UPNSuffixes = $Null
$tmp = $Null

Write-Verbose "$(Get-Date): `tDomains in forest"
$xRow++
$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "Domains in forest"
$Domains = $Forest.Domains | Sort
If($Domains -eq $Null)
{
	$Table.Cell($xRow,2).Range.Text = "<None>"
}
Else
{
	#redo list of domains so forest root domain is listed first
	$tmpDomains = @($ForestRootDomain)
	$tmpDomains2 = @("$($ForestRootDomain)`r")
	ForEach($Domain in $Domains)
	{
		If($Domain -ne $ForestRootDomain)
		{
			$tmpDomains += "$($Domain)"
			$tmpDomains2 += "$($Domain)`r"
		}
	}
	
	$Domains = $tmpDomains
	$Table.Cell($xRow,2).Range.Text = $tmpDomains2
}

Write-Verbose "$(Get-Date): `tSites"
$xRow++
$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "Sites"
$Sites = $Forest.Sites | Sort
If($Sites -eq $Null)
{
	$Table.Cell($xRow,2).Range.Text = "<None>"
}
Else
{
	$tmp = @()
	ForEach($Site in $Sites)
	{
		$tmp += "$($Site)`r"
	}
	$Table.Cell($xRow,2).Range.Text = $tmp
}
$Sites = $Null
$tmp = $Null

#set column widths
$xcols = $table.columns

ForEach($xcol in $xcols)
{
    switch ($xcol.Index)
    {
	  1 {$xcol.width = 125}
	  2 {$xcol.width = 300}
    }
}

$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
$Table.AutoFitBehavior($wdAutoFitFixed)

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null
$TableRange = $Null
$Table = $Null

[gc]::collect() 

Write-Verbose "$(Get-Date): `tDomain controllers"
WriteWordLine 3 0 "Domain Controllers"
#get all DCs in the forest
#http://www.superedge.net/2012/09/how-to-get-ad-forest-in-powershell.html
#http://msdn.microsoft.com/en-us/library/vstudio/system.directoryservices.activedirectory.forest.getforest%28v=vs.90%29
$ADContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("forest", $ADForest) 
$Forest2 = [system.directoryservices.activedirectory.Forest]::GetForest($ADContext)
$AllDCs = $Forest2.domains | ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name} 
$AllDCs = $AllDCs | Sort
$ADContext = $Null
$Forest2 = $Null

If($AllDCs -eq $Null)
{
	WriteWordLine 0 0 "<None>"
}
Else
{
	$TableRange = $doc.Application.Selection.Range
	[int]$Columns = 3
	If($AllDCs -is [array])
	{
		[int]$Rows = $AllDCs.Count + 1
	}
	Else
	{
		[int]$Rows = 2
	}
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$Table.Style = $myHash.Word_TableGrid
	$Table.Borders.InsideLineStyle = $wdLineStyleSingle
	$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
	$Table.rows.first.headingformat = $wdHeadingFormatTrue
	$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell(1,1).Range.Font.Bold = $True
	$Table.Cell(1,1).Range.Text = "Name"
	$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell(1,2).Range.Font.Bold = $True
	$Table.Cell(1,2).Range.Text = "Global Catalog"
	$Table.Cell(1,3).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell(1,3).Range.Font.Bold = $True
	$Table.Cell(1,3).Range.Text = "Read-only"
	[int]$xRow = 1
	ForEach($DC in $AllDCs)
	{
		$DCName = $DC.SubString(0,$DC.IndexOf("."))
		$SrvName = $DC.SubString($DC.IndexOf(".")+1)
		$xRow++
		$Table.Cell($xRow,1).Range.Text = $DC
		
		$Results = Get-ADDomainController -Identity $DCName -Server $SrvName
		
		If($? -and $Results -ne $Null)
		{
			$Table.Cell($xRow,2).Range.Text = $Results.IsGlobalCatalog
			$Table.Cell($xRow,3).Range.Text = $Results.IsReadOnly
		}
		Else
		{
			$Table.Cell($xRow,2).Range.Text = "Unknown"
			$Table.Cell($xRow,3).Range.Text = "Unknown"
		}
		$Results = $Null
	}
	$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
	$Table.AutoFitBehavior($wdAutoFitContent)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
	$TableRange = $Null
	$Table = $Null
}
$AllDCs = $Null
[gc]::collect() 

#Site information
Write-Verbose "$(Get-Date): Writing sites and services data"

$selection.InsertNewPage()
WriteWordLine 1 0 "Site and Services"

#get site information
#some of the following was taken from
#http://blogs.msdn.com/b/adpowershell/archive/2009/08/18/active-directory-powershell-to-manage-sites-and-subnets-part-3-getting-site-and-subnets.aspx

$tmp = $Forest.PartitionsContainer
$ConfigurationBase = $tmp.SubString($tmp.IndexOf(",") + 1)
$Sites = $Null
$Sites = Get-ADObject -Filter 'ObjectClass -eq "site"' -SearchBase $ConfigurationBase -Properties Name, SiteObjectBl -Server $ADForest | Sort Name

$siteContainerDN = ("CN=Sites," + $configNC)

If($? -and $Sites -ne $Null)
{
	WriteWordLine 2 0 "Inter-Site Transports"
	Write-Verbose "$(Get-Date): `tProcessing Inter-Site Transports"
	#adapted from code provided by Jeremy Saunders
	# Report of all site links and related settings
	$AllSiteLinks = Get-ADObject -Searchbase $ConfigNC -Server $ComputerName `
	-Filter 'objectClass -eq "siteLink"' -Property Description, Options, Cost, ReplInterval, SiteList, Schedule `
	| Select-Object Name, Description, @{Name="SiteCount";Expression={$_.SiteList.Count}}, Cost, ReplInterval, `
	@{Name="Schedule";Expression={If($_.Schedule){If(($_.Schedule -Join " ").Contains("240")){"NonDefault"}Else{"24x7"}}Else{"24x7"}}}, `
	Options, SiteList, DistinguishedName
	
	If($? -and $AllSiteLinks -ne $Null)
	{
		ForEach($SiteLink in $AllSiteLinks)
		{
			Write-Verbose "$(Get-Date): `t`tProcessing site link $($SiteLink.Name)"
			$SiteLinkTypeDN = @()
			$SiteLinkTypeDN = $SiteLink.DistinguishedName.Split(",")
			$SiteLinkType = $SiteLinkTypeDN[1].SubString(3)
			$SitesInLink = ""
			$SiteLinkSiteList = $SiteLink.SiteList
			ForEach($xSite in $SiteLinkSiteList)
			{
				$tmp = $xSite.Split(",")
				$SitesInLink += "$($tmp[0].SubString(3))`r"
			}
			
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If([String]::IsNullOrEmpty($SiteLink.Description))
			{
				[int]$Rows = 7
			}
			Else
			{
				[int]$Rows = 8
			}
			[int]$xRow = 0
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$Table.Style = $myHash.Word_TableGrid
			$Table.rows.first.headingformat = $wdHeadingFormatTrue
			$Table.Borders.InsideLineStyle = $wdLineStyleSingle
			$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
			$xRow++
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Name"
			$Table.Cell($xRow,2).Range.Text = $SiteLink.Name
			If(![String]::IsNullOrEmpty($SiteLink.Description))
			{
				$xRow++
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Description"
				$Table.Cell($xRow,2).Range.Text = $SiteLink.Description
			}
			$xRow++
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Sites in Link"
			If($SitesInLink -ne " ")
			{
				$Table.Cell($xRow,2).Range.Text = $SitesInLink
			}
			Else
			{
				$Table.Cell($xRow,2).Range.Text = "<None>"
			}
			$xRow++
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Cost"
			$Table.Cell($xRow,2).Range.Text = $SiteLink.Cost
			$xRow++
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Replication Interval"
			$Table.Cell($xRow,2).Range.Text = $SiteLink.ReplInterval
			$xRow++
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Schedule"
			$Table.Cell($xRow,2).Range.Text = $SiteLink.Schedule
			$xRow++
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Options"
			If([String]::IsNullOrEmpty($SiteLink.Options))
			{
				$Table.Cell($xRow,2).Range.Text = "Change Notification is Disabled"
			}
			ElseIf($SiteLink.Options -eq "1" -or $SiteLink.Options -eq "5")
			{
				$Table.Cell($xRow,2).Range.Text = "Change Notification is Enabled"
			}
			Else
			{
				$Table.Cell($xRow,2).Range.Text = "Unknown"
			}
			$xRow++
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Type"
			$Table.Cell($xRow,2).Range.Text = $SiteLinkType
			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
			$Table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	$AllSiteLinks = $Null
	
	ForEach($Site in $Sites)
	{
		Write-Verbose "$(Get-Date): `tProcessing site $($Site.Name)"
		WriteWordLine 2 0 $Site.Name

		WriteWordLine 3 0 "Subnets"
		Write-Verbose "$(Get-Date): `t`tProcessing subnets"
		$subnetArray = New-Object -Type string[] -ArgumentList $Site.siteObjectBL.Count
		$i = 0
		$SitesiteObjectBL = $Site.siteObjectBL
		foreach ($subnetDN in $SitesiteObjectBL) 
		{
			$subnetName = $subnetDN.SubString(3, $subnetDN.IndexOf(",CN=Subnets,CN=Sites,") - 3)
			$subnetArray[$i] = $subnetName
			$i++
		}
		$subnetArray = $subnetArray | Sort
		If($subnetArray -eq $Null)
		{
			WriteWordLine 0 0 "<None>"
		}
		Else
		{
			BuildMultiColumnTable $subnetArray "Subnets"
		}
		
		Write-Verbose "$(Get-Date): `t`tProcessing servers"
		WriteWordLine 3 0 "Servers"
		$siteName = $Site.Name
		
		#build array of connect objects
		Write-Verbose "$(Get-Date): `t`t`tProcessing automatic connection objects"
		$Connections = @()
		$ConnnectionObjects = $Null
		$ConnectionObjects = Get-ADObject -Filter 'objectClass -eq "nTDSConnection" -and options -bor 1' -Searchbase $ConfigNC -Property DistinguishedName, fromServer -Server $ADForest
		
		If($? -and $ConnectionObjects -ne $Null)
		{
			ForEach($ConnectionObject in $ConnectionObjects)
			{
				$xArray = $ConnectionObject.DistinguishedName.Split(",")
				#server name is 3rd item in array (element 2)
				$ToServer = $xArray[2].SubString($xArray[2].IndexOf("=")+1) #get past the = sign
				$xArray = $ConnectionObject.FromServer.Split(",")
				#server name is 2nd item in array (element 1)
				$FromServer = $xArray[1].SubString($xArray[1].IndexOf("=")+1) #get past the = sign
				#site name is 4th item in array (element 3)
				$FromServerSite = $xArray[3].SubString($xArray[3].IndexOf("=")+1) #get past the = sign
				$xArray = $Null
				$obj = New-Object -TypeName PSObject
				$obj | Add-Member -MemberType NoteProperty -Name Name           -Value "<automatically generated>"
				$obj | Add-Member -MemberType NoteProperty -Name ToServer       -Value $ToServer
				$obj | Add-Member -MemberType NoteProperty -Name FromServer     -Value $FromServer
				$obj | Add-Member -MemberType NoteProperty -Name FromServerSite -Value $FromServerSite
				$Connections += $obj
			}
		}
		
		Write-Verbose "$(Get-Date): `t`t`tProcessing manual connection objects"
		$ConnectionObjects = $Null
		$ConnectionObjects = Get-ADObject -Filter 'objectClass -eq "nTDSConnection" -and -not options -bor 1' -Searchbase $ConfigNC -Property Name, DistinguishedName, fromServer -Server $ADForest
		
		If($? -and $ConnectionObjects -ne $Null)
		{
			ForEach($ConnectionObject in $ConnectionObjects)
			{
				$xArray = $ConnectionObject.DistinguishedName.Split(",")
				#server name is 3rd item in array (element 2)
				$ToServer = $xArray[2].SubString($xArray[2].IndexOf("=")+1) #get past the = sign
				$xArray = $ConnectionObject.FromServer.Split(",")
				#server name is 2nd item in array (element 1)
				$FromServer = $xArray[1].SubString($xArray[1].IndexOf("=")+1) #get past the = sign
				#site name is 4th item in array (element 3)
				$FromServerSite = $xArray[3].SubString($xArray[3].IndexOf("=")+1) #get past the = sign
				$xArray = $Null
				$obj = New-Object -TypeName PSObject
				$obj | Add-Member -MemberType NoteProperty -Name Name           -Value $ConnectionObject.Name
				$obj | Add-Member -MemberType NoteProperty -Name ToServer       -Value $ToServer
				$obj | Add-Member -MemberType NoteProperty -Name FromServer     -Value $FromServer
				$obj | Add-Member -MemberType NoteProperty -Name FromServerSite -Value $FromServerSite
				$Connections += $obj
			}
		}

		If($Connections -ne $Null)
		{
			$Connections = $Connections | Sort Name, ToServer, FromServer
		}
		
		#list each server
		$serverContainerDN = "CN=Servers,CN=" + $siteName + "," + $siteContainerDN
		$SiteServers = $Null
		$SiteServers = Get-ADObject -SearchBase $serverContainerDN -SearchScope OneLevel -Filter { objectClass -eq "Server" } -Properties "DNSHostName" -Server $ADForest | Select DNSHostName, Name | Sort DNSHostName
		
		If($? -and $SiteServers -ne $Null)
		{
			$First = $True
			ForEach($SiteServer in $SiteServers)
			{
				If(!$First)
				{
					WriteWordLine 0 0 ""
				}
				WriteWordLine 0 0 $SiteServer.DNSHostName
				#for each server list each connection object
				If($Connections -ne $Null)
				{
					$Results = $Connections | Where {$_.ToServer -eq $SiteServer.Name}

					If($? -and $Results -ne $Null)
					{
						WriteWordLine 0 1 "Connection Objects to source server $($SiteServer.Name)"
						$TableRange = $doc.Application.Selection.Range
						[int]$Columns = 3
						If($Results -is [array])
						{
							[int]$Rows = $Results.Count + 1
						}
						Else
						{
							[int]$Rows = 2
						}
						[int]$xRow = 1
						$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
						$Table.Style = $myHash.Word_TableGrid
						$Table.rows.first.headingformat = $wdHeadingFormatTrue
						$Table.Borders.InsideLineStyle = $wdLineStyleSingle
						$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
						$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell($xRow,1).Range.Font.Bold = $True
						$Table.Cell($xRow,1).Range.Text = "Name"
						$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell($xRow,2).Range.Font.Bold = $True
						$Table.Cell($xRow,2).Range.Text = "From Server"
						$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell($xRow,3).Range.Font.Bold = $True
						$Table.Cell($xRow,3).Range.Text = "From Site"
						ForEach($Result in $Results)
						{
							$xRow++
							$Table.Cell($xRow,1).Range.Text = $Result.Name
							$Table.Cell($xRow,2).Range.Text = $Result.FromServer
							$Table.Cell($xRow,3).Range.Text = $Result.FromServerSite
						}
						$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
						$Table.AutoFitBehavior($wdAutoFitContent)

						#return focus back to document
						$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

						#move to the end of the current document
						$selection.EndKey($wdStory,$wdMove) | Out-Null
						$TableRange = $Null
						$Table = $Null
					}
				}
				Else
				{
					WriteWordLine 0 3 "Connection Objects: "
					WriteWordLine 0 4 "<None>"
				}
				$First = $False
			}
		}
		ElseIf(!$?)
		{
			Write-Warning "No Site Servers were retrieved."
			WriteWordLine 0 0 "Warning: No Site Servers were retrieved" "" $Null 0 $False $True
		}
		Else
		{
			WriteWordLine 0 0 "No servers in this site"
		}
	}
}
ElseIf(!$?)
{
	Write-Warning "No Sites were retrieved."
	WriteWordLine 0 0 "Warning: No Sites were retrieved" "" $Null 0 $False $True
}
Else
{
	Write-Warning "There were no sites found to retrieve."
	WriteWordLine 0 0 "There were no sites found to retrieve" "" $Null 0 $False $True
}
$Sites = $Null
$siteContainerDN = $Null
$AllSiteLinks = $Null
$SiteLinkTypeDN = $Null
$SiteLinkType = $Null
$SitesInLink = $Null
$SiteLinkSiteList = $Null
$subnetArray = $Null
$subnetName = $Null
$connections = $Null
$ConnectionObjects = $Null
$SiteName = $Null
$serverContainerDN = $Null
$SiteServers = $Null
$Results = $Null

#domains
Write-Verbose "$(Get-Date): Writing domain data"

$selection.InsertNewPage()
WriteWordLine 1 0 "Domain Information"
$AllDomainControllers = @()
$First = $True
#http://technet.microsoft.com/en-us/library/bb125224(v=exchg.150).aspx
#http://support.microsoft.com/kb/556086/he
$SchemaVersionTable = @{ 
"13" = "Windows 2000"; 
"30" = "Windows 2003 RTM, SP1, SP2"; 
"31" = "Windows 2003 R2";
"44" = "Windows 2008"; 
"47" = "Windows 2008 R2";
"56" = "Windows Server 2012"
"69" = "Windows Server 2012 R2"
"4397" = "Exchange 2000 RTM"; 
"4406" = "Exchange 2000 SP3";
"6870" = "Exchange 2003 RTM, SP1, SP2"; 
"6936" = "Exchange 2003 SP3"; 
"10637" = "Exchange 2007 RTM";
"11116" = "Exchange 2007 SP1"; 
"14622" = "Exchange 2007 SP2, Exchange 2010 RTM";
"14625" = "Exchange 2007 SP3";
"14726" = "Exchange 2010 SP1";
"14732" = "Exchange 2010 SP2";
"14734" = "Exchange 2010 SP3";
"15137" = "Exchange 2013 RTM";
"15254" = "Exchange 2013 CU1";
"15281" = "Exchange 2013 CU2";
"15283" = "Exchange 2013 CU3";
"15292" = "Exchange 2013 SP1"
}

ForEach($Domain in $Domains)
{
	Write-Verbose "$(Get-Date): `tProcessing domain $($Domain)"

	$DomainInfo = Get-ADDomain -Identity $Domain
	
	If($? -and $DomainInfo -ne $Null)
	{
		If(!$First)
		{
			#put each domain, starting with the second, on a new page
			$selection.InsertNewPage()
		}
		
		If($Domain -eq $ForestRootDomain)
		{
			WriteWordLine 2 0 "$($Domain) (Forest Root)"
		}
		Else
		{
			WriteWordLine 2 0 $Domain
		}

		Switch ($DomainInfo.DomainMode)
		{
			"Windows2000Domain"   {$DomainMode = "Windows 2000"}
			"Windows2003Mixed"    {$DomainMode = "Windows Server 2003 mixed"}
			"Windows2003Domain"   {$DomainMode = "Windows Server 2003"}
			"Windows2008Domain"   {$DomainMode = "Windows Server 2008"}
			"Windows2008R2Domain" {$DomainMode = "Windows Server 2008 R2"}
			"Windows2012Domain"   {$DomainMode = "Windows Server 2012"}
			"Windows2012R2Domain" {$DomainMode = "Windows Server 2012 R2"}
			"UnknownForest"       {$DomainMode = "Unknown Domain Mode"}
			Default               {$DomainMode = "Unable to determine Domain Mode: $($ADDomain.DomainMode)"}
		}
		
		#http://blogs.technet.com/b/poshchap/archive/2014/03/07/ad-schema-version.aspx
		$ADSchemaInfo = $Null
		$ExchangeSchemaInfo = $Null
		
		$ADSchemaInfo = Get-ADObject (Get-ADRootDSE -Server $Domain).schemaNamingContext -Property objectVersion -Server $Domain
		
		If($? -and $ADSchemaInfo -ne $Null)
		{
			$ADSchemaVersion = $ADSchemaInfo.objectversion
			$ADSchemaVersionName = $SchemaVersionTable.Get_Item("$ADSchemaVersion")
			If($ADSchemaVersionName -eq $Null)
			{
				$ADSchemaVersionName = "Unknown"
			}
		}
		Else
		{
			$ADSchemaVersion = "Unknown"
			$ADSchemaVersionName = "Unknown"
		}
		
		If($Domain -eq $ForestRootDomain)
		{
			$ExchangeSchemaInfo = Get-ADObject "cn=ms-exch-schema-version-pt,cn=Schema,cn=Configuration,$($DomainInfo.DistinguishedName)" -properties rangeupper -Server $Domain

			If($? -and $ExchangeSchemaInfo -ne $Null)
			{
				$ExchangeSchemaVersion = $ExchangeSchemaInfo.rangeupper
				$ExchangeSchemaVersionName = $SchemaVersionTable.Get_Item("$ExchangeSchemaVersion")
				If($ExchangeSchemaVersionName -eq $Null)
				{
					$ExchangeSchemaVersionName = "Unknown"
				}
			}
			Else
			{
				$ExchangeSchemaVersion = "Unknown"
				$ExchangeSchemaVersionName = "Unknown"
			}
		}
		
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		[int]$Rows = 22
		If(![String]::IsNullOrEmpty($DomainInfo.ManagedBy))
		{
			$Rows++
		}
		
		If(![String]::IsNullOrEmpty($ExchangeSchemaInfo))
		{
			$Rows++
		}
		
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.AutoFitBehavior($wdAutoFitFixed)	
		$Table.Style = $myHash.Word_TableGrid
		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
		
		[int]$xRow = 1
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Domain mode"
		$Table.Cell($xRow,2).Range.Text = $DomainMode

		$xRow++
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Domain name"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.Name

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "NetBIOS name"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.NetBIOSName

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "DNS root"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.DNSRoot

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Distinguished name"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.DistinguishedName

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Infrastructure master"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.InfrastructureMaster

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "PDC Emulator"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.PDCEmulator

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "RID Master"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.RIDMaster

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Default computers container"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.ComputersContainer

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Default users container"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.UsersContainer

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Deleted objects container"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.DeletedObjectsContainer

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Domain controllers container"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.DomainControllersContainer

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Foreign security principals container"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.ForeignSecurityPrincipalsContainer

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Lost and Found container"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.LostAndFoundContainer

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Quotas container"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.QuotasContainer

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Systems container"
		$Table.Cell($xRow,2).Range.Text = $DomainInfo.SystemsContainer

		$xRow++		
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "AD Schema"
		$Table.Cell($xRow,2).Range.Text = "($($ADSchemaVersion)) - $($ADSchemaVersionName)"
		
		If(![String]::IsNullOrEmpty($ExchangeSchemaInfo))
		{
			$xRow++		
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Exchange Schema"
			$Table.Cell($xRow,2).Range.Text = "($($ExchangeSchemaVersion)) - $($ExchangeSchemaVersionName)"
		}
		
		If(![String]::IsNullOrEmpty($DomainInfo.ManagedBy))
		{
			$xRow++
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Managed by"
			$Table.Cell($xRow,2).Range.Text = $DomainInfo.ManagedBy
		}

		Write-Verbose "$(Get-Date): `t`tGetting Allowed DNS Suffixes"
		$xRow++
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Allowed DNS Suffixes"
		$DNSSuffixes = $DomainInfo.AllowedDNSSuffixes | Sort
		If($DNSSuffixes -eq $Null)
		{
			$Table.Cell($xRow,2).Range.Text = "<None>"
		}
		Else
		{
			$tmp = @()
			ForEach($DNSSuffix in $DNSSuffixes)
			{
				$tmp += "$($DNSSuffix)`r"
			}
			$Table.Cell($xRow,2).Range.Text = $tmp
		}
		$DNSSuffixes = $Null
		$tmp = $Null

		Write-Verbose "$(Get-Date): `t`tGetting Child domains"
		$xRow++
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Child domains"
		$ChildDomains = $DomainInfo.ChildDomains | Sort
		If($ChildDomains -eq $Null)
		{
			$Table.Cell($xRow,2).Range.Text = "<None>"
		}
		Else
		{
			$tmp = @()
			ForEach($ChildDomain in $ChildDomains)
			{
				$tmp += "$($ChildDomain)`r"
			}
			$Table.Cell($xRow,2).Range.Text = $tmp
		}
		$ChildDOmains = $Null
		$tmp = $Null

		Write-Verbose "$(Get-Date): `t`tGetting Read-only replica directory servers"
		$xRow++
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Read-only replica directory servers"
		$ReadOnlyReplicas = $DomainInfo.ReadOnlyReplicaDirectoryServers | Sort
		If($ReadOnlyReplicas -eq $Null)
		{
			$Table.Cell($xRow,2).Range.Text = "<None>"
		}
		Else
		{
			$tmp = @()
			ForEach($ReadOnlyReplica in $ReadOnlyReplicas)
			{
				$tmp += "$($ReadOnlyReplica)`r"
			}
			$Table.Cell($xRow,2).Range.Text = $tmp
		}
		$ReadOnlyReplicas = $Null
		$tmp = $Null

		Write-Verbose "$(Get-Date): `t`tGetting Replica directory servers"
		$xRow++
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Replica directory servers"
		$Replicas = $DomainInfo.ReplicaDirectoryServers | Sort
		If($Replicas -eq $Null)
		{
			$Table.Cell($xRow,2).Range.Text = "<None>"
		}
		Else
		{
			$tmp = @()
			ForEach($Replica in $Replicas)
			{
				$tmp += "$($Replica)`r"
			}
			$Table.Cell($xRow,2).Range.Text = $tmp
		}
		$Replicas = $Null
		$tmp = $Null

		Write-Verbose "$(Get-Date): `t`tGetting Subordinate references"
		$xRow++
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Subordinate references"
		$SubordinateReferences = $DomainInfo.SubordinateReferences | Sort
		If($SubordinateReferences -eq $Null)
		{
			$Table.Cell($xRow,2).Range.Text = "<None>"
		}
		Else
		{
			$tmp = @()
			ForEach($SubordinateReference in $SubordinateReferences)
			{
				$tmp += "$($SubordinateReference)`r"
			}
			$Table.Cell($xRow,2).Range.Text = $tmp
		}
		$SubordinateReferences = $Null
		$tmp = $Null
		
		#set column widths
		$xcols = $table.columns

		ForEach($xcol in $xcols)
		{
		    switch ($xcol.Index)
		    {
			  1 {$xcol.width = 175}
			  2 {$xcol.width = 300}
		    }
		}
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
		$Table.AutoFitBehavior($wdAutoFitFixed)	

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null

		Write-Verbose "$(Get-Date): `t`tGetting domain trusts"
		WriteWordLine 3 0 "Domain trusts"
		
		$ADDomainTrusts = $Null
		$ADDomainTrusts = Get-ADObject -Filter {ObjectClass -eq "trustedDomain"} -Server $Domain -Properties *

		If($? -and $ADDomainTrusts -ne $Null)
		{
			
			ForEach($Trust in $ADDomainTrusts) 
			{ 
				$TableRange = $doc.Application.Selection.Range
				[int]$Columns = 2
				If([String]::IsNullOrEmpty($Trust.Description))
				{
					[int]$Rows = 6
				}
				Else
				{
					[int]$Rows = 7
				}
				[int]$xRow = 0
				$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				$Table.Style = $myHash.Word_TableGrid
				$Table.Borders.InsideLineStyle = $wdLineStyleSingle
				$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
				$xRow++
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Name"
				$Table.Cell($xRow,2).Range.Text = $Trust.Name 
				
				If(![String]::IsNullOrEmpty($Trust.Description))
				{
					$xRow++
					$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Description"
					$Table.Cell($xRow,2).Range.Text = $Trust.Description
				}
				
				$xRow++
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Created"
				$Table.Cell($xRow,2).Range.Text = $Trust.Created
				
				$xRow++
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Modified"
				$Table.Cell($xRow,2).Range.Text = $Trust.Modified

				$TrustDirectionNumber = $Trust.TrustDirection
				$TrustTypeNumber = $Trust.TrustType
				$TrustAttributesNumber = $Trust.TrustAttributes

				#http://msdn.microsoft.com/en-us/library/cc234293.aspx
				Switch ($TrustTypeNumber) 
				{ 
					1 { $TrustType = "Trust with a Windows domain not running Active Directory"} 
					2 { $TrustType = "Trust with a Windows domain running Active Directory"} 
					3 { $TrustType = "Trust with a non-Windows-compliant Kerberos distribution"} 
					4 { $TrustType = "Trust with a DCE realm (not used)"} 
					Default { $TrustType = "Invalid Trust Type of $($TrustTypeNumber)" }
				} 
				$xRow++
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Type"
				If($TrustTypeNumber -lt 1 -or $TrustTypeNumber -gt 4)
				{
					$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorRed
					$Table.Cell($xRow,2).Range.Font.Bold  = $True
					$Table.Cell($xRow,2).Range.Font.Color = $WDColorBlack
				}
				$Table.Cell($xRow,2).Range.Text = $TrustType

				#http://msdn.microsoft.com/en-us/library/cc223779.aspx
				#thanks to fellow CTP Jeremy Saunders for the following switch stmt for trustAttributes
				#I adapted his code
				$attributes = @()
				$hextrustAttributesValue = '{0:X}' -f $trustAttributesNumber
				Switch ($hextrustAttributesValue)
				{
					{($hextrustAttributesValue -bor 0x00000001) -eq $hextrustAttributesValue} 
						{$attributes += "Non-Transitive"}
					
					{($hextrustAttributesValue -bor 0x00000002) -eq $hextrustAttributesValue} 
						{$attributes += "Uplevel clients only"}
					
					{($hextrustAttributesValue -bor 0x00000004) -eq $hextrustAttributesValue} 
						{$attributes += "Quarantined Domain (External, SID Filtering)"}
					
					{($hextrustAttributesValue -bor 0x00000008) -eq $hextrustAttributesValue} 
						{$attributes += "Cross-Organizational Trust (Selective Authentication)"}
					
					{($hextrustAttributesValue -bor 0x00000010) -eq $hextrustAttributesValue} 
						{$attributes += "Intra-Forest Trust"}
					
					{($hextrustAttributesValue -bor 0x00000020) -eq $hextrustAttributesValue} 
						{$attributes += "Inter-Forest Trust"}
					
					{($hextrustAttributesValue -bor 0x00000040) -eq $hextrustAttributesValue} 
						{$attributes += "MIT Trust using RC4 Encryption"}
					
					{($hextrustAttributesValue -bor 0x00000200) -eq $hextrustAttributesValue} 
						{$attributes += "Cross organization Trust no TGT delegation"}
				}

				$xRow++
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Attributes"
				$tmp = @()
				ForEach($attribute in $attributes)
				{
					$tmp += "$($attribute)`r"
				}
				$Table.Cell($xRow,2).Range.Text = $tmp

				#http://msdn.microsoft.com/en-us/library/cc223768.aspx
				Switch ($TrustDirectionNumber) 
				{ 
					0 { $TrustDirection = "Disabled"} 
					1 { $TrustDirection = "Inbound"} 
					2 { $TrustDirection = "Outbound"} 
					3 { $TrustDirection = "Bidirectional"} 
					Default { $TrustDirection = $TrustDirectionNumber }
				}
				$xRow++
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Direction"
				$Table.Cell($xRow,2).Range.Text = $TrustDirection
				
				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
				$Table.AutoFitBehavior($wdAutoFitContent)

				#return focus back to document
				$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

				#move to the end of the current document
				$selection.EndKey($wdStory,$wdMove) | Out-Null
				$TableRange = $Null
				$Table = $Null
				WriteWordLine 0 0 ""
			}
		}
		ElseIf(!$?)
		{
			#error retrieving domain trusts
			Write-Warning "Error retrieving domain trusts for $($Domain)"
			WriteWordLine 0 0 "Error retrieving domain trusts for $($Domain)" "" $Null 0 $False $True
		}
		Else
		{
			#no domain trust data
			WriteWordLine 0 0 "<None>"
		}

		Write-Verbose "$(Get-Date): `t`tProcessing domain controllers"
		$DomainControllers = $Null
		$DomainControllers = Get-ADDomainController -Filter * -Server $DomainInfo.DNSRoot | Sort Name
		
		If($? -and $DomainControllers -ne $Null)
		{
			$AllDomainControllers += $DomainControllers
			WriteWordLine 3 0 "Domain Controllers"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 1
			If($DomainControllers -is [array])
			{
				[int]$Rows = $DomainControllers.Count
			}
			Else
			{
				[int]$Rows = 1
			}
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$Table.AutoFitBehavior($wdAutoFitFixed)
			$Table.Style = $myHash.Word_TableGrid
			$Table.Borders.InsideLineStyle = $wdLineStyleSingle
			$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
			[int]$xRow = 0
			ForEach($DomainController in $DomainControllers)
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = $DomainController.Name
			}
			#set column widths
			$xcols = $table.columns

			ForEach($xcol in $xcols)
			{
			    switch ($xcol.Index)
			    {
				  1 {$xcol.width = 100}
			    }
			}

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
			$Table.AutoFitBehavior($wdAutoFitFixed)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		ElseIf(!$?)
		{
			Write-Warning "Error retrieving domain controller data for domain $($Domain)"
			WriteWordLine 0 0 "Error retrieving domain controller data for domain $($Domain)" "" $Null 0 $False $True
		}
		Else
		{
			WriteWordLine 0 0 "No Domain controller data was retrieved for domain $($Domain)" "" $Null 0 $False $True
		}
		
		$DomainControllers = $Null
		$LinkedGPOs = $Null
		$SubordinateReferences = $Null
		$Replicas = $Null
		$ReadOnlyReplicas = $Null
		$ChildDomains = $Null
		$DNSSuffixes = $Null
		$First = $False
	}
	ElseIf(!$?)
	{
		Write-Warning "Error retrieving domain data for domain $($Domain)."
		WriteWordLine 0 0 "Error retrieving domain data for domain $($Domain)" "" $Null 0 $False $True
	}
	Else
	{
		WriteWordLine 0 0 "No Domain data was retrieved for domain $($Domain)" "" $Null 0 $False $True
	}
}
$DomainControllers = $Null
$LinkedGPOs = $Null
$SubordinateReferences = $Null
$Replicas = $Null
$ReadOnlyReplicas = $Null
$ChildDomains = $Null
$DNSSuffixes = $Null
$First = $False
$SchemaVersionTable = $Null
$DomainInfo = $Null
$ADSchemaInfo = $Null
$ExchangeSchemaInfo = $Null
$ADDomainTrusts = $Null
$attributes = $Null
[gc]::collect() 

#domain controllers
Write-Verbose "$(Get-Date): Writing domain controller data"

$selection.InsertNewPage()
WriteWordLine 1 0 "Domain Controllers in $($ForestName)"
$AllDomainControllers = $AllDomainControllers | Sort Name
$First = $True

ForEach($DC in $AllDomainControllers)
{
	Write-Verbose "$(Get-Date): `tProcessing domain controller $($DC.name)"
	
	If(!$First)
	{
		#put each DC, starting with the second, on a new page
		$selection.InsertNewPage()
	}
	
	WriteWordLine 2 0 $DC.Name
	$TableRange = $doc.Application.Selection.Range
	[int]$Columns = 2
	If(!$Hardware)
	{
		[int]$Rows = 16
	}
	Else
	{
		[int]$Rows = 14
	}
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$Table.AutoFitBehavior($wdAutoFitFixed)
	$Table.Style = $myHash.Word_TableGrid
	$Table.Borders.InsideLineStyle = $wdLineStyleSingle
	$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
	$xRow = 1
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "Default partition"
	$Table.Cell($xRow,2).Range.Text = $DC.DefaultPartition
	
	$xRow++
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "Domain"
	$Table.Cell($xRow,2).Range.Text = $DC.domain
	
	$xRow++
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "Enabled"
	If($DC.Enabled -eq $True)
	{
		$Table.Cell($xRow,2).Range.Text = "True"
	}
	Else
	{
		$Table.Cell($xRow,2).Range.Text = "False"
	}
	
	$xRow++
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "Hostname"
	$Table.Cell($xRow,2).Range.Text = $DC.HostName
	
	$xRow++
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "Global Catalog"
	If($DC.IsGlobalCatalog -eq $True)
	{
		$Table.Cell($xRow,2).Range.Text = "Yes" 
	}
	Else
	{
		$Table.Cell($xRow,2).Range.Text = "No"
	}
	
	$xRow++
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "Read-only"
	If($DC.IsReadOnly -eq $True)
	{
		$Table.Cell($xRow,2).Range.Text = "Yes"
	}
	Else
	{
		$Table.Cell($xRow,2).Range.Text = "No"
	}
	
	$xRow++
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "LDAP port"
	$Table.Cell($xRow,2).Range.Text = $DC.LdapPort
	
	$xRow++
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "SSL port"
	$Table.Cell($xRow,2).Range.Text = $DC.SslPort
	
	$xRow++
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "Operation Master roles"
	$FSMORoles = $DC.OperationMasterRoles | Sort
	If($FSMORoles -eq $Null)
	{
		$Table.Cell($xRow,2).Range.Text = "<None>"
	}
	Else
	{
		$tmp = ""
		ForEach($FSMORole in $FSMORoles)
		{
			$tmp += ($FSMORole.ToString() + "`n")
		}
		$Table.Cell($xRow,2).Range.Text = $tmp
	}
	
	$xRow++
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "Partitions"
	$Partitions = $DC.Partitions | Sort
	If($Partitions -eq $Null)
	{
		$Table.Cell($xRow,2).Range.Text = "<None>"
	}
	Else
	{
		$tmp = ""
		ForEach($Partition in $Partitions)
		{
			$tmp += ($Partition + "`n")
		}
		$Table.Cell($xRow,2).Range.Text = $tmp
	}
	
	$xRow++
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "Site"
	$Table.Cell($xRow,2).Range.Text = $DC.Site

	$xRow++
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "Operating System"
	$Table.Cell($xRow,2).Range.Text = $DC.OperatingSystem
	
	$xRow++
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "Service Pack"
	$Table.Cell($xRow,2).Range.Text = $DC.OperatingSystemServicePack
	
	$xRow++
	$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell($xRow,1).Range.Font.Bold = $True
	$Table.Cell($xRow,1).Range.Text = "Operating System version"
	$Table.Cell($xRow,2).Range.Text = $DC.OperatingSystemVersion
	
	If(!$Hardware)
	{
		
		$xRow++
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "IPv4 Address"
		If([String]::IsNullOrEmpty($DC.IPv4Address))
		{
			$Table.Cell($xRow,2).Range.Text = "<None>"
		}
		Else
		{
			$Table.Cell($xRow,2).Range.Text = $DC.IPv4Address
		}
		
		$xRow++
		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "IPv6 Address"
		If([String]::IsNullOrEmpty($DC.IPv6Address))
		{
			$Table.Cell($xRow,2).Range.Text = "<None>"
		}
		Else
		{
			$Table.Cell($xRow,2).Range.Text = $DC.IPv6Address
		}
	}
	
	#set column widths
	$xcols = $table.columns

	ForEach($xcol in $xcols)
	{
	    switch ($xcol.Index)
	    {
		  1 {$xcol.width = 140}
		  2 {$xcol.width = 300}
	    }
	}

	$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
	$Table.AutoFitBehavior($wdAutoFitFixed)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
	$TableRange = $Null
	$Table = $Null
	
	If($Hardware -or $Services)
	{
		If(Test-Connection -ComputerName $DC.name -quiet -EA 0)
		{
			If($Hardware)
			{
				GetComputerWMIInfo $DC.Name
			}
			
			If($Services)
			{
				GetComputerServices $DC.Name
			}
		}
		Else
		{
			Write-Verbose "$(Get-Date): `t`t$($DC.Name) is offline or unreachable.  Hardware inventory is skipped."
			WriteWordLine 0 0 "Server $($DC.Name) was offline or unreachable at "(get-date).ToString()
			If($Hardware -and -not $Services)
			{
				WriteWordLine 0 0 "Hardware inventory was skipped."
			}
			ElseIf($Services -and -not $Hardware)
			{
				WriteWordLine 0 0 "Services was skipped."
			}
			ElseIf($Hardware -and $Services)
			{
				WriteWordLine 0 0 "Hardware inventory and Services were skipped."
			}
		}
	}
	$First = $False
}
$AllDomainControllers = $Null
[gc]::collect() 

#organizational units
Write-Verbose "$(Get-Date): Writing OU data by Domain"
$selection.InsertNewPage()
WriteWordLine 1 0 "Organizational Units"
$First = $True

ForEach($Domain in $Domains)
{
	Write-Verbose "$(Get-Date): `tProcessing domain $($Domain)"
	If(!$First)
	{
		#put each domain, starting with the second, on a new page
		$selection.InsertNewPage()
	}
	If($Domain -eq $ForestRootDomain)
	{
		WriteWordLine 2 0 "OUs in Domain $($Domain) (Forest Root)"
	}
	Else
	{
		WriteWordLine 2 0 "OUs in Domain $($Domain)"
	}
	#get all OUs for the domain
	$OUs = $Null
	$OUs = Get-ADOrganizationalUnit -Filter * -Server $Domain `
	-Properties CanonicalName, DistinguishedName, Name, Created, ProtectedFromAccidentalDeletion | `
	Select CanonicalName, DistinguishedName, Name, Created, ProtectedFromAccidentalDeletion | `
	Sort CanonicalName
	
	If($? -and $OUs -ne $Null)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 6
		If($OUs -is [array])
		{
			[int]$Rows = $OUs.Count + 1
			[int]$NumOUs = $OUs.Count
		}
		Else
		{
			[int]$Rows = 2
			[int]$NumOUs = 1
		}
		[int]$xRow = 1
		[int]$OUCount = 0

		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.AutoFitBehavior($wdAutoFitFixed)
		$Table.Style = $myHash.Word_TableGrid
		$Table.rows.first.headingformat = $wdHeadingFormatTrue
		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

		$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Name"
		
		$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,2).Range.Font.Bold = $True
		$Table.Cell($xRow,2).Range.Text = "Created"
		
		$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,3).Range.Font.Bold = $True
		$Table.Cell($xRow,3).Range.Text = "Protected"
		
		$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,4).Range.Font.Bold = $True
		$Table.Cell($xRow,4).Range.Text = "# Users"
		
		$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,5).Range.Font.Bold = $True
		$Table.Cell($xRow,5).Range.Text = "# Computers"
		
		$Table.Cell($xRow,6).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,6).Range.Font.Bold = $True
		$Table.Cell($xRow,6).Range.Text = "# Groups"

		ForEach($OU in $OUs)
		{
			$xRow++
			$OUCount++
			If($xRow % 2 -eq 0)
			{
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
				$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
				$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray05
				$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray05
				$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorGray05
				$Table.Cell($xRow,6).Shading.BackgroundPatternColor = $wdColorGray05
			}
			$OUDisplayName = $OU.CanonicalName.SubString($OU.CanonicalName.IndexOf("/")+1)
			Write-Verbose "$(Get-Date): `t`tProcessing OU $($OU.CanonicalName) - OU # $OUCount of $NumOUs"
			
			#get counts of users, computers and groups in the OU
			Write-Verbose "$(Get-Date): `t`t`tGetting user count"
			
			[int]$UserCount = 0
			[int]$ComputerCount = 0
			[int]$GroupCount = 0
			
			$Results = Get-ADUser -Filter * -SearchBase $OU.DistinguishedName -Server $Domain
			If($Results -eq $Null)
			{
				$UserCount = 0
			}
			ElseIf($Results -is [array])
			{
				$UserCount = $Results.Count
			}
			Else
			{
				$UserCount = 1
			}
			Write-Verbose "$(Get-Date): `t`t`tGetting computer count"
			$Results = Get-ADComputer -Filter * -SearchBase $OU.DistinguishedName -Server $Domain
			If($Results -eq $Null)
			{
				$ComputerCount = 0
			}
			ElseIf($Results -is [array])
			{
				$ComputerCount = $Results.Count
			}
			Else
			{
				$ComputerCount = 1
			}
			Write-Verbose "$(Get-Date): `t`t`tGetting group count"
			$Results = Get-ADGroup -Filter * -SearchBase $OU.DistinguishedName -Server $Domain
			If($Results -eq $Null)
			{
				$GroupCount = 0
			}
			ElseIf($Results -is [array])
			{
				$GroupCount = $Results.Count
			}
			Else
			{
				$GroupCount = 1
			}
			
			Write-Verbose "$(Get-Date): `t`t`tPopulating table row"
			$Table.Cell($xRow,1).Range.Text = $OUDisplayName
			$Table.Cell($xRow,2).Range.Text = $OU.Created
			If($OU.ProtectedFromAccidentalDeletion -eq $True)
			{
				$Table.Cell($xRow,3).Range.Text = "Yes"
			}
			Else
			{
				$Table.Cell($xRow,3).Range.Text = "No"
			}
			
			[string]$UserCountStr = "{0,7:N0}" -f $UserCount
			[string]$ComputerCountStr = "{0,7:N0}" -f $ComputerCount
			[string]$GroupCountStr = "{0,7:N0}" -f $GroupCount

			$Table.Cell($xRow,4).Range.ParagraphFormat.Alignment = $wdCellAlignVerticalTop
			$Table.Cell($xRow,4).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
			$Table.Cell($xRow,4).Range.Text = $UserCountStr
			$Table.Cell($xRow,5).Range.ParagraphFormat.Alignment = $wdCellAlignVerticalTop
			$Table.Cell($xRow,5).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
			$Table.Cell($xRow,5).Range.Text = $ComputerCountStr
			$Table.Cell($xRow,6).Range.ParagraphFormat.Alignment = $wdCellAlignVerticalTop
			$Table.Cell($xRow,6).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
			$Table.Cell($xRow,6).Range.Text = $GroupCountStr
			$Results = $Null
			$UserCountStr = $Null
			$ComputerCountStr = $Null
			$GroupCountStr = $Null
		}
		
		#set column widths
		$xcols = $table.columns

		ForEach($xcol in $xcols)
		{
		    switch ($xcol.Index)
		    {
			  1 {$xcol.width = 214}
			  2 {$xcol.width = 68}
			  3 {$xcol.width = 56}
			  4 {$xcol.width = 56}
			  5 {$xcol.width = 70}
			  6 {$xcol.width = 56}
		    }
		}
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
		$Table.AutoFitBehavior($wdAutoFitFixed)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null
		$Results = $Null
		$UserCountStr = $Null
		$ComputerCountStr = $Null
		$GroupCountStr = $Null
	}
	ElseIf(!$?)
	{
		Write-Warning "Error retrieving OU data for domain $($Domain)"
		WriteWordLine 0 0 "Error retrieving OU data for domain $($Domain)" "" $Null 0 $False $True
	}
	Else
	{
		WriteWordLine 0 0 "No OU data was retrieved for domain $($Domain)" "" $Null 0 $False $True
	}
	$First = $False
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
}
$OUs = $Null
$OUDisplayName = $Null
$Results = $Null
$UserCountStr = $Null
$ComputerCountStr = $Null
$GroupCountStr = $Null
[gc]::collect() 

#Group information
Write-Verbose "$(Get-Date): Writing group data"

$selection.InsertNewPage()
WriteWordLine 1 0 "Groups"
$First = $True

ForEach($Domain in $Domains)
{
	Write-Verbose "$(Get-Date): `tProcessing groups in domain $($Domain)"
	If(!$First)
	{
		#put each domain, starting with the second, on a new page
		$selection.InsertNewPage()
	}
	If($Domain -eq $ForestRootDomain)
	{
		WriteWordLine 2 0 "Domain $($Domain) (Forest Root)"
	}
	Else
	{
		WriteWordLine 2 0 "Domain $($Domain)"
	}

	#get all Groups for the domain
	$Groups = $Null
	$Groups = Get-ADGroup -Filter * -Server $Domain -Properties Name, GroupCategory, GroupType | Sort Name

	If($? -and $Groups -ne $Null)
	{
		#get counts
		
		Write-Verbose "$(Get-Date): `t`tGetting counts"
		
		[int]$SecurityCount = 0
		[int]$DistributionCount = 0
		[int]$GlobalCount = 0
		[int]$UniversalCount = 0
		[int]$DomainLocalCount = 0
		[int]$ContactsCount = 0
		[int]$GroupsWithSIDHistory = 0
		
		Write-Verbose "$(Get-Date): `t`t`tSecurity Groups"
		$Results = $groups | Where {$_.groupcategory -eq "Security"}
		
		If($Results -eq $Null)
		{
			[int]$SecurityCount = 0
		}
		ElseIf($Results -is [array])
		{
			[int]$SecurityCount = $Results.Count
		}
		Else
		{
			[int]$SecurityCount = 1
		}
		
		Write-Verbose "$(Get-Date): `t`t`tDistribution Groups"
		$Results = $groups | Where {$_.groupcategory -eq "Distribution"}
		
		If($Results -eq $Null)
		{
			[int]$DistributionCount = 0
		}
		ElseIf($Results -is [array])
		{
			[int]$DistributionCount = $Results.Count
		}
		Else
		{
			[int]$DistributionCount = 1
		}

		Write-Verbose "$(Get-Date): `t`t`tGlobal Groups"
		$Results = $groups | Where {$_.groupscope -eq "Global"}

		If($Results -eq $Null)
		{
			[int]$GlobalCount = 0
		}
		ElseIf($Results -is [array])
		{
			[int]$GlobalCount = $Results.Count
		}
		Else
		{
			[int]$GlobalCount = 1
		}

		Write-Verbose "$(Get-Date): `t`t`tUniversal Groups"
		$Results = $groups | Where {$_.groupscope -eq "Universal"}

		If($Results -eq $Null)
		{
			[int]$UniversalCount = 0
		}
		ElseIf($Results -is [array])
		{
			[int]$UniversalCount = $Results.Count
		}
		Else
		{
			[int]$UniversalCount = 1
		}
		
		Write-Verbose "$(Get-Date): `t`t`tDomain Local Groups"
		$Results = $groups | Where {$_.groupscope -eq "DomainLocal"}

		If($Results -eq $Null)
		{
			[int]$DomainLocalCount = 0
		}
		ElseIf($Results -is [array])
		{
			[int]$DomainLocalCount = $Results.Count
		}
		Else
		{
			[int]$DomainLocalCount = 1
		}

		Write-Verbose "$(Get-Date): `t`t`tGroups with SID History"
		$Results = $Null
		$Results = Get-ADObject -LDAPFilter "(sIDHistory=*)" -Server $Domain -Property objectClass, sIDHistory

		If($Results -eq $Null)
		{
			[int]$GroupsWithSIDHistory = 0
		}
		ElseIf($Results -is [array])
		{
			[int]$GroupsWithSIDHistory = ($Results | Where {$_.objectClass -eq 'group'}).Count
		}
		Else
		{
			[int]$GroupsWithSIDHistory = 1
		}

		Write-Verbose "$(Get-Date): `t`t`tContacts"
		$Results = $Null
		$Results = Get-ADObject -LDAPFilter "objectClass=Contact" -Server $Domain

		If($Results -eq $Null)
		{
			[int]$ContactsCount = 0
		}
		ElseIf($Results -is [array])
		{
			[int]$ContactsCount = $Results.Count
		}
		Else
		{
			[int]$ContactsCount = 1
		}

		[string]$TotalCountStr = "{0,7:N0}" -f ($SecurityCount + $DistributionCount)
		[string]$SecurityCountStr = "{0,7:N0}" -f $SecurityCount
		[string]$DomainLocalCountStr = "{0,7:N0}" -f $DomainLocalCount
		[string]$GlobalCountStr = "{0,7:N0}" -f $GlobalCount
		[string]$UniversalCountStr = "{0,7:N0}" -f $UniversalCount
		[string]$DistributionCountStr = "{0,7:N0}" -f $DistributionCount
		[string]$GroupsWithSIDHistoryStr = "{0,7:N0}" -f $GroupsWithSIDHistory
		[string]$ContactsCountStr = "{0,7:N0}" -f $ContactsCount
		
		Write-Verbose "$(Get-Date): `t`tBuild groups table"
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		[int]$Rows = 8
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.Style = $myHash.Word_TableGrid
		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "Total Groups"
		$Table.Cell(1,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(1,2).Range.Text = $TotalCountStr
		$Table.Cell(2,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(2,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(2,1).Range.Font.Bold = $True
		$Table.Cell(2,1).Range.Text = "`tSecurity Groups"
		$Table.Cell(2,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(2,2).Range.Text = $SecurityCountStr
		$Table.Cell(3,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(3,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(3,1).Range.Font.Bold = $True
		$Table.Cell(3,1).Range.Text = "`t`tDomain Local"
		$Table.Cell(3,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(3,2).Range.Text = $DomainLocalCountStr
		$Table.Cell(4,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(4,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(4,1).Range.Font.Bold = $True
		$Table.Cell(4,1).Range.Text = "`t`tGlobal"
		$Table.Cell(4,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(4,2).Range.Text = $GlobalCountStr
		$Table.Cell(5,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(5,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(5,1).Range.Font.Bold = $True
		$Table.Cell(5,1).Range.Text = "`t`tUniversal"
		$Table.Cell(5,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(5,2).Range.Text = $UniversalCountStr
		$Table.Cell(6,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(6,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(6,1).Range.Font.Bold = $True
		$Table.Cell(6,1).Range.Text = "`tDistribution Groups"
		$Table.Cell(6,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(6,2).Range.Text = $DistributionCountStr
		$Table.Cell(7,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(7,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(7,1).Range.Font.Bold = $True
		$Table.Cell(7,1).Range.Text = "Groups with SID History"
		$Table.Cell(7,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(7,2).Range.Text = $GroupsWithSIDHistoryStr
		$Table.Cell(8,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(8,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(8,1).Range.Font.Bold = $True
		$Table.Cell(8,1).Range.Text = "Contacts"
		$Table.Cell(8,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(8,2).Range.Text = $ContactsCountStr

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
		$Table.AutoFitBehavior($wdAutoFitContent)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null
		
		#get members of privileged groups
		$DomainInfo = $Null
		$DomainInfo = Get-ADDomain -Identity $Domain
		
		If($? -and $DomainInfo -ne $Null)
		{
			$DomainAdminsSID = "$($DomainInfo.DomainSID)-512"
			$EnterpriseAdminsSID = "$($DomainInfo.DomainSID)-519"
			$SchemaAdminsSID = "$($DomainInfo.DomainSID)-518"
		}
		Else
		{
			$DomainAdminsSID = $Null
			$EnterpriseAdminsSID = $Null
			$SchemaAdminsSID = $Null
		}
		
		WriteWordLine 3 0 "Privileged Groups"
		Write-Verbose "$(Get-Date): `t`tListing domain admins"
		$Admins = $Null
		$Admins = Get-ADGroupMember -Identity $DomainAdminsSID -Server $Domain
		
		If($? -and $Admins -ne $Null)
		{
			If($Admins -is [array])
			{
				[int]$AdminsCount = $Admins.Count
			}
			Else
			{
				[int]$AdminsCount = 1
			}
			$Admins = $Admins | Sort Name
			[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
			
			WriteWordLine 4 0 "Domain Admins ($($AdminsCountStr) members):"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 4
			[int]$Rows = $AdminsCount + 1
			[int]$xRow = 1
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$Table.AutoFitBehavior($wdAutoFitFixed)
			$Table.Style = $myHash.Word_TableGrid
			$Table.rows.first.headingformat = $wdHeadingFormatTrue
			$Table.Borders.InsideLineStyle = $wdLineStyleSingle
			$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Name"
			$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,2).Range.Font.Bold = $True
			$Table.Cell($xRow,2).Range.Text = "Password Last Changed"
			$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,3).Range.Font.Bold = $True
			$Table.Cell($xRow,3).Range.Text = "Password Never Expires"
			$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,4).Range.Font.Bold = $True
			$Table.Cell($xRow,4).Range.Text = "Account Enabled"
			#$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorGray15
			#$Table.Cell($xRow,5).Range.Font.Bold = $True
			#$Table.Cell($xRow,5).Range.Text = "Password Policy"
			ForEach($Admin in $Admins)
			{
				$User = Get-ADUser -Identity $Admin.SID -Server $Domain -Properties PasswordLastSet, Enabled, PasswordNeverExpires 

				$xRow++
				
				If($? -and $User -ne $Null)
				{
					$Table.Cell($xRow,1).Range.Text = $User.Name
					If($User.PasswordLastSet -eq $Null)
					{
						$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorRed
						$Table.Cell($xRow,2).Range.Font.Bold  = $True
						$Table.Cell($xRow,2).Range.Font.Color = $WDColorBlack
						$Table.Cell($xRow,2).Range.Text = "No Date Set"
					}
					Else
					{
						$Table.Cell($xRow,2).Range.Text = (get-date $User.PasswordLastSet -f d)
					}
					If($User.PasswordNeverExpires -eq $True)
					{
						$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorRed
						$Table.Cell($xRow,3).Range.Font.Bold  = $True
						$Table.Cell($xRow,3).Range.Font.Color = $WDColorBlack
					}
					$Table.Cell($xRow,3).Range.Text = $User.PasswordNeverExpires
					If($User.Enabled -eq $False)
					{
						$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorRed
						$Table.Cell($xRow,4).Range.Font.Bold  = $True
						$Table.Cell($xRow,4).Range.Font.Color = $WDColorBlack
					}
					$Table.Cell($xRow,4).Range.Text = $User.Enabled
					#$Table.Cell($xRow,5).Range.Text = ""
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $Admin.SID
					$Table.Cell($xRow,3).Range.Text = "Unknown"
					$Table.Cell($xRow,4).Range.Text = "Unknown"
					#$Table.Cell($xRow,5).Range.Text = "Unknown"
				}
			}
			
			#set column widths
			$xcols = $table.columns

			ForEach($xcol in $xcols)
			{
			    switch ($xcol.Index)
			    {
				  1 {$xcol.width = 200}
				  2 {$xcol.width = 66}
				  3 {$xcol.width = 56}
				  4 {$xcol.width = 50}
				  5 {$xcol.width = 142}
			    }
			}
			
			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
			$Table.AutoFitBehavior($wdAutoFitFixed)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		ElseIf(!$?)
		{
			WriteWordLine 0 0 "Unable to retrieve Domain Admins group membership" "" $Null 0 $False $True
		}
		Else
		{
			WriteWordLine 4 0 "Domain Admins: "
			WriteWordLine 0 0 "<None>"
		}

		Write-Verbose "$(Get-Date): `t`tListing enterprise admins"
		
		If($Domain -eq $ForestRootDomain)
		{
			$Admins = Get-ADGroupMember -Identity $EnterpriseAdminsSID -Server $Domain 
			
			If($? -and $Admins -ne $Null)
			{
				If($Admins -is [array])
				{
					[int]$AdminsCount = $Admins.Count
				}
				Else
				{
					[int]$AdminsCount = 1
				}
				$Admins = $Admins | Sort Name
				[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
				
				WriteWordLine 4 0 "Enterprise Admins ($($AdminsCountStr) members):"
				$TableRange = $doc.Application.Selection.Range
				[int]$Columns = 4
				[int]$Rows = $AdminsCount + 1
				[int]$xRow = 1
				$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				$Table.AutoFitBehavior($wdAutoFitFixed)
				$Table.Style = $myHash.Word_TableGrid
				$Table.rows.first.headingformat = $wdHeadingFormatTrue
				$Table.Borders.InsideLineStyle = $wdLineStyleSingle
				$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Name"
				$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,2).Range.Font.Bold = $True
				$Table.Cell($xRow,2).Range.Text = "Password Last Changed"
				$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,3).Range.Font.Bold = $True
				$Table.Cell($xRow,3).Range.Text = "Password Never Expires"
				$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,4).Range.Font.Bold = $True
				$Table.Cell($xRow,4).Range.Text = "Account Enabled"
				#$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorGray15
				#$Table.Cell($xRow,5).Range.Font.Bold = $True
				#$Table.Cell($xRow,5).Range.Text = "Password Policy"
				ForEach($Admin in $Admins)
				{
					$User = Get-ADUser -Identity $Admin.SID -Server $Domain -Properties PasswordLastSet, Enabled, PasswordNeverExpires 

					$xRow++
					
					If($? -and $User -ne $Null)
					{
						$Table.Cell($xRow,1).Range.Text = $User.Name
						If($User.PasswordLastSet -eq $Null)
						{
							$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorRed
							$Table.Cell($xRow,2).Range.Font.Bold  = $True
							$Table.Cell($xRow,2).Range.Font.Color = $WDColorBlack
							$Table.Cell($xRow,2).Range.Text = "No Date Set"
						}
						Else
						{
							$Table.Cell($xRow,2).Range.Text = (get-date $User.PasswordLastSet -f d)
						}
						If($User.PasswordNeverExpires -eq $True)
						{
							$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorRed
							$Table.Cell($xRow,3).Range.Font.Bold  = $True
							$Table.Cell($xRow,3).Range.Font.Color = $WDColorBlack
						}
						$Table.Cell($xRow,3).Range.Text = $User.PasswordNeverExpires
						If($User.Enabled -eq $False)
						{
							$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorRed
							$Table.Cell($xRow,4).Range.Font.Bold  = $True
							$Table.Cell($xRow,4).Range.Font.Color = $WDColorBlack
						}
						$Table.Cell($xRow,4).Range.Text = $User.Enabled
						#$Table.Cell($xRow,5).Range.Text = ""
					}
					Else
					{
						$Table.Cell($xRow,2).Range.Text = $Admin.SID
						$Table.Cell($xRow,3).Range.Text = "Unknown"
						$Table.Cell($xRow,4).Range.Text = "Unknown"
						#$Table.Cell($xRow,5).Range.Text = "Unknown"
					}
				}
			
				#set column widths
				$xcols = $table.columns

				ForEach($xcol in $xcols)
				{
				    switch ($xcol.Index)
				    {
					  1 {$xcol.width = 200}
					  2 {$xcol.width = 66}
					  3 {$xcol.width = 56}
					  4 {$xcol.width = 50}
					  5 {$xcol.width = 142}
				    }
				}

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
				$Table.AutoFitBehavior($wdAutoFitFixed)

				#return focus back to document
				$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

				#move to the end of the current document
				$selection.EndKey($wdStory,$wdMove) | Out-Null
				$TableRange = $Null
				$Table = $Null
			}
			ElseIf(!$?)
			{
				WriteWordLine 0 0 "Unable to retrieve Enterprise Admins group membership" "" $Null 0 $False $True
			}
			Else
			{
				WriteWordLine 0 0 "<None>"
			}
		}
		Else
		{
			WriteWordLine 4 0 "Enterprise Admins: "
			WriteWordLine 0 0 "<None>"
		}
		
		Write-Verbose "$(Get-Date): `t`tListing schema admins"
		
		If($Domain -eq $ForestRootDomain)
		{
			$Admins = Get-ADGroupMember -Identity $SchemaAdminsSID -Server $Domain 
			
			If($? -and $Admins -ne $Null)
			{
				If($Admins -is [array])
				{
					[int]$AdminsCount = $Admins.Count
				}
				Else
				{
					[int]$AdminsCount = 1
				}
				$Admins = $Admins | Sort Name
				[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
				
				WriteWordLine 4 0 "Schema Admins ($($AdminsCountStr) members):"
				$TableRange = $doc.Application.Selection.Range
				[int]$Columns = 4
				[int]$Rows = $AdminsCount + 1
				[int]$xRow = 1
				$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				$Table.AutoFitBehavior($wdAutoFitFixed)
				$Table.Style = $myHash.Word_TableGrid
				$Table.rows.first.headingformat = $wdHeadingFormatTrue
				$Table.Borders.InsideLineStyle = $wdLineStyleSingle
				$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
				
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Name"
				$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,2).Range.Font.Bold = $True
				$Table.Cell($xRow,2).Range.Text = "Password Last Changed"
				$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,3).Range.Font.Bold = $True
				$Table.Cell($xRow,3).Range.Text = "Password Never Expires"
				$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,4).Range.Font.Bold = $True
				$Table.Cell($xRow,4).Range.Text = "Account Enabled"
				#$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorGray15
				#$Table.Cell($xRow,5).Range.Font.Bold = $True
				#$Table.Cell($xRow,5).Range.Text = "Password Policy"
				ForEach($Admin in $Admins)
				{
					$User = Get-ADUser -Identity $Admin.SID -Server $Domain -Properties PasswordLastSet, Enabled, PasswordNeverExpires 

					$xRow++

					If($? -and $User -ne $Null)
					{
						$Table.Cell($xRow,1).Range.Text = $User.Name
						If($User.PasswordLastSet -eq $Null)
						{
							$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorRed
							$Table.Cell($xRow,2).Range.Font.Bold  = $True
							$Table.Cell($xRow,2).Range.Font.Color = $WDColorBlack
							$Table.Cell($xRow,2).Range.Text = "No Date Set"
						}
						Else
						{
							$Table.Cell($xRow,2).Range.Text = (get-date $User.PasswordLastSet -f d)
						}
						If($User.PasswordNeverExpires -eq $True)
						{
							$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorRed
							$Table.Cell($xRow,3).Range.Font.Bold  = $True
							$Table.Cell($xRow,3).Range.Font.Color = $WDColorBlack
						}
						$Table.Cell($xRow,3).Range.Text = $User.PasswordNeverExpires
						If($User.Enabled -eq $False)
						{
							$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorRed
							$Table.Cell($xRow,4).Range.Font.Bold  = $True
							$Table.Cell($xRow,4).Range.Font.Color = $WDColorBlack
						}
						$Table.Cell($xRow,4).Range.Text = $User.Enabled
						#$Table.Cell($xRow,5).Range.Text = ""
					}
					Else
					{
						$Table.Cell($xRow,2).Range.Text = $Admin.SID
						$Table.Cell($xRow,3).Range.Text = "Unknown"
						$Table.Cell($xRow,4).Range.Text = "Unknown"
						#$Table.Cell($xRow,5).Range.Text = "Unknown"
					}
				}
			
				#set column widths
				$xcols = $table.columns

				ForEach($xcol in $xcols)
				{
				    switch ($xcol.Index)
				    {
					  1 {$xcol.width = 200}
					  2 {$xcol.width = 66}
					  3 {$xcol.width = 56}
					  4 {$xcol.width = 50}
					  5 {$xcol.width = 142}
				    }
				}
				
				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
				$Table.AutoFitBehavior($wdAutoFitFixed)

				#return focus back to document
				$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

				#move to the end of the current document
				$selection.EndKey($wdStory,$wdMove) | Out-Null
				$TableRange = $Null
				$Table = $Null
			}
			ElseIf(!$?)
			{
				WriteWordLine 0 0 "Unable to retrieve Schema Admins group membership" "" $Null 0 $False $True
			}
			Else
			{
				WriteWordLine 0 0 "<None>"
			}
		}
		Else
		{
			WriteWordLine 4 0 "Schema Admins: "
			WriteWordLine 0 0 "<None>"
		}

		#http://www.shariqsheikh.com/blog/index.php/200908/use-powershell-to-look-up-admincount-from-adminsdholder-and-sdprop/		
		Write-Verbose "$(Get-Date): `t`tListing users with AdminCount = 1"
		$AdminCounts = Get-ADUser -LDAPFilter "(admincount=1)"  -Server $Domain 
		
		If($? -and $AdminCounts -ne $Null)
		{
			$AdminCounts = $AdminCounts | Sort Name
			If($AdminCounts -is [array])
			{
				[int]$AdminsCount = $AdminCounts.Count
			}
			Else
			{
				[int]$AdminsCount = 1
			}
			[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
			
			WriteWordLine 4 0 "Users with AdminCount=1 ($($AdminsCountStr) members):"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 4
			[int]$Rows = $AdminCounts.Count + 1
			[int]$xRow = 1
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$Table.AutoFitBehavior($wdAutoFitFixed)
			$Table.Style = $myHash.Word_TableGrid
			$Table.rows.first.headingformat = $wdHeadingFormatTrue
			$Table.Borders.InsideLineStyle = $wdLineStyleSingle
			$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
			
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Name"
			$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,2).Range.Font.Bold = $True
			$Table.Cell($xRow,2).Range.Text = "Password Last Changed"
			$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,3).Range.Font.Bold = $True
			$Table.Cell($xRow,3).Range.Text = "Password Never Expires"
			$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,4).Range.Font.Bold = $True
			$Table.Cell($xRow,4).Range.Text = "Account Enabled"
			#$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorGray15
			#$Table.Cell($xRow,5).Range.Font.Bold = $True
			#$Table.Cell($xRow,5).Range.Text = "Password Policy"
			ForEach($Admin in $AdminCounts)
			{
				$User = Get-ADUser -Identity $Admin.SID -Server $Domain -Properties PasswordLastSet, Enabled, PasswordNeverExpires 

				$xRow++
				
				If($? -and $User -ne $Null)
				{
					$Table.Cell($xRow,1).Range.Text = $User.Name
					If($User.PasswordLastSet -eq $Null)
					{
						$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorRed
						$Table.Cell($xRow,2).Range.Font.Bold  = $True
						$Table.Cell($xRow,2).Range.Font.Color = $WDColorBlack
						$Table.Cell($xRow,2).Range.Text = "No Date Set"
					}
					Else
					{
						$Table.Cell($xRow,2).Range.Text = (get-date $User.PasswordLastSet -f d)
					}
					If($User.PasswordNeverExpires -eq $True)
					{
						$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorRed
						$Table.Cell($xRow,3).Range.Font.Bold  = $True
						$Table.Cell($xRow,3).Range.Font.Color = $WDColorBlack
					}
					$Table.Cell($xRow,3).Range.Text = $User.PasswordNeverExpires
					If($User.Enabled -eq $False)
					{
						$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorRed
						$Table.Cell($xRow,4).Range.Font.Bold  = $True
						$Table.Cell($xRow,4).Range.Font.Color = $WDColorBlack
					}
					$Table.Cell($xRow,4).Range.Text = $User.Enabled
					#$Table.Cell($xRow,5).Range.Text = ""
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $Admin.SID
					$Table.Cell($xRow,3).Range.Text = "Unknown"
					$Table.Cell($xRow,4).Range.Text = "Unknown"
					#$Table.Cell($xRow,5).Range.Text = "Unknown"
				}
			}
			
			#set column widths
			$xcols = $table.columns

			ForEach($xcol in $xcols)
			{
			    switch ($xcol.Index)
			    {
				  1 {$xcol.width = 200}
				  2 {$xcol.width = 66}
				  3 {$xcol.width = 56}
				  4 {$xcol.width = 50}
				  5 {$xcol.width = 142}
			    }
			}

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
			$Table.AutoFitBehavior($wdAutoFitFixed)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		ElseIf(!$?)
		{
			WriteWordLine 0 0 "Unable to retrieve users with AdminCount=1" "" $Null 0 $False $True
		}
		Else
		{
			WriteWordLine 4 0 "Users with AdminCount=1: "
			WriteWordLIne 0 0 "<None>"
		}
		
		Write-Verbose "$(Get-Date): `t`tListing groups with AdminCount = 1"
		$AdminCounts = Get-ADGroup -LDAPFilter "(admincount=1)" -Server $Domain  | Select Name
		
		If($? -and $AdminCounts -ne $Null)
		{
			$AdminCounts = $AdminCounts | Sort Name
			If($AdminCounts -is [array])
			{
				[int]$AdminsCount = $AdminCounts.Count
			}
			Else
			{
				[int]$AdminsCount = 1
			}
			[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
			
			WriteWordLine 4 0 "Groups with AdminCount=1 ($($AdminsCountStr) members):"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			[int]$Rows = $AdminCounts.Count + 1
			[int]$xRow = 1
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$Table.AutoFitBehavior($wdAutoFitFixed)
			$Table.Style = $myHash.Word_TableGrid
			$Table.rows.first.headingformat = $wdHeadingFormatTrue
			$Table.Borders.InsideLineStyle = $wdLineStyleSingle
			$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Group Name"
			$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,2).Range.Font.Bold = $True
			$Table.Cell($xRow,2).Range.Text = "Members"
			ForEach($Admin in $AdminCounts)
			{
				Write-Verbose "$(Get-Date): `t`t`t$($Admin.Name)"
				$xRow++
				$Members = Get-ADGroupMember -Identity $Admin.Name -Server $Domain | Sort Name
				
				If($? -and $Members -ne $Null)
				{
					If($Members -is [array])
					{
						$MembersCount = $Members.Count
					}
					Else
					{
						$MembersCount = 1
					}
				}
				Else
				{
					$MembersCount = 0
				}

				[string]$MembersCountStr = "{0:N0}" -f $MembersCount
				$Table.Cell($xRow,1).Range.Text = "$($Admin.Name) ($($MembersCountStr) members)"
				$MbrStr = ""
				If($MembersCount -gt 0)
				{
					ForEach($Member in $Members)
					{
						$MbrStr += "$($Member.Name)`r"
					}
					$Table.Cell($xRow,2).Range.Text = $MbrStr
				}
			}
			
			#set column widths
			$xcols = $table.columns

			ForEach($xcol in $xcols)
			{
			    switch ($xcol.Index)
			    {
				  1 {$xcol.width = 200}
				  2 {$xcol.width = 172}
			    }
			}
			
			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
			$Table.AutoFitBehavior($wdAutoFitFixed)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		ElseIf(!$?)
		{
			WriteWordLine 0 0 "Unable to retrieve Groups with AdminCount=1" "" $Null 0 $False $True
		}
		Else
		{
			WriteWordLine 4 0 "Groups with AdminCount=1: "
			WriteWordLine 0 0 "<None>"
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "Error retrieving Group data for domain $($Domain)"
		WriteWordLine 0 0 "Error retrieving Group data for domain $($Domain)" "" $Null 0 $False $True
	}
	Else
	{
		WriteWordLine 0 0 "No Group data was retrieved for domain $($Domain)" "" $Null 0 $False $True
	}
	$First = $False
}
$Groups = $Null
$Admins = $Null
$Members = $Null
$Results = $Null
$DomainInfo = $Null
$User = $Null
$AdminCounts = $Null
$TotalCountStr = $Null
$SecurityCountStr = $Null
$DomainLocalCountStr = $Null
$GlobalCountStr = $Null
$UniversalCountStr = $Null
$DistributionCountStr = $Null
$GroupsWithSIDHistoryStr = $Null
$ContactsCountStr = $Null
$DomainAdminsSID = $Null
$EnterpriseAdminsSID = $Null
$SchemaAdminsSID = $Null
$AdminsCountStr = $Null
$MembersCountStr = $Null
[gc]::collect() 

#GPOs by domain
Write-Verbose "$(Get-Date): Writing domain group policy data"

$selection.InsertNewPage()
WriteWordLine 1 0 "Group Policies by Domain"
$First = $True

ForEach($Domain in $Domains)
{
	Write-Verbose "$(Get-Date): `tProcessing group policies for domain $($Domain)"

	$DomainInfo = Get-ADDomain -Identity $Domain 
	
	If($? -and $DomainInfo -ne $Null)
	{
		If(!$First)
		{
			#put each domain, starting with the second, on a new page
			$selection.InsertNewPage()
		}
		
		If($Domain -eq $ForestRootDomain)
		{
			WriteWordLine 2 0 "$($Domain) (Forest Root)"
		}
		Else
		{
			WriteWordLine 2 0 $Domain
		}

		Write-Verbose "$(Get-Date): `t`tGetting linked GPOs"
		WriteWordLine 3 0 "Linked Group Policy Objects" 
		$LinkedGPOs = $DomainInfo.LinkedGroupPolicyObjects | Sort
		If($LinkedGpos -eq $Null)
		{
			WriteWordLine 0 0 "<None>"
		}
		Else
		{
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 1
			If($LinkedGpos -is [array])
			{
				[int]$Rows = $LinkedGpos.Count
			}
			Else
			{
				[int]$Rows = 1
			}
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$Table.Style = $myHash.Word_TableGrid
			$Table.Borders.InsideLineStyle = $wdLineStyleSingle
			$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
			$GPOArray = @()
			ForEach($LinkedGpo in $LinkedGpos)
			{
				#taken from Michael B. Smith's work on the XenApp 6.x scripts
				#this way we don't need the GroupPolicy module
				$gpObject = [ADSI]( "LDAP://" + $LinkedGPO )
				If($gpObject.DisplayName -eq $Null)
				{
					$p1 = $LinkedGPO.IndexOf("{")
					#38 is length of guid (32) plus the four "-"s plus the beginning "{" plus the ending "}"
					$GUID = $LinkedGPO.SubString($p1,38)
					$tmp = "GPO with GUID $($GUID) was not found in this domain"
				}
				Else
				{
					$tmp = $gpObject.DisplayName	### name of the group policy object
				}
				$GPOArray += $tmp
			}

			$GPOArray = $GPOArray | Sort
			
			[int]$xRow = 0
			ForEach($Item in $GPOArray)
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = $Item
			}
			$GPOArray = $Null

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
			$Table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		$LinkedGPOs = $Null
		$First = $False
	}
	ElseIf(!$?)
	{
		Write-Warning "Error retrieving domain data for domain $($Domain)"
		WriteWordLine 0 0 "Error retrieving domain data for domain $($Domain)" "" $Null 0 $False $True
	}
	Else
	{
		WriteWordLine 0 0 "No Domain data was retrieved for domain $($Domain)" "" $Null 0 $False $True
	}
}
$DomainInfo = $Null
$LinkedGPOs = $Null
$GPOArray = $Null
[gc]::collect() 

#group policies by organizational units
Write-Verbose "$(Get-Date): Writing Group Policy data by Domain by OU"
$selection.InsertNewPage()
WriteWordLine 1 0 "Group Policies by Organizational Unit"
$First = $True

ForEach($Domain in $Domains)
{
	Write-Verbose "$(Get-Date): `tProcessing domain $($Domain)"
	If(!$First)
	{
		#put each domain, starting with the second, on a new page
		$selection.InsertNewPage()
	}
	If($Domain -eq $ForestRootDomain)
	{
		WriteWordLine 2 0 "Group Policies by OUs in Domain $($Domain) (Forest Root)"
	}
	Else
	{
		WriteWordLine 2 0 "Group Policies by OUs in Domain $($Domain)"
	}
	#print disclaimer line in 8 point bold italics
	WriteWordLine 0 0 "(Contains only OUs with linked Group Policies)" "" $Null 8 $True $True
	#get all OUs for the domain
	$OUs = Get-ADOrganizationalUnit -Filter * -Server $Domain -Properties CanonicalName, DistinguishedName, Name  | Select CanonicalName, DistinguishedName, Name | Sort CanonicalName
	
	If($? -and $OUs -ne $Null)
	{
		If($OUs -is [array])
		{
			[int]$NumOUs = $OUs.Count
		}
		Else
		{
			[int]$NumOUs = 1
		}
		[int]$OUCount = 0

		ForEach($OU in $OUs)
		{
			$OUCount++
			$OUDisplayName = $OU.CanonicalName.SubString($OU.CanonicalName.IndexOf("/")+1)
			Write-Verbose "$(Get-Date): `t`tProcessing OU $($OU.CanonicalName) - OU # $OUCount of $NumOUs"
			
			#get data for the individual OU
			$OUInfo = Get-ADOrganizationalUnit -Identity $OU.DistinguishedName -Server $Domain -Properties * 
			
			If($? -and $OUInfo -ne $Null)
			{
				Write-Verbose "$(Get-Date): `t`t`tGetting linked GPOs"
				$LinkedGPOs = $OUInfo.LinkedGroupPolicyObjects | Sort
				If($LinkedGpos -eq $Null)
				{
					# do nothing
				}
				Else
				{
					WriteWordLine 3 0 "$($OUDisplayName)"
					$TableRange = $doc.Application.Selection.Range
					[int]$Columns = 1
					If($LinkedGpos -is [array])
					{
						[int]$Rows = $LinkedGpos.Count
					}
					Else
					{
						[int]$Rows = 1
					}
					$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.Style = $myHash.Word_TableGrid
					$Table.Borders.InsideLineStyle = $wdLineStyleSingle
					$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
					$GPOArray = @()
					ForEach($LinkedGpo in $LinkedGpos)
					{
						#taken from Michael B. Smith's work on the XenApp 6.x scripts
						#this way we don't need the GroupPolicy module
						$gpObject = [ADSI]( "LDAP://" + $LinkedGPO )
						If($gpObject.DisplayName -eq $Null)
						{
							$p1 = $LinkedGPO.IndexOf("{")
							#38 is length of guid (32) plus the four "-"s plus the beginning "{" plus the ending "}"
							$GUID = $LinkedGPO.SubString($p1,38)
							$tmp = "GPO with GUID $($GUID) was not found in this domain"
						}
						Else
						{
							$tmp = $gpObject.DisplayName	### name of the group policy object
						}
						$GPOArray += $tmp
					}

					$GPOArray = $GPOArray | Sort
					
					[int]$xRow = 0
					ForEach($Item in $GPOArray)
					{
						$xRow++
						$Table.Cell($xRow,1).Range.Text = $Item
					}
					$GPOArray = $Null

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
					$Table.AutoFitBehavior($wdAutoFitContent)

					#return focus back to document
					$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

					#move to the end of the current document
					$selection.EndKey($wdStory,$wdMove) | Out-Null
					$TableRange = $Null
					$Table = $Null
				}
			}
			ElseIf(!$?)
			{
				Write-Warning "Error retrieving OU data for OU $($OU.CanonicalName)"
			}
			Else
			{
				$Table.Cell($xRow,1).Range.Text = "<None>"
				$Table.Cell($xRow,2).Range.Text = "<None>"
				$Table.Cell($xRow,3).Range.Text = "<None>"
				$Table.Cell($xRow,4).Range.Text = "<None>"
				$Table.Cell($xRow,5).Range.Text = "<None>"
				$Table.Cell($xRow,6).Range.Text = "<None>"
			}
		}
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
		$Table.AutoFitBehavior($wdAutoFitContent)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf(!$?)
	{
		Write-Warning "Error retrieving OU data for domain $($Domain)"
		WriteWordLine 0 0 "Error retrieving OU data for domain $($Domain)" "" $Null 0 $False $True
	}
	Else
	{
		WriteWordLine 0 0 "No OU data was retrieved for domain $($Domain)" "" $Null 0 $False $True
	}
	$First = $False
}
$OUs = $Null
$OUDisplayName = $Null
$OUInfo = $Null
$LinkedGPOs = $Null
$GPOArray = $Null
[gc]::collect() 

#misc info by domain
Write-Verbose "$(Get-Date): Writing miscellaneous data by domain"

$selection.InsertNewPage()
WriteWordLine 1 0 "Miscellaneous Data by Domain"
$First = $True

ForEach($Domain in $Domains)
{
	Write-Verbose "$(Get-Date): `tProcessing misc data for domain $($Domain)"

	$DomainInfo = Get-ADDomain -Identity $Domain 
	
	If($? -and $DomainInfo -ne $Null)
	{
		If(!$First)
		{
			#put each domain, starting with the second, on a new page
			$selection.InsertNewPage()
		}
		
		If($Domain -eq $ForestRootDomain)
		{
			WriteWordLine 2 0 "$($Domain) (Forest Root)"
		}
		Else
		{
			WriteWordLine 2 0 $Domain
		}

		Write-Verbose "$(Get-Date): `t`tGathering user misc data"
		
		$Users = Get-ADUser -Filter * -Server $Domain -Properties CannotChangePassword, Enabled, LockedOut, PasswordExpired, PasswordNeverExpires, PasswordNotRequired, lastLogonTimestamp, DistinguishedName 
		
		If($? -and $Users -ne $Null)
		{
			If($Users -is [array])
			{
				[int]$UsersCount = $Users.Count
			}
			Else
			{
				[int]$UsersCount = 1
			}
			
			Write-Verbose "$(Get-Date): `t`t`tDisabled users"
			$Results = $Users | Where {$_.Enabled -eq $False}
		
			If($Results -eq $Null)
			{
				[int]$UsersDisabled = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersDisabled = $Results.Count
			}
			Else
			{
				[int]$UsersDisabled = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tUnknown users"
			$Results = $Users | Where {$_.Enabled -eq $Null}
		
			If($Results -eq $Null)
			{
				[int]$UsersUnknown = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersUnknown = $Results.Count
			}
			Else
			{
				[int]$UsersUnknown = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tLocked out users"
			$Results = $Users | Where {$_.LockedOut -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$UsersLockedOut = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersLockedOut = $Results.Count
			}
			Else
			{
				[int]$UsersLockedOut = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tAll users with password expired"
			$Results = $Users | Where {$_.PasswordExpired -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$UsersPasswordExpired = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersPasswordExpired = $Results.Count
			}
			Else
			{
				[int]$UsersPasswordExpired = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tAll users password never expires"
			$Results = $Users | Where {$_.PasswordNeverExpires -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$UsersPasswordNeverExpires = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersPasswordNeverExpires = $Results.Count
			}
			Else
			{
				[int]$UsersPasswordNeverExpires = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tAll users password not required"
			$Results = $Users | Where {$_.PasswordNotRequired -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$UsersPasswordNotRequired = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersPasswordNotRequired = $Results.Count
			}
			Else
			{
				[int]$UsersPasswordNotRequired = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tAll users who cannot change password"
			$Results = $Users | Where {$_.CannotChangePassword -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$UsersCannotChangePassword = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersCannotChangePassword = $Results.Count
			}
			Else
			{
				[int]$UsersCannotChangePassword = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tAll users with SID History"
			$Results = $Null
			$Results = Get-ADObject -LDAPFilter "(sIDHistory=*)" -Server $Domain -Property objectClass, sIDHistory

			If($Results -eq $Null)
			{
				[int]$UsersWithSIDHistory = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersWithSIDHistory = ($Results | Where {$_.objectClass -eq 'user'}).Count
			}
			Else
			{
				[int]$UserssWithSIDHistory = 1
			}

			#active users now
			Write-Verbose "$(Get-Date): `t`t`tActive users"
			$EnabledUsers = $Users | Where {$_.Enabled -eq $True}
		
			If($EnabledUsers -eq $Null)
			{
				[int]$ActiveUsersCount = 0
			}
			ElseIf($EnabledUsers -is [array])
			{
				[int]$ActiveUsersCount = $EnabledUsers.Count
			}
			Else
			{
				[int]$ActiveUsersCount = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tActive users password expired"
			$Results = $EnabledUsers | Where {$_.PasswordExpired -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$ActiveUsersPasswordExpired = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$ActiveUsersPasswordExpired = $Results.Count
			}
			Else
			{
				[int]$ActiveUsersPasswordExpired = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tActive users password never expires"
			$Results = $EnabledUsers | Where {$_.PasswordNeverExpires -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$ActiveUsersPasswordNeverExpires = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$ActiveUsersPasswordNeverExpires = $Results.Count
			}
			Else
			{
				[int]$ActiveUsersPasswordNeverExpires = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tActive users password not required"
			$Results = $EnabledUsers | Where {$_.PasswordNotRequired -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$ActiveUsersPasswordNotRequired = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$ActiveUsersPasswordNotRequired = $Results.Count
			}
			Else
			{
				[int]$ActiveUsersPasswordNotRequired = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tActive Users cannot change password"
			$Results = $EnabledUsers | Where {$_.CannotChangePassword -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$ActiveUsersCannotChangePassword = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$ActiveUsersCannotChangePassword = $Results.Count
			}
			Else
			{
				[int]$ActiveUsersCannotChangePassword = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tActive Users no lastLogonTimestamp"
			$Results = $EnabledUsers | Where {$_.lastLogonTimestamp -eq $Null}
		
			If($Results -eq $Null)
			{
				[int]$ActiveUserslastLogonTimestamp = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$ActiveUserslastLogonTimestamp = $Results.Count
			}
			Else
			{
				[int]$ActiveUserslastLogonTimestamp = 1
			}
		}
		Else
		{
			[int]$UsersCount = 0
			[int]$UsersDisabled = 0
			[int]$UsersLockedOut = 0
			[int]$UsersPasswordExpired = 0
			[int]$UsersPasswordNeverExpires = 0
			[int]$UsersPasswordNotRequired = 0
			[int]$UsersCannotChangePassword = 0
			[int]$UsersWithSIDHistory = 0
			[int]$ActiveUsersCount = 0
			[int]$ActiveUsersPasswordExpired = 0
			[int]$ActiveUsersPasswordNeverExpires = 0
			[int]$ActiveUsersPasswordNotRequired = 0
			[int]$ActiveUsersCannotChangePassword = 0
			[int]$ActiveUserslastLogonTimestamp = 0
		}

		Write-Verbose "$(Get-Date): `t`tFormat numbers into strings"
		[string]$UsersCountStr = "{0,7:N0}" -f $UsersCount
		[string]$UsersDisabledStr = "{0,7:N0}" -f $UsersDisabled
		[string]$UsersUnknownStr = "{0,7:N0}" -f $UsersUnknown
		[string]$UsersLockedOutStr = "{0,7:N0}" -f $UsersLockedOut
		[string]$UsersPasswordExpiredStr = "{0,7:N0}" -f $UsersPasswordExpired
		[string]$UsersPasswordNeverExpiresStr = "{0,7:N0}" -f $UsersPasswordNeverExpires
		[string]$UsersPasswordNotRequiredStr = "{0,7:N0}" -f $UsersPasswordNotRequired
		[string]$UsersCannotChangePasswordStr = "{0,7:N0}" -f $UsersCannotChangePassword
		[string]$UsersWithSIDHistoryStr = "{0,7:N0}" -f $UsersWithSIDHistory
		[string]$ActiveUsersCountStr = "{0,7:N0}" -f $ActiveUsersCount
		[string]$ActiveUsersPasswordExpiredStr = "{0,7:N0}" -f $ActiveUsersPasswordExpired
		[string]$ActiveUsersPasswordNeverExpiresStr = "{0,7:N0}" -f $ActiveUsersPasswordNeverExpires
		[string]$ActiveUsersPasswordNotRequiredStr = "{0,7:N0}" -f $ActiveUsersPasswordNotRequired
		[string]$ActiveUsersCannotChangePasswordStr = "{0,7:N0}" -f $ActiveUsersCannotChangePassword
		[string]$ActiveUserslastLogonTimestampStr = "{0,7:N0}" -f $ActiveUserslastLogonTimestamp

		Write-Verbose "$(Get-Date): `t`tBuild table for All Users"
		WriteWordLine 3 0 "All Users"
		$TableRange   = $doc.Application.Selection.Range
		[int]$Columns = 3
		[int]$Rows = 9
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.Style = $myHash.Word_TableGrid
		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "Total Users"
		$Table.Cell(1,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(1,2).Range.Text = $UsersCountStr
		$Table.Cell(2,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(2,1).Range.Font.Bold = $True
		$Table.Cell(2,1).Range.Text = "Disabled users"
		$Table.Cell(2,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(2,2).Range.Text = $UsersDisabledStr
		[single]$pct = (($UsersDisabled / $UsersCount)*100)
		$pctstr = "{0,5:N2}" -f $pct
		$Table.Cell(2,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(2,3).Range.Text = "$($pctstr)% of Total Users"
		$Table.Cell(3,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(3,1).Range.Font.Bold = $True
		$Table.Cell(3,1).Range.Text = "Unknown users*"
		$Table.Cell(3,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(3,2).Range.Text = $UsersUnknownStr
		[single]$pct = (($UsersUnknown / $UsersCount)*100)
		$pctstr = "{0,5:N2}" -f $pct
		$Table.Cell(3,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(3,3).Range.Text = "$($pctstr)% of Total Users"
		$Table.Cell(4,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(4,1).Range.Font.Bold = $True
		$Table.Cell(4,1).Range.Text = "Locked out users"
		$Table.Cell(4,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(4,2).Range.Text = $UsersLockedOutStr
		[single]$pct = (($UsersLockedOut / $UsersCount)*100)
		$pctstr = "{0,5:N2}" -f $pct
		$Table.Cell(4,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(4,3).Range.Text = "$($pctstr)% of Total Users"
		$Table.Cell(5,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(5,1).Range.Font.Bold = $True
		$Table.Cell(5,1).Range.Text = "Password expired"
		$Table.Cell(5,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(5,2).Range.Text = $UsersPasswordExpiredStr
		[single]$pct = (($UsersPasswordExpired / $UsersCount)*100)
		$pctstr = "{0,5:N2}" -f $pct
		$Table.Cell(5,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(5,3).Range.Text = "$($pctstr)% of Total Users"
		$Table.Cell(6,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(6,1).Range.Font.Bold = $True
		$Table.Cell(6,1).Range.Text = "Password never expires"
		$Table.Cell(6,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(6,2).Range.Text = $UsersPasswordNeverExpiresStr
		[single]$pct = (($UsersPasswordNeverExpires / $UsersCount)*100)
		$pctstr = "{0,5:N2}" -f $pct
		$Table.Cell(6,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(6,3).Range.Text = "$($pctstr)% of Total Users"
		$Table.Cell(7,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(7,1).Range.Font.Bold = $True
		$Table.Cell(7,1).Range.Text = "Password not required"
		$Table.Cell(7,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(7,2).Range.Text = $UsersPasswordNotRequiredStr
		[single]$pct = (($UsersPasswordNotRequired / $UsersCount)*100)
		$pctstr = "{0,5:N2}" -f $pct
		$Table.Cell(7,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(7,3).Range.Text = "$($pctstr)% of Total Users"
		$Table.Cell(8,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(8,1).Range.Font.Bold = $True
		$Table.Cell(8,1).Range.Text = "Can't change password"
		$Table.Cell(8,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(8,2).Range.Text = $UsersCannotChangePasswordStr
		[single]$pct = (($UsersCannotChangePassword / $UsersCount)*100)
		$pctstr = "{0,5:N2}" -f $pct
		$Table.Cell(8,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(8,3).Range.Text = "$($pctstr)% of Total Users"
		$Table.Cell(9,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(9,1).Range.Font.Bold = $True
		$Table.Cell(9,1).Range.Text = "With SID History"
		$Table.Cell(9,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(9,2).Range.Text = $UsersWithSIDHistoryStr
		[single]$pct = (($UsersWithSIDHistory / $UsersCount)*100)
		$pctstr = "{0,5:N2}" -f $pct
		$Table.Cell(9,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(9,3).Range.Text = "$($pctstr)% of Total Users"

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
		$Table.AutoFitBehavior($wdAutoFitContent)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null

		WriteWordLine 0 0 "*Unknown users are user accounts with no Enabled property." "" $Null 8 $False $True
		If($DARights -eq $False)
		{
			WriteWordLine 0 0 "*Rerun the script with Domain Admin rights in $($ADForest)." "" $Null 8 $False $True
		}
		Else
		{
			WriteWordLine 0 0 "*This may be a permissions issue if this is a Trusted Forest." "" $Null 8 $False $True
		}
		
		Write-Verbose "$(Get-Date): `t`tBuild table for Active Users"
		WriteWordLine 3 0 "Active Users"
		$TableRange   = $doc.Application.Selection.Range
		[int]$Columns = 3
		[int]$Rows = 6
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.Style = $myHash.Word_TableGrid
		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "Total Active Users"
		$Table.Cell(1,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(1,2).Range.Text = $ActiveUsersCountStr
		$Table.Cell(2,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(2,1).Range.Font.Bold = $True
		$Table.Cell(2,1).Range.Text = "Password expired"
		$Table.Cell(2,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(2,2).Range.Text = $ActiveUsersPasswordExpiredStr
		[single]$pct = (($ActiveUsersPasswordExpired / $ActiveUsersCount)*100)
		$pctstr = "{0,5:N2}" -f $pct
		$Table.Cell(2,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(2,3).Range.Text = "$($pctstr)% of Active Users"
		$Table.Cell(3,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(3,1).Range.Font.Bold = $True
		$Table.Cell(3,1).Range.Text = "Password never expires"
		$Table.Cell(3,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(3,2).Range.Text = $ActiveUsersPasswordNeverExpiresStr
		[single]$pct = (($ActiveUsersPasswordNeverExpires / $ActiveUsersCount)*100)
		$pctstr = "{0,5:N2}" -f $pct
		$Table.Cell(3,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(3,3).Range.Text = "$($pctstr)% of Active Users"
		$Table.Cell(4,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(4,1).Range.Font.Bold = $True
		$Table.Cell(4,1).Range.Text = "Password not required"
		$Table.Cell(4,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(4,2).Range.Text = $ActiveUsersPasswordNotRequiredStr
		[single]$pct = (($ActiveUsersPasswordNotRequired / $ActiveUsersCount)*100)
		$pctstr = "{0,5:N2}" -f $pct
		$Table.Cell(4,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(4,3).Range.Text = "$($pctstr)% of Active Users"
		$Table.Cell(5,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(5,1).Range.Font.Bold = $True
		$Table.Cell(5,1).Range.Text = "Can't change password"
		$Table.Cell(5,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(5,2).Range.Text = $ActiveUsersCannotChangePasswordStr
		[single]$pct = (($ActiveUsersCannotChangePassword / $ActiveUsersCount)*100)
		$pctstr = "{0,5:N2}" -f $pct
		$Table.Cell(5,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(5,3).Range.Text = "$($pctstr)% of Active Users"
		$Table.Cell(6,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(6,1).Range.Font.Bold = $True
		$Table.Cell(6,1).Range.Text = "No lastLogonTimestamp"
		$Table.Cell(6,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(6,2).Range.Text = $ActiveUserslastLogonTimestampStr
		[single]$pct = (($ActiveUserslastLogonTimestamp / $ActiveUsersCount)*100)
		$pctstr = "{0,5:N2}" -f $pct
		$Table.Cell(6,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
		$Table.Cell(6,3).Range.Text = "$($pctstr)% of Active Users"

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
		$Table.AutoFitBehavior($wdAutoFitContent)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null

		#put computer info on a separate page
		$selection.InsertNewPage()
		
		GetComputerCountByOS $Domain
		
	}
	ElseIf(!$?)
	{
		Write-Warning "Error retrieving domain data for domain $($Domain)"
		WriteWordLine 0 0 "Error retrieving domain data for domain $($Domain)" "" $Null 0 $False $True
	}
	Else
	{
		WriteWordLine 0 0 "No Domain data was retrieved for domain $($Domain)" "" $Null 0 $False $True
	}
	$First = $False
}
$Domains = $Null
$DomainInfo = $Null
$Users = $Null
$EnabledUsers = $Null
$Results = $Null
$UsersCountStr = $Null
$UsersDisabledStr = $Null
$UsersUnknownStr = $Null
$UsersLockedOutStr = $Null
$UsersPasswordExpiredStr = $Null
$UsersPasswordNeverExpiresStr = $Null
$UsersPasswordNotRequiredStr = $Null
$UsersCannotChangePasswordStr = $Null
$UsersWithSIDHistoryStr = $Null
$ActiveUsersCountStr = $Null
$ActiveUsersPasswordExpiredStr = $Null
$ActiveUsersPasswordNeverExpiresStr = $Null
$ActiveUsersPasswordNotRequiredStr = $Null
$ActiveUsersCannotChangePasswordStr = $Null
$ActiveUserslastLogonTimestampStr = $Null
$pctstr = $Null
[gc]::collect() 

Write-Verbose "$(Get-Date): Finishing up Word document"
#end of document processing

#Update document properties
If($CoverPagesExist)
{
	Write-Verbose "$(Get-Date): Set Cover Page Properties"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Company" $CompanyName
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Title" $title
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Subject" "Active Directory Inventory"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Author" $username

	#Get the Coverpage XML part
	$cp = $doc.CustomXMLParts | Where {$_.NamespaceURI -match "coverPageProps$"}

	#get the abstract XML part
	$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}

	#set the text
	If([String]::IsNullOrEmpty($CompanyName))
	{
		[string]$abstract = "Microsoft Active Directory Inventory"
	}
	Else
	{
		[string]$abstract = "Microsoft Active Directory Inventory for $CompanyName"
	}

	$ab.Text = $abstract

	$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
	#set the text
	[string]$abstract = (Get-Date -Format d).ToString()
	$ab.Text = $abstract

	Write-Verbose "$(Get-Date): Update the Table of Contents"
	#update the Table of Contents
	$doc.TablesOfContents.item(1).Update()
	$cp = $Null
	$ab = $Null
	$abstract = $Null
}

#bug fix 1-Apr-2014
#reset Grammar and Spelling options back to their original settings
$Word.Options.CheckGrammarAsYouType = $CurrentGrammarOption
$Word.Options.CheckSpellingAsYouType = $CurrentSpellingOption

Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
If($WordVersion -eq $wdWord2007)
{
	#Word 2007
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
	}
	Else
	{
		Write-Verbose "$(Get-Date): Saving DOCX file"
	}
	Write-Verbose "$(Get-Date): Running Word 2007 and detected operating system $($RunningOS)"
	If($RunningOS.Contains("Server 2008 R2") -or $RunningOS.Contains("Server 2012"))
	{
		$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
		$doc.SaveAs($filename1, $SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$SaveFormat = $wdSaveFormatPDF
			$doc.SaveAs($filename2, $SaveFormat)
		}
	}
	Else
	{
		#works for Server 2008 and Windows 7
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$doc.SaveAs([REF]$filename1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$doc.SaveAs([REF]$filename2, [ref]$saveFormat)
		}
	}
}
Else
{
	#the $saveFormat below passes StrictMode 2
	#I found this at the following two links
	#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
	#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
	}
	Else
	{
		Write-Verbose "$(Get-Date): Saving DOCX file"
	}
	$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
	$doc.SaveAs([REF]$filename1, [ref]$SaveFormat)
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Now saving as PDF"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
		$doc.SaveAs([REF]$filename2, [ref]$saveFormat)
	}
}

Write-Verbose "$(Get-Date): Closing Word"
$doc.Close()
$Word.Quit()
If($PDF)
{
	Write-Verbose "$(Get-Date): Deleting $($filename1) since only $($filename2) is needed"
	Remove-Item $filename1
}
Write-Verbose "$(Get-Date): System Cleanup"
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
If(Test-Path variable:global:word)
{
	Remove-Variable -Name word -Scope Global
}
$SaveFormat = $Null
[gc]::collect() 
[gc]::WaitForPendingFinalizers()
Write-Verbose "$(Get-Date): Script has completed"
Write-Verbose "$(Get-Date): "

If($PDF)
{
	Write-Verbose "$(Get-Date): $($filename2) is ready for use"
}
Else
{
	Write-Verbose "$(Get-Date): $($filename1) is ready for use"
}
Write-Verbose "$(Get-Date): "

#http://poshtips.com/measuring-elapsed-time-in-powershell/
Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
$runtime = $(Get-Date) - $Script:StartTime
$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
	$runtime.Days, `
	$runtime.Hours, `
	$runtime.Minutes, `
	$runtime.Seconds,
	$runtime.Milliseconds)
Write-Verbose "$(Get-Date): Elapsed time: $($Str)"
$runtime = $Null
$Str = $Null
$ErrorActionPreference = $SaveEAPreference
# SIG # Begin signature block
# MIIiywYJKoZIhvcNAQcCoIIivDCCIrgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUBYuJmVu0I0C/KB/4mEj8yZnv
# 6yCggh41MIIDtzCCAp+gAwIBAgIQDOfg5RfYRv6P5WD8G/AwOTANBgkqhkiG9w0B
# AQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMzExMTEwMDAwMDAwWjBlMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3Qg
# Q0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCtDhXO5EOAXLGH87dg
# +XESpa7cJpSIqvTO9SA5KFhgDPiA2qkVlTJhPLWxKISKityfCgyDF3qPkKyK53lT
# XDGEKvYPmDI2dsze3Tyoou9q+yHyUmHfnyDXH+Kx2f4YZNISW1/5WBg1vEfNoTb5
# a3/UsDg+wRvDjDPZ2C8Y/igPs6eD1sNuRMBhNZYW/lmci3Zt1/GiSw0r/wty2p5g
# 0I6QNcZ4VYcgoc/lbQrISXwxmDNsIumH0DJaoroTghHtORedmTpyoeb6pNnVFzF1
# roV9Iq4/AUaG9ih5yLHa5FcXxH4cDrC0kqZWs72yl+2qp/C3xag/lRbQ/6GW6whf
# GHdPAgMBAAGjYzBhMA4GA1UdDwEB/wQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBRF66Kv9JLLgjEtUYunpyGd823IDzAfBgNVHSMEGDAWgBRF66Kv9JLL
# gjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEAog683+Lt8ONyc3pklL/3
# cmbYMuRCdWKuh+vy1dneVrOfzM4UKLkNl2BcEkxY5NM9g0lFWJc1aRqoR+pWxnmr
# EthngYTffwk8lOa4JiwgvT2zKIn3X/8i4peEH+ll74fg38FnSbNd67IJKusm7Xi+
# fT8r87cmNW1fiQG2SVufAQWbqz0lwcy2f8Lxb4bG+mRo64EtlOtCt/qMHt1i8b5Q
# Z7dsvfPxH2sMNgcWfzd8qVttevESRmCD1ycEvkvOl77DZypoEd+A5wwzZr8TDRRu
# 838fYxAe+o0bJW1sj6W3YQGx0qMmoRBxna3iw/nDmVG3KwcIzi7mULKn+gpFL6Lw
# 8jCCBmowggVSoAMCAQICEAOf7e3LeVuN7TIMiRnwNokwDQYJKoZIhvcNAQEFBQAw
# YjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBD
# QS0xMB4XDTEzMDUyMTAwMDAwMFoXDTE0MDYwNDAwMDAwMFowRzELMAkGA1UEBhMC
# VVMxETAPBgNVBAoTCERpZ2lDZXJ0MSUwIwYDVQQDExxEaWdpQ2VydCBUaW1lc3Rh
# bXAgUmVzcG9uZGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAumlK
# gU1vpRQWqorNZ75Lv8Zpj1gc4HnoHp1YJpjaXNR8o/nbK4wSNsP8+WQGsbvCqJgK
# Fw3hletAtOuWbZi/po95z7yKknttnBgGUdilGFMyAScZYeiEQd/G8OjK/netX9ie
# e4xgb4VcRr1r5w+AzucDw3wxz7dlVcb74JkI5HNa+5fa0Ey+tLbGD38mkqm4/Dju
# tOQ6pEjQTOqpRidbz5IRk5wWp/7SrR8ixR6swXHvvErbAQlE35gcLWe6qIoDM8lR
# tfcCTQmkTf6AXsXXRcN9CKoBM8wz2E8wFuT/IjIu63478PkeMuuVJdLy/m1UhLrV
# 5dTR3RuvvVl7lIUwAQIDAQABo4IDNTCCAzEwDgYDVR0PAQH/BAQDAgeAMAwGA1Ud
# EwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwggG/BgNVHSAEggG2MIIB
# sjCCAaEGCWCGSAGG/WwHATCCAZIwKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9DUFMwggFkBggrBgEFBQcCAjCCAVYeggFSAEEAbgB5ACAAdQBz
# AGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAYwBv
# AG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBuAGMAZQAgAG8AZgAg
# AHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQAFMAIABhAG4AZAAg
# AHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAgAEEAZwByAGUAZQBt
# AGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABpAGEAYgBpAGwAaQB0
# AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwByAGEAdABlAGQAIABo
# AGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBlAC4wCwYJYIZIAYb9
# bAMVMB8GA1UdIwQYMBaAFBUAEisTmLKZB+0e36K+Vw0rZwLNMB0GA1UdDgQWBBRj
# L8nfeZJ7tSPKu+Gk7jN+4+Kd+jB9BgNVHR8EdjB0MDigNqA0hjJodHRwOi8vY3Js
# My5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDA4oDagNIYy
# aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5j
# cmwwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCr
# dL1AAEx2FSVXPdMcA/99RchFEmbnKGVg2N87s/oNwawzj/SBuWHxnfuYVdfeR0O6
# gD3xSMw/ZzBWH8700EyEvYeknsXhD6gGXdAvbl7cGejwh+rgTq89bCCOc29+1ocY
# 4IbTmvye6oxy6UEPuHG1OCz4KbLVHKKdG+xfKrjcNyDhy7vw0GxspbPLn0r2VOMm
# ND0uuMErHLf2wz3+0S0eUPSUyPj97nPbSbUb9PX/pZDBORQb2O1xG2qY+/pAmkSp
# KQ5VXni4t6SDw3AB8GZA5a55NOErTQOhLebbVGIY7dUJi6Kq1gzITxq+mSV4aZmJ
# 1FmJ3t+I8NNnXnSlnaZEMIIGkDCCBXigAwIBAgIQBKVRftX3ANDrw0+OjYS9xjAN
# BgkqhkiG9w0BAQUFADBvMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQg
# SW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMS4wLAYDVQQDEyVEaWdpQ2Vy
# dCBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQS0xMB4XDTExMDkzMDAwMDAwMFoX
# DTE0MTAwODEyMDAwMFowXDELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAlROMRIwEAYD
# VQQHEwlUdWxsYWhvbWExFTATBgNVBAoTDENhcmwgV2Vic3RlcjEVMBMGA1UEAxMM
# Q2FybCBXZWJzdGVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAz2g4
# Kup2X6Mscbuq96HnetDDiITbncV1LtQ8Rxf8ZtN00+O/TliIZsWtufMq7GsLj1D8
# ikWfcgWGqMngWMsVYB4vdr1B8aQuHmKWld7W+j8FhKp3l+rNuFviTGa62sR6fEVW
# 1N6lDtJJHpfSIg/FUFfAqOKl0gFc45PU7iWCh08+oG5FJdhZ3WY0SosS1QujKEA4
# riSjeXPV6XSLsAHTE/fmHlGuu7NzJyMUzNNz2gPOFxYupHygbduhM5aAItD6GJ1h
# ajlovRt71tAMyeIPWNjj9B2luXxfRbgO9eufw91uFrXnougBPa7/eQ25YdW3NcGf
# tosYjvVI6Ptw/AaSiQIDAQABo4IDOTCCAzUwHwYDVR0jBBgwFoAUe2jOKarAF75J
# euHlP9an90WPNTIwHQYDVR0OBBYEFMHndyU+4pRT+JRECX9EG4y1laDkMA4GA1Ud
# DwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzBzBgNVHR8EbDBqMDOgMaAv
# hi1odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vYXNzdXJlZC1jcy0yMDExYS5jcmww
# M6AxoC+GLWh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9hc3N1cmVkLWNzLTIwMTFh
# LmNybDCCAcQGA1UdIASCAbswggG3MIIBswYJYIZIAYb9bAMBMIIBpDA6BggrBgEF
# BQcCARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5
# Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAg
# AHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0
# AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABE
# AGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABS
# AGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3
# AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBk
# ACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBu
# ACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjCBggYIKwYBBQUHAQEEdjB0MCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTAYIKwYBBQUHMAKG
# QGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENv
# ZGVTaWduaW5nQ0EtMS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQUFAAOC
# AQEAm1zhveo2Zy2lp8UNpR2E2CE8/NvEk0NDLszcBBuMda3N8Du23CikXCgrVvE0
# 3mMaeu/cIMDVU01ityLaqvDuovmTsvAKqaSJNztV9yTeWK9H4+h+35UEIU5TvYLs
# uzEW+rI5M2KcCXR6/LF9ZPmnBf9hHnK44hweHpmDWbo8HPqMatnIo7ideucuDn/D
# BM6s63eTMsFQCPYwte5vxuyVLqodOubLvIOMezZzByrpvJp9+gWAL151CE4qR6xQ
# jpgk5KqSkkkyvl72D+3PhNwZuxZDbZil5PIcrjmaBYoG8wfJzoNrtPFq3aG8dnQr
# xjXJjl+IN1iHYehBAUoBX98EozCCBqMwggWLoAMCAQICEA+oSQYV1wCgviF2/cXs
# bb0wDQYJKoZIhvcNAQEFBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGln
# aUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTExMDIxMTEyMDAwMFoXDTI2MDIx
# MDEyMDAwMFowbzELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZ
# MBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEuMCwGA1UEAxMlRGlnaUNlcnQgQXNz
# dXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EtMTCCASIwDQYJKoZIhvcNAQEBBQADggEP
# ADCCAQoCggEBAJx8+aCPCsqJS1OaPOwZIn8My/dIRNA/Im6aT/rO38bTJJH/qFKT
# 53L48UaGlMWrF/R4f8t6vpAmHHxTL+WD57tqBSjMoBcRSxgg87e98tzLuIZARR9P
# +TmY0zvrb2mkXAEusWbpprjcBt6ujWL+RCeCqQPD/uYmC5NJceU4bU7+gFxnd7XV
# b2ZklGu7iElo2NH0fiHB5sUeyeCWuAmV+UuerswxvWpaQqfEBUd9YCvZoV29+1aT
# 7xv8cvnfPjL93SosMkbaXmO80LjLTBA1/FBfrENEfP6ERFC0jCo9dAz0eotyS+BW
# tRO2Y+k/Tkkj5wYW8CWrAfgoQebH1GQ7XasCAwEAAaOCA0MwggM/MA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDAzCCAcMGA1UdIASCAbowggG2MIIB
# sgYIYIZIAYb9bAMwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRpZ2ljZXJ0
# LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIwggFWHoIB
# UgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkA
# YwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEA
# bgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMA
# UABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkA
# IABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwA
# aQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8A
# cgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMA
# ZQAuMBIGA1UdEwEB/wQIMAYBAf8CAQAweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUF
# BzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6
# Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5j
# cnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmw0LmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwHQYDVR0OBBYE
# FHtozimqwBe+SXrh5T/Wp/dFjzUyMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA0GCSqGSIb3DQEBBQUAA4IBAQB7ch1k/4jIOsG36eepxIe725SS15BZ
# M/orh96oW4AlPxOPm4MbfEPE5ozfOT7DFeyw2jshJXskwXJduEeRgRNG+pw/alE4
# 3rQly/Cr38UoAVR5EEYk0TgPJqFhkE26vSjmP/HEqpv22jVTT8nyPdNs3CPtqqBN
# ZwnzOoA9PPs2TJDndqTd8jq/VjUvokxl6ODU2tHHyJFqLSNPNzsZlBjU1ZwQPNWx
# HBn/j8hrm574rpyZlnjRzZxRFVtCJnJajQpKI5JA6IbeIsKTOtSbaKbfKX8GuTwO
# vZ/EhpyCR0JxMoYJmXIJeUudcWn1Qf9/OXdk8YSNvosesn1oo6WQsQz/MIIGzTCC
# BbWgAwIBAgIQBv35A5YDreoACus/J7u6GzANBgkqhkiG9w0BAQUFADBlMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0Ew
# HhcNMDYxMTEwMDAwMDAwWhcNMjExMTEwMDAwMDAwWjBiMQswCQYDVQQGEwJVUzEV
# MBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29t
# MSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQDogi2Z+crCQpWlgHNAcNKeVlRcqcTSQQaPyTP8
# TUWRXIGf7Syc+BZZ3561JBXCmLm0d0ncicQK2q/LXmvtrbBxMevPOkAMRk2T7It6
# NggDqww0/hhJgv7HxzFIgHweog+SDlDJxofrNj/YMMP/pvf7os1vcyP+rFYFkPAy
# IRaJxnCI+QWXfaPHQ90C6Ds97bFBo+0/vtuVSMTuHrPyvAwrmdDGXRJCgeGDboJz
# PyZLFJCuWWYKxI2+0s4Grq2Eb0iEm09AufFM8q+Y+/bOQF1c9qjxL6/siSLyaxhl
# scFzrdfx2M8eCnRcQrhofrfVdwonVnwPYqQ/MhRglf0HBKIJAgMBAAGjggN6MIID
# djAOBgNVHQ8BAf8EBAMCAYYwOwYDVR0lBDQwMgYIKwYBBQUHAwEGCCsGAQUFBwMC
# BggrBgEFBQcDAwYIKwYBBQUHAwQGCCsGAQUFBwMIMIIB0gYDVR0gBIIByTCCAcUw
# ggG0BgpghkgBhv1sAAEEMIIBpDA6BggrBgEFBQcCARYuaHR0cDovL3d3dy5kaWdp
# Y2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIB
# Vh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkA
# ZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAA
# dABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAA
# LwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIA
# dAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQA
# IABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIA
# cABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUA
# bgBjAGUALjALBglghkgBhv1sAxUwEgYDVR0TAQH/BAgwBgEB/wIBADB5BggrBgEF
# BQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBD
# BggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0
# QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDig
# NoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNybDAdBgNVHQ4EFgQUFQASKxOYspkH7R7for5XDStnAs0wHwYDVR0jBBgw
# FoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQEFBQADggEBAEZQPsm3
# KCSnOB22WymvUs9S6TFHq1Zce9UNC0Gz7+x1H3Q48rJcYaKclcNQ5IK5I9G6OoZy
# rTh4rHVdFxc0ckeFlFbR67s2hHfMJKXzBBlVqefj56tizfuLLZDCwNK1lL1eT7EF
# 0g49GqkUW6aGMWKoqDPkmzmnxPXOHXh2lCVz5Cqrz5x2S+1fwksW5EtwTACJHvzF
# ebxMElf+X+EevAJdqP77BzhPDcZdkbkPZ0XN1oPt55INjbFpjE/7WeAjD9KqrgB8
# 7pxCDs+R1ye3Fu4Pw718CqDuLAhVhSK46xgaTfwqIa1JMYNHlXdx3LEbS0scEJx3
# FMGdTy9alQgpECYxggQAMIID/AIBATCBgzBvMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMS4wLAYD
# VQQDEyVEaWdpQ2VydCBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQS0xAhAEpVF+
# 1fcA0OvDT46NhL3GMAkGBSsOAwIaBQCgQDAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAjBgkqhkiG9w0BCQQxFgQUKQiWzyIheMqQQ5wOvDijZHjzzDowDQYJKoZI
# hvcNAQEBBQAEggEAagAwfEPgHunFsG+K/H28030/Hea5JRyf6OWRI0Y7aFZ0kRB4
# k9bUwH9JkByuchqMmI6dTy+/AQ50iSPplz4iazWnTsuaWQ3l+IvMIfVwBjUK0IrH
# 3CzzK+PsVTi3A1r5cnIbdqU2EUYHzE1e6FNP+6jWZV51Org2CbSsMaczQ87NG2t3
# 1Io+UYsaLN83ZQ4Fv/iQPpKMI5Y4YRYkkXkBbGGLOvBPRS60bgvLFv1X8B1Tj6TN
# AsgpBtaUAloFlHWFIu410h5h6kbVk6bH6vyWFaY0uv1CLz54HiZZTsVE9OT3TDk+
# 89GNlmbllCg11zAQnXcNdzcY/67lM19RmM9/Q6GCAg8wggILBgkqhkiG9w0BCQYx
# ggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IEFzc3VyZWQgSUQgQ0EtMQIQA5/t7ct5W43tMgyJGfA2iTAJBgUrDgMCGgUAoF0w
# GAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTQwNTI4
# MDIzMDUwWjAjBgkqhkiG9w0BCQQxFgQUD33LkHnxheUmyU8bkOP5/iKRmfYwDQYJ
# KoZIhvcNAQEBBQAEggEAiEX/msGhVrpcidwsWOwNl90GNCUkgfggSR93WasCjyGu
# hZ47gb+jzc1cfBPSmUZxCETqP1u1Jl9wVmaVvl1fhGbZOWAzfMXzn98GLu4gnbz0
# 4IzosxnLzg9ighZMFSZ0Lq/IE6HLGrniblCK/iMDe67SX7UfB+8fPfqQ4SZFH/zP
# +NZ4Oz2vshK4BIT79OuvCE7fQtsCXi5+7Id+gQr5gvfUKdAI7NTDqIHBkAhB/ruc
# YILCoUKihKS2M8UzQij9He6WO1cdqOm45NhX6xWdV3mvTh120OzRJtDNyZGaS0Wm
# t29gwuU5puoRzPf0u0ncZu6CXUd3sMvk/RWUB3DCqQ==
# SIG # End signature block
