#include <WinAPIFiles.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <WindowsConstants.au3>
#include <Date.au3>
#include <File.au3>
#include <Math.au3>
#include <Misc.au3>

If _Singleton("maj_scanbac", 1) = 0 Then Exit

Local $sFilePath = @ScriptDir & "\config_maj.ini"
Local $ini_ana_enr = StringSplit(IniRead($sFilePath, "General", "ini_ana_enr", ""), "|") ; Codes analyse à surveiller à l'enregistrement pour intégration
Local $ini_lign = StringSplit(IniRead($sFilePath, "General", "OBX_a_copier", ""), "|") ; Codes des lignes d'analyses à surveiller pour type en bactério
Local $ini_ana_edi = StringSplit(IniRead($sFilePath, "General", "ini_ana_edi", ""), "|") ; Codes analyse à surveiller à l'edition pour suppression
Local $ini_ana_enr2 = StringSplit(IniRead($sFilePath, "General", "ini_ana_enr2", ""), "|") ; Codes analyse à surveiller à l'enregistrement pour intégration CERBA
Local $ini_ana_edi2 = StringSplit(IniRead($sFilePath, "General", "ini_ana_edi", ""), "|") ; Codes analyse à surveiller à l'edition pour suppression CERBA


#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Interface SCANNER BioWin", 624, 135, 192, 124)
$Label1 = GUICtrlCreateLabel("", 8, 8, 601, 45)
$Progress1 = GUICtrlCreateProgress(8, 64, 601, 25)
$Button1 = GUICtrlCreateButton("Quitter", 216, 96, 185, 33)
GUICtrlSetState(-1, $GUI_DISABLE)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

;~ COPIE DES FICHIERS
GUICtrlSetData($Progress1, 0)
GUICtrlSetData($Label1, "Chargement..." & @CRLF & "Etape 1/5" & @CRLF & "Copie des fichiers log depuis BioWin...")
local $afich[3]
$afich[0] = "TRACE_RESULAT_" & StringReplace(_DateAdd("D", -2, stringleft(_nowcalc(), 10)), "/", "") & ".log"
$afich[1] = "TRACE_RESULAT_" & StringReplace(_DateAdd("D", -1, stringleft(_nowcalc(), 10)), "/", "") & ".log"
$afich[2] = "TRACE_RESULAT_" & StringReplace(stringleft(_nowcalc(), 10), "/", "") & ".log"
FileCopy("Y:\Biowin\DEBUG\" & $afich[0], @ScriptDir & "\" & $afich[0], 1)
FileCopy("Y:\Biowin\DEBUG\" & $afich[1], @ScriptDir & "\" & $afich[1], 1)
FileCopy("Y:\Biowin\DEBUG\" & $afich[2], @ScriptDir & "\" & $afich[2], 1)


FileCopy("Y:\Biowin\Doscli.fic", @ScriptDir & "\Doscli.fic", 1)
FileCopy("Y:\Biowin\Doscli.ndx", @ScriptDir & "\Doscli.ndx", 1)
Runwait(@ScriptDir & "\hyperfile2xml.exe Doscli.fic Doscli.xml")

;SUPPRESSSION ANCIENS FICHIERS LOG (> 20 jours)
for $i = -60 to -21
	FileDelete("TRACE_RESULAT_" & StringReplace(_DateAdd("D", $i, stringleft(_nowcalc(), 10)), "/", "") & ".log")
Next

;LECTURE FICHIER DOSCLI
GUICtrlSetData($Progress1, 25)
GUICtrlSetData($Label1, "Chargement..." & @CRLF & "Etape 2/5" & @CRLF & "Lecture du fichier Patients...")
Local $hFileOpen = FileOpen(@ScriptDir & "\Doscli.xml", $FO_READ)
FileReadLine($hFileOpen)
FileReadLine($hFileOpen)
Local $adoscli[1][4]
While 1 = 1
	FileReadLine($hFileOpen)
	if @error = -1 then ExitLoop ; Si on est à la fin, on quitte la bouche
	$txt = FileReadLine($hFileOpen) ; Lecture de la ligne du fichier source
	if @error = -1 then ExitLoop ; Si on est à la fin, on quitte la bouche
	Local $add[1][4]
	$add[0][0] = StringStripWS(StringSplit(StringSplit($txt, ">")[2], "<")[1], 8) & StringStripWS(StringSplit(StringSplit(FileReadLine($hFileOpen), ">")[2], "<")[1], 3)
	$add[0][1] = StringStripWS(StringSplit(StringSplit(FileReadLine($hFileOpen), ">")[2], "<")[1], 3)
	$add[0][2] = StringStripWS(StringSplit(StringSplit(FileReadLine($hFileOpen), ">")[2], "<")[1], 3)
	for $t = 1 to 11
		$txt = FileReadLine($hFileOpen) ; Lecture de la ligne du fichier source
		if @error = -1 then ExitLoop(2) ; Si on est à la fin, on quitte les 2 boucles
		if StringStripWS($txt, 6) = "<MED2>" Then ExitLoop
	Next

	$add[0][3] = StringStripWS(StringSplit(StringSplit(FileReadLine($hFileOpen), ">")[2], "<")[1], 3)
	_ArrayAdd($adoscli, $add, 0)
	for $t = 1 to 1000
		$txt = FileReadLine($hFileOpen) ; Lecture de la ligne du fichier source
		if @error = -1 then ExitLoop(2) ; Si on est à la fin, on quitte les 2 boucles
		if StringStripWS($txt, 8) = "</Data>" Then ExitLoop
	Next
	if $t = 1000 Then
		msgox(0, "", "Erreur lecture doscli.xml")
		Exit
	EndIf
WEnd
;~ _ArrayDisplay($adoscli)
FileClose($hFileOpen)

;TRAITEMENT FICHIERS LOG
Global $encours, $encoursCERBA
for $fich = 0 to 2
	GUICtrlSetData($Progress1, 25 + $fich * 25)
	GUICtrlSetData($Label1, "Chargement..." & @CRLF & "Etape " & 2 + $fich & "/5" & @CRLF & "Lecture du fichier log " & $afich[$fich] & "...")
	_FileReadToArray(@ScriptDir & "\encours.txt", $encours, 0, "|")
	_FileReadToArray(StringReplace(@ScriptDir & "\encoursCERBA.txt", "scanbac", "cerba"), $encoursCERBA, 0, "|")
	Local $hFileOpen = FileOpen(@ScriptDir & "\" & $afich[$fich], $FO_READ)
	While 1 = 1
		$txt = FileReadLine($hFileOpen) ; Lecture de la ligne du fichier source
		if @error = -1 then ExitLoop ; Si on est à la fin, on quitte la bouche
		$decod = False
		StringReplace($txt, ";", "")
		if @extended > 9 Then
			if StringSplit($txt, ";")[6] = "DOSSIER" or StringSplit($txt, ";")[6] = "NOUVANA" or StringSplit($txt, ";")[6] = "RESULDOS" or StringSplit($txt, ";")[6] = "EDITRES" then
;~ ConsoleWrite($afich[$fich]&":"&$txt&@crlf)
				for $i = 1 to $ini_ana_enr[0] ; INTERCEPTION AJOUT ANALYSE SUR ECRAN GESTION DES DOSSIERS
					if StringSplit($txt, ";")[10] = "Analyse = " & $ini_ana_enr[$i] and StringSplit($txt, ";")[2] = "AJOUT" and StringSplit($txt, ";")[6] = "DOSSIER" Then ;Création en gestion de dossier
;~ 							  ConsoleWrite($i)
						$ret = _ArrayFindAll($adoscli, StringSplit($txt, ";")[9], 0, 0, 0, 0, 0, 0)
						if UBound($ret) > 2 Then
							msgbox(0, "Erreur sur le dossier " & $txt1)
							Exit
						EndIf
						$nom = ""
						$prenom = ""
						$serv = ""
						if UBound($ret) = 1 Then
							$nom = $adoscli[$ret[0]][1]
							$prenom = $adoscli[$ret[0]][2]
							$serv = $adoscli[$ret[0]][3]
						EndIf

						MAJ( _
								StringSplit($txt, ";")[9], _ ;Dossier
								Stringleft(StringSplit($txt, ";")[1], 19), _ ; Date-Heure
								$nom, _ ; Nom
								$prenom, _ ; Prenom
								$serv, _ ; Service
								$ini_ana_enr[$i], _ ; Examen
								"0", _ ; Etat
								Stringleft(StringSplit($txt, ";")[1], 19)) ; MAJ
						$decod = True
					EndIf
				Next

				if $decod = False Then
					for $i = 1 to $ini_ana_enr2[0] ; INTERCEPTION AJOUT ANALYSE SUR ECRAN GESTION DES DOSSIERS (CERBA)
						if StringSplit($txt, ";")[10] = "Analyse = " & $ini_ana_enr2[$i] and StringSplit($txt, ";")[2] = "AJOUT" and StringSplit($txt, ";")[6] = "DOSSIER" Then ;Création en gestion de dossier
;~ 							  ConsoleWrite($i)
							$ret = _ArrayFindAll($adoscli, StringSplit($txt, ";")[9], 0, 0, 0, 0, 0, 0)
							if UBound($ret) > 2 Then
								msgbox(0, "Erreur sur le dossier " & $txt1)
								Exit
							EndIf
							$nom = ""
							$prenom = ""
							$serv = ""
							if UBound($ret) = 1 Then
								$nom = $adoscli[$ret[0]][1]
								$prenom = $adoscli[$ret[0]][2]
								$serv = $adoscli[$ret[0]][3]
							EndIf

							MAJ_CERBA( _
									StringSplit($txt, ";")[9], _ ;Dossier
									Stringleft(StringSplit($txt, ";")[1], 19), _ ; Date-Heure
									$nom, _ ; Nom
									$prenom, _ ; Prenom
									$serv, _ ; Service
									$ini_ana_enr2[$i], _ ; Examen
									"0", _ ; Etat
									Stringleft(StringSplit($txt, ";")[1], 19)) ; MAJ
							$decod = True
						EndIf
					Next
				Endif


				if $decod = False Then
					for $i = 1 to $ini_ana_enr[0] ; INTERCEPTION AJOUT ANALYSE SUR ECRAN SAISIE PAR DOSSIER
						if StringSplit($txt, ";")[10] = "Analyse = " & $ini_ana_enr[$i] and StringSplit($txt, ";")[2] = "AJOUT" and StringSplit($txt, ";")[6] = "NOUVANA" Then ;Création en gestion de dossier
							$ret = _ArrayFindAll($adoscli, StringSplit($txt, ";")[9], 0, 0, 0, 0, 0, 0)
							if UBound($ret) > 2 Then
								msgbox(0, "Erreur sur le dossier " & $txt1)
								Exit
							EndIf
							$nom = ""
							$prenom = ""
							$serv = ""
							if UBound($ret) = 1 Then
								$nom = $adoscli[$ret[0]][1]
								$prenom = $adoscli[$ret[0]][2]
								$serv = $adoscli[$ret[0]][3]
							EndIf

							MAJ( _
									StringSplit($txt, ";")[9], _ ;Dossier
									Stringleft(StringSplit($txt, ";")[1], 19), _ ; Date-Heure
									$nom, _ ; Nom
									$prenom, _ ; Prenom
									$serv, _ ; Service
									$ini_ana_enr[$i], _ ; Examen
									"0", _ ; Etat
									Stringleft(StringSplit($txt, ";")[1], 19)) ; MAJ
							$decod = True
						EndIf
					Next
				EndIf

				if $decod = False Then
					for $i = 1 to $ini_ana_enr2[0] ; INTERCEPTION AJOUT ANALYSE SUR ECRAN SAISIE PAR DOSSIER (CERBA)
						if StringSplit($txt, ";")[10] = "Analyse = " & $ini_ana_enr2[$i] and StringSplit($txt, ";")[2] = "AJOUT" and StringSplit($txt, ";")[6] = "NOUVANA" Then ;Création en gestion de dossier
							$ret = _ArrayFindAll($adoscli, StringSplit($txt, ";")[9], 0, 0, 0, 0, 0, 0)
							if UBound($ret) > 2 Then
								msgbox(0, "Erreur sur le dossier " & $txt1)
								Exit
							EndIf
							$nom = ""
							$prenom = ""
							$serv = ""
							if UBound($ret) = 1 Then
								$nom = $adoscli[$ret[0]][1]
								$prenom = $adoscli[$ret[0]][2]
								$serv = $adoscli[$ret[0]][3]
							EndIf

							MAJ_CERBA( _
									StringSplit($txt, ";")[9], _ ;Dossier
									Stringleft(StringSplit($txt, ";")[1], 19), _ ; Date-Heure
									$nom, _ ; Nom
									$prenom, _ ; Prenom
									$serv, _ ; Service
									$ini_ana_enr2[$i], _ ; Examen
									"0", _ ; Etat
									Stringleft(StringSplit($txt, ";")[1], 19)) ; MAJ
							$decod = True
						EndIf
					Next
				EndIf

				if $decod = False Then
					for $i = 1 to $ini_ana_enr[0] ; INTERCEPTION MODIFICATION ANALYSE SUR ECRAN SAISIE PAR DOSSIER
						if StringSplit($txt, ";")[10] = "Analyse = " & $ini_ana_enr[$i] and StringSplit($txt, ";")[2] = "MODIFICATION" and StringSplit($txt, ";")[6] = "RESULDOS" Then ;Création en gestion de dossier
							$ret = _ArrayFindAll($adoscli, StringSplit($txt, ";")[9], 0, 0, 0, 0, 0, 0)
							if UBound($ret) > 2 Then
								msgbox(0, "Erreur sur le dossier " & $txt1)
								Exit
							EndIf
							$nom = ""
							$prenom = ""
							$serv = ""
							if UBound($ret) = 1 Then
								$nom = $adoscli[$ret[0]][1]
								$prenom = $adoscli[$ret[0]][2]
								$serv = $adoscli[$ret[0]][3]
							EndIf

							MAJ( _
									StringSplit($txt, ";")[9], _ ;Dossier
									Stringleft(StringSplit($txt, ";")[1], 19), _ ; Date-Heure
									$nom, _ ; Nom
									$prenom, _ ; Prenom
									$serv, _ ; Service
									$ini_ana_enr[$i], _ ; Examen
									"0", _ ; Etat
									Stringleft(StringSplit($txt, ";")[1], 19)) ; MAJ
							$decod = True
						EndIf
					Next
				EndIf

				if $decod = False Then
					for $i = 1 to $ini_ana_enr2[0] ; INTERCEPTION MODIFICATION ANALYSE SUR ECRAN SAISIE PAR DOSSIER (CERBA)
						if StringSplit($txt, ";")[10] = "Analyse = " & $ini_ana_enr2[$i] and StringSplit($txt, ";")[2] = "MODIFICATION" and StringSplit($txt, ";")[6] = "RESULDOS" Then ;Création en gestion de dossier
							$ret = _ArrayFindAll($adoscli, StringSplit($txt, ";")[9], 0, 0, 0, 0, 0, 0)
							if UBound($ret) > 2 Then
								msgbox(0, "Erreur sur le dossier " & $txt1)
								Exit
							EndIf
							$nom = ""
							$prenom = ""
							$serv = ""
							if UBound($ret) = 1 Then
								$nom = $adoscli[$ret[0]][1]
								$prenom = $adoscli[$ret[0]][2]
								$serv = $adoscli[$ret[0]][3]
							EndIf

							MAJ_CERBA( _
									StringSplit($txt, ";")[9], _ ;Dossier
									Stringleft(StringSplit($txt, ";")[1], 19), _ ; Date-Heure
									$nom, _ ; Nom
									$prenom, _ ; Prenom
									$serv, _ ; Service
									$ini_ana_enr2[$i], _ ; Examen
									"0", _ ; Etat
									Stringleft(StringSplit($txt, ";")[1], 19)) ; MAJ
							$decod = True
						EndIf
					Next
				EndIf

				if $decod = False Then ; INTERCEPTION RESULTAT SUR ECRAN GESTION DES DOSSIERS
					if Stringleft(StringSplit($txt, ";")[10], 11) = "*BACT 01" and StringSplit($txt, ";")[2] = "AJOUT" and StringSplit($txt, ";")[6] = "DOSSIER" Then ;Type d'examen
						MAJ( _
								StringSplit($txt, ";")[9], _ ;Dossier
								"", _ ; Date-Heure
								"", _ ; Nom
								"", _ ; Prenom
								"", _ ; Service
								StringSplit($txt, ";")[11], _ ; Examen
								1, _ ; Etat
								Stringleft(StringSplit($txt, ";")[1], 19)) ; MAJ
						$decod = True
					EndIf
				EndIf

				if $decod = False Then ; INTERCEPTION RESULTAT SUR ECRAN SAISIE PAR DOSSIER
					if Stringleft(StringSplit($txt, ";")[10], 11) = "*BACT 01" and StringSplit($txt, ";")[2] = "AJOUT" and StringSplit($txt, ";")[6] = "RESULDOS" Then ;Type d'examen
						MAJ( _
								StringSplit($txt, ";")[9], _ ;Dossier
								"", _ ; Date-Heure
								"", _ ; Nom
								"", _ ; Prenom
								"", _ ; Service
								StringSplit($txt, ";")[11], _ ; Examen
								1, _ ; Etat
								Stringleft(StringSplit($txt, ";")[1], 19)) ; MAJ
						$decod = True
					EndIf
				EndIf

				if $decod = False Then ; INTERCEPTION EDITION POUR SUPPRESSION DE LA LISTE
					for $i = 1 to $ini_ana_edi[0]
						if StringSplit($txt, ";")[10] = "Analyse = " & $ini_ana_edi[$i] and StringSplit($txt, ";")[6] = "EDITRES" Then
							$ret = _ArrayFindAll($encours, StringSplit($txt, ";")[9], 0, 0, 0, 0, 0, 0)
							if UBound($ret) > 2 Then
								msgbox(0, "Erreur sur le dossier " & $txt1)
							elseif UBound($ret) = 1 then
								_ArrayDelete($encours, $ret[0])
							Else
							EndIf
						EndIf
					Next
				EndIf

;~ if $decod=False Then ; INTERCEPTION EDITION POUR SUPPRESSION DE LA LISTE (CERBA)
;~       for $i=1 to $ini_ana_edi2[0]
;~ 	  						   if StringSplit($txt,";")[10]="Analyse = "&$ini_ana_edi2[$i] and StringSplit($txt,";")[6]="EDITRES" Then
;~ $ret=_ArrayFindAll($encoursCERBA, StringSplit($txt,";")[9], 0, 0, 0, 0, 0, 0)
;~ if UBound($ret)>2 Then
;~    msgbox(0,"Erreur sur le dossier "&$txt1)
;~ elseif UBound($ret)=1 then
;~    _ArrayDelete ($encoursCERBA, $ret[0] )
;~    Else
;~ EndIf
;~ EndIf
;~ Next
;~ EndIf



			EndIf
		Endif
	WEnd

	_FileWriteFromArray(@ScriptDir & "\encours.txt", $encours, 0)
	_FileWriteFromArray(StringReplace(@ScriptDir & "\encoursCERBA.txt", "scanbac", "cerba"), $encoursCERBA, 0)
Next









Func MAJ($txt1 = "", $txt2 = "", $txt3 = "", $txt4 = "", $txt5 = "", $txt6 = "", $txt7 = "", $txt8 = "")

	$ret = _ArrayFindAll($encours, $txt1, 0, 0, 0, 0, 0, 0)
	if UBound($ret) > 2 Then
		msgbox(0, "Erreur sur le dossier " & $txt1)
		Exit
	EndIf
	if UBound($ret) = 1 Then ;Dossier déja présent dans la table

		$dat1 = StringMid($encours[$ret[0]][7], 7, 4) & "/" & StringMid($encours[$ret[0]][7], 4, 2) & "/" & StringMid($encours[$ret[0]][7], 1, 2) & StringMid($encours[$ret[0]][7], 11) ;Dat1 = date dans la bdd
		$dat2 = StringMid($txt8, 7, 4) & "/" & StringMid($txt8, 4, 2) & "/" & StringMid($txt8, 1, 2) & StringMid($txt8, 11) ;Date de la fonction
;~       _ArrayDisplay($ret)
		if _DateDiff("s", $dat1, $dat2) > 0 Then ; Maj à faire
			if $txt1 <> "" Then $encours[$ret[0]][0] = $txt1
			if $txt2 <> "" Then $encours[$ret[0]][1] = $txt2
			if $txt3 <> "" Then $encours[$ret[0]][2] = $txt3
			if $txt4 <> "" Then $encours[$ret[0]][3] = $txt4
			if $txt5 <> "" Then $encours[$ret[0]][4] = $txt5
			if $txt6 <> "" Then $encours[$ret[0]][5] = $txt6
			if $txt7 <> "" Then
				if $txt7 > $encours[$ret[0]][6] Then $encours[$ret[0]][6] = $txt7
			EndIf
			if $txt8 <> "" Then $encours[$ret[0]][7] = $txt8
		EndIf
	EndIf

	if UBound($ret) = 0 then ;Ajouter le dossier dans la table
;~    ConsoleWrite($txt1&"|"&$txt2&"|"&$txt3&"|"&$txt4&"|"&$txt5&"|"&$txt6&"|"&$txt7&"|"&$txt8&@crlf)
		Local $add[1][8]
		$add[0][0] = $txt1
		$add[0][1] = $txt2
		$add[0][2] = $txt3
		$add[0][3] = $txt4
		$add[0][4] = $txt5
		$add[0][5] = $txt6
		$add[0][6] = $txt7
		$add[0][7] = $txt8
		_ArrayAdd($encours, $add, 0)
	EndIf
EndFunc   ;==>MAJ



Func MAJ_CERBA($txt1 = "", $txt2 = "", $txt3 = "", $txt4 = "", $txt5 = "", $txt6 = "", $txt7 = "", $txt8 = "")

	$ret = _ArrayFindAll($encoursCERBA, $txt1, 0, 0, 0, 0, 0, 0)
	if UBound($ret) > 2 Then
		msgbox(0, "Erreur sur le dossier " & $txt1)
		Exit
	EndIf
	if UBound($ret) = 1 Then ;Dossier déja présent dans la table

		$dat1 = StringMid($encoursCERBA[$ret[0]][7], 7, 4) & "/" & StringMid($encoursCERBA[$ret[0]][7], 4, 2) & "/" & StringMid($encoursCERBA[$ret[0]][7], 1, 2) & StringMid($encoursCERBA[$ret[0]][7], 11) ;Dat1 = date dans la bdd
		$dat2 = StringMid($txt8, 7, 4) & "/" & StringMid($txt8, 4, 2) & "/" & StringMid($txt8, 1, 2) & StringMid($txt8, 11) ;Date de la fonction
;~       _ArrayDisplay($ret)
		if _DateDiff("s", $dat1, $dat2) > 0 Then ; Maj à faire
			if $txt1 <> "" Then $encoursCERBA[$ret[0]][0] = $txt1
			if $txt2 <> "" Then $encoursCERBA[$ret[0]][1] = $txt2
			if $txt3 <> "" Then $encoursCERBA[$ret[0]][2] = $txt3
			if $txt4 <> "" Then $encoursCERBA[$ret[0]][3] = $txt4
			if $txt5 <> "" Then $encoursCERBA[$ret[0]][4] = $txt5
			if $txt6 <> "" Then $encoursCERBA[$ret[0]][5] = $txt6
			if $txt7 <> "" Then
				if $txt7 > $encoursCERBA[$ret[0]][6] Then $encoursCERBA[$ret[0]][6] = $txt7
			EndIf
			if $txt8 <> "" Then $encoursCERBA[$ret[0]][7] = $txt8
		EndIf
	EndIf

	if UBound($ret) = 0 then ;Ajouter le dossier dans la table
;~    ConsoleWrite($txt1&"|"&$txt2&"|"&$txt3&"|"&$txt4&"|"&$txt5&"|"&$txt6&"|"&$txt7&"|"&$txt8&@crlf)
		Local $add[1][8]
		$add[0][0] = $txt1
		$add[0][1] = $txt2
		$add[0][2] = $txt3
		$add[0][3] = $txt4
		$add[0][4] = $txt5
		$add[0][5] = $txt6
		$add[0][6] = $txt7
		$add[0][7] = $txt8
		_ArrayAdd($encoursCERBA, $add, 0)
	EndIf
EndFunc   ;==>MAJ_CERBA

