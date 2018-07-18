#include <WinAPIFiles.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <Date.au3>
#include <File.au3>
#include <Math.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiImageList.au3>
#include <ColorConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <Misc.au3>
#include "ExtMsgBox.au3"




If _Singleton("apicrypt", 1) = 0 Then Exit

;~ CONFIGURATION
$apitunnel = "Y:\Apitunnel"
$dos = "Y:\SCAN2"

Local $hBitmap, $hImage, $sCLSID, $tData, $tParams
Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
$id = -1

$aFiles = _FileListToArray($apitunnel, "*.pdf", 1, 1)
if UBound($afiles) <> 0 Then

	#Region ### START Koda GUI section ### Form=
	$Form11 = GUICreate("SCANBAC APICRYPT", 624, 135, 192, 124)
	$Label11 = GUICtrlCreateLabel("Recupération des PDF en cours...", 8, 8, 192, 45)
	$Progress11 = GUICtrlCreateProgress(8, 64, 601, 25)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###
	GUICtrlSetData($Progress11, 0)


;~ SUPPRIMER LES FICHIERS TEMPORAIRES
	DirRemove($apitunnel & "\tmp\", 1)
	DirCreate($apitunnel & "\tmp\")

;~ DEPLACER LES FICHIERS RES
	FileMove($apitunnel & "\*.res", $apitunnel & "\BACKUP_RES\", 1)


	For $i = 1 to $afiles[0]
		GUICtrlSetData($Progress11, $i / $afiles[0] * 100)
		GUICtrlSetData($Label11, "Recupération des PDF en cours..." & @CRLF & $afiles[$i])
		Local $aPathSplit = _PathSplit($afiles[$i], $sDrive, $sDir, $sFileName, $sExtension)
;~    RunWait(@comspec & ' /c pdftotext.exe ' & '"' & $afiles[$i] & '" ', '.', @SW_HIDE)
;~ ConsoleWrite("pdftotext.exe "&$afiles[$i])
		RunWait("pdftotext.exe " & $afiles[$i], @ScriptDir, @SW_HIDE)
		Local $hFileOpen = FileOpen(StringReplace($afiles[$i], ".pdf", ".txt"), $FO_READ)
		$yy = _FileCountLines(StringReplace($afiles[$i], ".pdf", ".txt"))
		$ref = ""
		for $y = 1 to $yy
			$a = FileReadLine($hFileOpen)
			if StringLeft($a, 13) = "Info sur CR :" Then
				$a = StringReplace($a, "R2F", "REF")
				$ref = StringRegExpReplace($a, "\D", "")
				ExitLoop
			EndIf
		Next
		FileClose($hFileOpen)
		FileDelete(StringReplace($afiles[$i], ".pdf", ".txt"))
		if $ref = "" Then $ref = $aPathSplit[3]
		DirCreate($apitunnel & '\tmp\' & $aPathSplit[3] & '\')
		RunWait("pdftopng.exe " & $apitunnel & "\" & $aPathSplit[3] & ".pdf " & $apitunnel & "\tmp\" & $aPathSplit[3] & "\" & $aPathSplit[3] & " -gray", @ScriptDir, @SW_HIDE)
		$aFiles2 = _FileListToArray($apitunnel & "\tmp\" & $aPathSplit[3], "*.png", 1, 1)
		_GDIPlus_Startup()
		if UBound($afiles2) <> 0 Then
			For $y = 1 to $afiles2[0]
				$hImage = _GDIPlus_ImageLoadFromFile($afiles2[$y])
				$sCLSID = _GDIPlus_EncodersGetCLSID("JPG")
				_GDIPlus_ImageSaveToFileEx($hImage, $dos & "\" & $ref & "-" & $y & ".jpg", $sCLSID, DllStructGetPtr($tParams))
				DirRemove($apitunnel & "\" & $aPathSplit[3], 1)
			Next
			_GDIPlus_ShutDown()
			FileMove($afiles[$i], $apitunnel & "\BACKUP_PDF", 1)

			Local $hFileOpen = FileOpen("log.txt", $FO_APPEND)
			FileWrite($hFileOpen, _now() & " : PDF TRAITE : " & $afiles[$i] & " (REF " & $ref & ")" & @CRLF)
			FileClose($hFileOpen)
			if stringlen($ref) = 11 then
				Local $hFileOpen = FileOpen("ref.txt", $FO_APPEND)
				FileWrite($hFileOpen, StringRegExpReplace($aPathSplit[3], "\D", "") & "|" & $ref & @CRLF)
				FileClose($hFileOpen)
			EndIf
		EndIf
	Next
	GUIDelete($Form11)
EndIf


#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Attribution via Apicrypt", 880, 658, 400, 21)
$ListView1 = GUICtrlCreateListView("Dossier|Date-Heure|Nom|Prénom|Examen|UF|Etat|Maj", 8, 8, 425, 593, $LVS_SINGLESEL, $WS_EX_CLIENTEDGE)
_GUICtrlListView_SetExtendedListViewStyle($ListView1, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES))
GUICtrlSetBkColor($ListView1, $GUI_BKCOLOR_LV_ALTERNATE)
$Pic1 = GUICtrlCreatePic("", 440, 8, 430, 593)
$Button1 = GUICtrlCreateButton("Précédent", 440, 608, 81, 41)
$Button2 = GUICtrlCreateButton("Suivant", 528, 608, 73, 41)
$Button3 = GUICtrlCreateButton("Attribuer", 672, 608, 81, 41)
$Button4 = GUICtrlCreateButton("Supprimer", 776, 616, 73, 25)
$Button5 = GUICtrlCreateButton("Imprim Page", 8, 608, 105, 41)
$Button6 = GUICtrlCreateButton("Quitter", 120, 608, 97, 41)
$CheckBox = GUICtrlCreateCheckbox("Mode Fax", 230, 615)
$Button7 = GUICtrlCreateButton("Supr. Dos.", 375, 608, 60, 41)
$Button8 = GUICtrlCreateButton("Créa. Dos.", 305, 608, 60, 41)
$LabelA = GUICtrlCreateLabel("2AAMMJJDDDD", 440, 292, 430, 13, $SS_CENTER)
#EndRegion ### END Koda GUI section ###
GUICtrlSetState($LabelA, $GUI_HIDE)
GUISetBkColor($COLOR_AQUA, $Form1)
GUICtrlSetState($Button1, $GUI_DISABLE)
GUICtrlSetState($Button2, $GUI_DISABLE)
$hImage = _GUIImageList_Create()
_GUIImageList_Add($hImage, _GUICtrlListView_CreateSolidBitMap($ListView1, 0xFF0000, 16, 16))
_GUIImageList_Add($hImage, _GUICtrlListView_CreateSolidBitMap($ListView1, 0xFFA500, 16, 16))
_GUIImageList_Add($hImage, _GUICtrlListView_CreateSolidBitMap($ListView1, 0x008000, 16, 16))
_GUICtrlListView_SetImageList($ListView1, $hImage, 1)

;LECTURE FICHIER SACNNERS (DOUBLONS)
;~ FileCopy("Y:\Biowin\Scanner.fic", @ScriptDir & "\Scanner.fic", 1)
;~ FileCopy("Y:\Biowin\Scanner.ndx", @ScriptDir & "\Scanner.ndx", 1)
;~ Runwait(@ScriptDir & "\hyperfile2xml.exe Scanner.fic Scanner.xml")
;~ Local $hFileOpen = FileOpen(@ScriptDir & "\Scanner.xml", $FO_READ)
;~ $lign = _FileCountLines(@ScriptDir & "\scanner.xml")
;~ FileReadLine($hFileOpen)
;~ FileReadLine($hFileOpen)
;~ Local $aSCAN[(($lign - 3) / 11) - 1][8]
;~ for $i = 0 to UBound($aSCAN, 1) - 1
;~ 	FileReadLine($hFileOpen)
;~ 	$aSCAN[$i][0] = StringStripWS(StringSplit(StringSplit(FileReadLine($hFileOpen), ">")[2], "<")[1], 8) & StringStripWS(StringSplit(StringSplit(FileReadLine($hFileOpen), ">")[2], "<")[1], 8)
;~ 	for $j = 2 to 8
;~ 		$aSCAN[$i][$j - 1] = StringStripWS(StringSplit(StringSplit(FileReadLine($hFileOpen), ">")[2], "<")[1], 8)
;~ 	Next
;~ 	FileReadLine($hFileOpen)
;~ Next
;~ FileClose($hFileOpen)

;LECTURE FICHIERS EN COURS
Global $encours, $adelai
_FileReadToArray(@ScriptDir & "\encours.txt", $encours, 0, "|")
_FileReadToArray(@ScriptDir & "\delais.txt", $adelai, 0, "|")
_ArraySort($encours, 0, 0, 0, 0)
for $y = 0 To UBound($encours) - 2
	GUICtrlCreateListViewItem(StringLeft($encours[$y][0], 7) & " " & StringRight($encours[$y][0], 4) & "|" & $encours[$y][1] & "|" & $encours[$y][2] & "|" & $encours[$y][3] & "|" & $encours[$y][5] & "|" & $encours[$y][4] & "|" & $encours[$y][6] & "|" & $encours[$y][7], $ListView1)
	_GUICtrlListView_SetColumnWidth($ListView1, 0, $LVSCW_AUTOSIZE)
	GUICtrlSetBkColor(-1, 0xD3D3D3)
	GUICtrlSetColor(-1, "0x000000")
	$aret = _ArrayFindAll($encours, $encours[$y][2], 0, 0, 0, 0, 2)
	if UBound($aret, 1) > 1 Then
		GUICtrlSetColor(-1, "0x0000FF")
	EndIf
	if UBound($aret, 1) = 3 Then ; NE PAS COLORIER LES 3 HEMOCS QUI SE SUIVENT
		If $encours[$aret[0]][0] + 1 = $encours[$aret[1]][0] AND $encours[$aret[2]][0] - 1 = $encours[$aret[1]][0] Then GUICtrlSetColor(-1, "0x000000")
	EndIf
	if UBound($aret, 1) = 2 Then ; NE PAS COLORIER LES 2 HEMOCS QUI SE SUIVENT
		If $encours[$aret[0]][0] + 1 = $encours[$aret[1]][0] Then GUICtrlSetColor(-1, "0x000000")
	EndIf

	$dat2 = StringMid($encours[$y][1], 7, 4) & "/" & StringMid($encours[$y][1], 4, 2) & "/" & StringMid($encours[$y][1], 1, 2)
	$delai1 = _DateDiff("d", $dat2, stringleft(_nowcalc(), 10))
	$delai2 = 0
	for $t = 0 to UBound($adelai) - 1
		if $encours[$y][5] = $adelai[$t][0] Then
			$delai2 = $adelai[$t][1]
		EndIf
	Next
	if $delai1 > $delai2 Then
		_GUICtrlListView_SetItemImage($ListView1, $y, 0)
	Else
		_GUICtrlListView_SetItemImage($ListView1, $y, 2)
	EndIf
Next
_GUICtrlListView_HideColumn($ListView1, 1)
_GUICtrlListView_HideColumn($ListView1, 6)
_GUICtrlListView_HideColumn($ListView1, 7)


GUICtrlSetState($CheckBox, $GUI_CHECKED)
GUISetBkColor($COLOR_RED, $Form1)

;DEBUT BOUCLE GUI
$file = 1
GUISetState(@SW_SHOW)
while 1
	GUICtrlSetState($Button1, $GUI_DISABLE)
	GUICtrlSetState($Button2, $GUI_DISABLE)
	GUICtrlSetState($Button3, $GUI_DISABLE)
	GUICtrlSetState($Button4, $GUI_DISABLE)
	GUICtrlSetState($Button5, $GUI_DISABLE)
	$aJPG = _FileListToArray("Y:\SCAN2\", "*.jpg", 1, 1)
	if @error = 4 or @error = 1 Then $file = 0

 if $file <> 0 then $file = _Min(UBound($aJPG)-1, $file)
;~ 	ConsoleWrite($file&@crlf)
	if $file > 1 Then GUICtrlSetState($Button1, $GUI_ENABLE)
	if $file < UBound($aJPG)-1 then GUICtrlSetState($Button2, $GUI_ENABLE)


	if $file <> 0 Then
		GUICtrlSetImage($Pic1, $aJPG[$file])
		GUICtrlSetState($Button3, $GUI_ENABLE)
		GUICtrlSetState($Button4, $GUI_ENABLE)
		GUICtrlSetState($Button5, $GUI_ENABLE)
	Else
		GUICtrlSetImage($Pic1, "")
	EndIf


;~ RECONNAISSANCE AUTOMATIQUE DU DOSSIER
	if $file <> 0 Then
		GUICtrlSetState($LabelA, $GUI_HIDE)
		if $id <> -1 then _GUICtrlListView_SetItemSelected($ListView1, $id, 0, 0)
		$id = -1

		Local $aPathSplit = _PathSplit($aJPG[$file], $sDrive, $sDir, $sFileName, $sExtension)
		if stringlen($aPathSplit[3]) = 13 Then
			$id = _ArraySearch($encours, stringleft($aPathSplit[3], 11), 0, 0, 0, 0, 1, 0)
			if $id <> -1 Then
;~ 	  ConsoleWrite($id)


				GUICtrlSetData($LabelA, $aPathSplit[3])
				GUICtrlSetState($LabelA, $GUI_SHOW)
				_GUICtrlListView_SetItemSelected($ListView1, $id)
				_WinAPI_SetFocus(ControlGetHandle("Attribution via Apicrypt", "", $ListView1))
			EndIf
		EndIf
	EndIf

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				exit
			Case $Button6
				Exit
			Case $Button1
				$file = $file - 1
				ExitLoop(1)
			Case $Button2
				$file = $file + 1
				ExitLoop(1)

			Case $Button5 ;IMPRESSION
				GUISetState(@SW_DISABLE, $Form1)
				SplashTextOn("Impression en cours...", "Veuillez patienter..." & @CRLF & "Impression en cours...", 200, 70, -1, -1, 0, "", 12)
				if @ComputerName = "CHA-PLABO2" Then
					RunWait('SumatraPDF.exe ' & $aJPG[$file] & ' -print-to "\\cha-vmsimp1\cha-ilab1"')
				Else

					RunWait('SumatraPDF.exe ' & $aJPG[$file] & ' -print-to "Lexmark T652 (MS)"')
				EndIf
				SplashOff()
				GUISetState(@SW_ENABLE, $Form1)
				Local $hFileOpen2 = FileOpen("log_imprim.txt", $FO_APPEND)
				FileWrite($hFileOpen2, _now() & " : IMPRESSION FICHIER : " & $aJPG[$file] & " sur imprimante Lexmark T652 (MS)" & @CRLF)

				FileClose($hFileOpen2)

			Case $CheckBox
				If _IsChecked($CheckBox) Then
					GUISetBkColor($COLOR_RED, $Form1)
				Else
					GUISetBkColor($COLOR_AQUA, $Form1)
				EndIf

			case $Pic1
				if $file > 0 Then
					$hGUI = GUICreate("", @DesktopWidth, @DesktopHeight, -1, -1, $WS_POPUP, $WS_EX_TOPMOST)
					$Pic2 = GUICtrlCreatePic($aJPG[$file], 0, 0, @DesktopWidth, @DesktopHeight)
					GUISetState(@SW_SHOW, $hGUI)
					While 1
						if GUIGetMsg() = $Pic2 Then ExitLoop
					WEnd
					GUIDelete($hGUI)
				EndIf

			Case $Button5 ;ACTUALISATION
				ExitLoop


			Case $Button4 ;SUPPRIMER
				if MsgBox(1, "Supprimer ?", "Etes-vous sur de vouloir suprimer cet élément ?") = 1 Then
					FileRecycle($aJPG[$file])
					ExitLoop
				EndIf

			Case $Button3 ;ATTRIBUER
				$selection = GUICtrlRead($ListView1)
				If $selection <> 0 Then
					$index = ControlListView("Attribution via Apicrypt", "", $ListView1, "GetSelected")
					$item = ControlListView("Attribution via Apicrypt", "", $ListView1, "GetText", $index)
					$item = StringLeft($item, 7) & StringRight($item, 4)
					$aSCAN2 = _FileListToArray("Y:\SCANNER\", stringright($item, 10) & "-*.jpg", 1)
					$action = ""

					if ubound($aSCAN2) > 1 Then
						GUICtrlSetState($Button3, $GUI_DISABLE)
						GUICtrlSetState($Button4, $GUI_DISABLE)

						#Region ### START Koda GUI section ### Form=
						$Form2 = GUICreate("Fichiers déjà présents sur BioWin", 747, 376, 560, 265, 0,$WS_EX_TOPMOST)
						$Button21 = GUICtrlCreateButton("Ajouter les fichiers au dossier", 561, 315, 179, 33)
						$Button22 = GUICtrlCreateButton("Annuler", 10, 315, 75, 33)
						$Button23 = GUICtrlCreateButton("Fichiers déjà scannés (fax). Ne pas ajouter au dossier (juste saisir la date de retour)", 105, 315, 435, 33)
						If _IsChecked($CheckBox) = True Then GUICtrlSetState($Button23, $GUI_DISABLE)
						$Label21 = GUICtrlCreateLabel("Un ou plusieurs fichiers sont déjà présents sur BioWin. Voulez-vous continuer ?", 104, 8, 547, 20)
						GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
						GUICtrlSetColor(-1, 0xFF0000)
						$pic21 = 9999
						$pic22 = 9999
						$pic23 = 9999
						$pic24 = 9999
						if ubound($aSCAN2) > 1 then $pic21 = GUICtrlCreatePic("Y:\SCANNER\" & $aSCAN2[1], 8, 32, 177, 273)
						if ubound($aSCAN2) > 2 then $pic22 = GUICtrlCreatePic("Y:\SCANNER\" & $aSCAN2[2], 192, 32, 177, 273)
						if ubound($aSCAN2) > 3 then $pic23 = GUICtrlCreatePic("Y:\SCANNER\" & $aSCAN2[3], 376, 32, 177, 273)
						if ubound($aSCAN2) > 4 then $pic24 = GUICtrlCreatePic("Y:\SCANNER\" & $aSCAN2[4], 560, 32, 177, 273)
						GUISetState(@SW_SHOW)
						#EndRegion ### END Koda GUI section ###

						While 1
							$nMsg = GUIGetMsg()
							Switch $nMsg

								Case $pic21 ;AFFICHAGE PLEIN ECRAN IMAGE 1
									$hGUI = GUICreate("", @DesktopWidth, @DesktopHeight, -1, -1, $WS_POPUP, $WS_EX_TOPMOST)
									$Pic2 = GUICtrlCreatePic("Y:\SCANNER\" & $aSCAN2[1], 0, 0, @DesktopWidth, @DesktopHeight)
									GUISetState(@SW_SHOW, $hGUI)
									While 1
										if GUIGetMsg() = $Pic2 Then ExitLoop
									WEnd
									GUIDelete($hGUI)

								Case $pic22 ;AFFICHAGE PLEIN ECRAN IMAGE 2
									$hGUI = GUICreate("", @DesktopWidth, @DesktopHeight, -1, -1, $WS_POPUP, $WS_EX_TOPMOST)
									$Pic2 = GUICtrlCreatePic("Y:\SCANNER\" & $aSCAN2[2], 0, 0, @DesktopWidth, @DesktopHeight)
									GUISetState(@SW_SHOW, $hGUI)
									While 1
										if GUIGetMsg() = $Pic2 Then ExitLoop
									WEnd
									GUIDelete($hGUI)

								Case $pic23 ;AFFICHAGE PLEIN ECRAN IMAGE 3
									$hGUI = GUICreate("", @DesktopWidth, @DesktopHeight, -1, -1, $WS_POPUP, $WS_EX_TOPMOST)
									$Pic2 = GUICtrlCreatePic("Y:\SCANNER\" & $aSCAN2[3], 0, 0, @DesktopWidth, @DesktopHeight)
									GUISetState(@SW_SHOW, $hGUI)
									While 1
										if GUIGetMsg() = $Pic2 Then ExitLoop
									WEnd
									GUIDelete($hGUI)

								Case $pic24 ;AFFICHAGE PLEIN ECRAN IMAGE 4
									$hGUI = GUICreate("", @DesktopWidth, @DesktopHeight, -1, -1, $WS_POPUP, $WS_EX_TOPMOST)
									$Pic2 = GUICtrlCreatePic("Y:\SCANNER\" & $aSCAN2[4], 0, 0, @DesktopWidth, @DesktopHeight)
									GUISetState(@SW_SHOW, $hGUI)
									While 1
										if GUIGetMsg() = $Pic2 Then ExitLoop
									WEnd
									GUIDelete($hGUI)

								Case $Button23 ;NE SAISIR QUE LA DATE DE RETOUR
									$action = "R"
									GUIDelete($Form2)
									exitloop

								Case $Button22 ;ANNULER
									GUIDelete($Form2)
									exitloop(2)

								Case $Button21 ;AJOUTER LES FICHIERS AU DOSSIER
									If _IsChecked($CheckBox) Then
										$action = "F"
									Else
										$action = "S"
									EndIf
									GUIDelete($Form2)
									GUICtrlSetState($Button3, $GUI_ENABLE)
									GUICtrlSetState($Button4, $GUI_ENABLE)
									exitloop
							EndSwitch
						WEnd

						$nfich = "Y:\SCANNER\" & stringright($item, 10) & "-" & UBound($aSCAN2) - 1 & ".jpg"

					Else
						If _IsChecked($CheckBox) Then
							$action = "F"
						Else
							$action = "S"
						EndIf
						$nfich = "Y:\SCANNER\" & stringright($item, 10) & "-0.jpg"
					EndIf
					if $action = "S" or $action = "F" Then
						$retour = FileMove($aJPG[$file], $nfich)
						FileSetTime($nfich, "")
						if $retour = 0 Then
							msgbox(0, "", "Erreur lors du déplacement du JPEG")
							Exit
						EndIf
					Else

						FileRecycle($aJPG[$file])
					EndIf
					if _FileCountLines("integration.txt") > 0 Then

						Local $hFileOpen = FileOpen("integration.txt", $FO_APPEND)
						FileWrite($hFileOpen, $nfich & "|" & $action & @CRLF)
						FileClose($hFileOpen)
					Else
						Local $hFileOpen = FileOpen("integration.txt", $FO_OVERWRITE)
						FileWrite($hFileOpen, $nfich & "|" & $action & @CRLF)
						FileClose($hFileOpen)
					EndIf

					ExitLoop
				EndIf

			Case $Button7 ;SUPRIMER DOSSIER
				$selection = GUICtrlRead($ListView1)
				If $selection <> 0 Then
					$index = ControlListView("Attribution via Apicrypt", "", $ListView1, "GetSelected")

					if MsgBox(4, "Confirmation de la suppression", "Etes-vous sur de vouloir supprimer le dossier suivant : " & @crlf & "n°" & $encours[$index][0] & " " & $encours[$index][2] & " " & $encours[$index][3] & " (" & $encours[$index][5] & ") ?") = 6 Then
						_ArrayDelete($encours, $index)
						_ArraySort($encours, 1, 0, 0, 0)
						_FileWriteFromArray(@ScriptDir & "\encours.txt", $encours)
						_ArraySort($encours, 0, 0, 0, 0)
						_GUICtrlListView_DeleteAllItems($ListView1)
						for $y = 0 To UBound($encours) - 2
							GUICtrlCreateListViewItem(StringLeft($encours[$y][0], 7) & " " & StringRight($encours[$y][0], 4) & "|" & $encours[$y][1] & "|" & $encours[$y][2] & "|" & $encours[$y][3] & "|" & $encours[$y][5] & "|" & $encours[$y][4] & "|" & $encours[$y][6] & "|" & $encours[$y][7], $ListView1)
							_GUICtrlListView_SetColumnWidth($ListView1, 0, $LVSCW_AUTOSIZE)
							GUICtrlSetBkColor(-1, 0xD3D3D3)
							GUICtrlSetColor(-1, "0x000000")
							$aret = _ArrayFindAll($encours, $encours[$y][2], 0, 0, 0, 0, 2)
							if UBound($aret, 1) > 1 Then
								GUICtrlSetColor(-1, "0x0000FF")
							EndIf
							if UBound($aret, 1) = 3 Then ; NE PAS COLORIER LES 3 HEMOCS QUI SE SUIVENT
								If $encours[$aret[0]][0] + 1 = $encours[$aret[1]][0] AND $encours[$aret[2]][0] - 1 = $encours[$aret[1]][0] Then GUICtrlSetColor(-1, "0x000000")
							EndIf
							if UBound($aret, 1) = 2 Then ; NE PAS COLORIER LES 2 HEMOCS QUI SE SUIVENT
								If $encours[$aret[0]][0] + 1 = $encours[$aret[1]][0] Then GUICtrlSetColor(-1, "0x000000")
							EndIf

							$dat2 = StringMid($encours[$y][1], 7, 4) & "/" & StringMid($encours[$y][1], 4, 2) & "/" & StringMid($encours[$y][1], 1, 2)
							$delai1 = _DateDiff("d", $dat2, stringleft(_nowcalc(), 10))
							$delai2 = 0
							for $t = 0 to UBound($adelai) - 1
								if $encours[$y][5] = $adelai[$t][0] Then
									$delai2 = $adelai[$t][1]
								EndIf
							Next
							if $delai1 > $delai2 Then
								_GUICtrlListView_SetItemImage($ListView1, $y, 0)
							Else
								_GUICtrlListView_SetItemImage($ListView1, $y, 2)
							EndIf
						Next
						_GUICtrlListView_HideColumn($ListView1, 1)
						_GUICtrlListView_HideColumn($ListView1, 6)
						_GUICtrlListView_HideColumn($ListView1, 7)
					EndIf
				EndIf

			Case $Button8 ;AJOUTER DOSSIER

				#Region ### START Koda GUI section ### Form=
				$Form10 = GUICreate("Ajouter un dossier", 298, 168, 192, 134, 0,$WS_EX_TOPMOST)
				$Label101 = GUICtrlCreateLabel("Dossier (2AAMMJJDDDD) : ", 8, 8, 161, 17)
				GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
				$Input101 = GUICtrlCreateInput("", 168, 8, 121, 21)
				$Label102 = GUICtrlCreateLabel("Nom marital : ", 8, 32, 82, 17)
				GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
				$Label103 = GUICtrlCreateLabel("Prénom : ", 8, 56, 58, 17)
				GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
				$Label104 = GUICtrlCreateLabel("Analyse : ", 8, 80, 60, 17)
				GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
				$Input102 = GUICtrlCreateInput("", 88, 32, 201, 21)
				$Input103 = GUICtrlCreateInput("", 72, 56, 217, 21)
				$Input104 = GUICtrlCreateInput("", 72, 80, 57, 21)
				$Button101 = GUICtrlCreateButton("Ajouter", 0, 104, 113, 33)
				$Button102 = GUICtrlCreateButton("Annuler", 176, 104, 113, 33)
				GUISetState(@SW_SHOW)
				#EndRegion ### END Koda GUI section ###

				While 1
					$nMsg = GUIGetMsg()
					Switch $nMsg
						Case $Button102
							GUIDelete($Form10)
							ExitLoop
						Case $Button101
							if GUICtrlRead($Input101) = "" or GUICtrlRead($Input102) = "" or GUICtrlRead($Input103) = "" or GUICtrlRead($Input104) = "" or stringlen(GUICtrlRead($Input101)) <> 11 Then
								MsgBox(0, "Erreur", "Données non conformes pour l'ajout du dossier !" & @crlf & "Merci de vérifier.")
							Else
								_ArrayAdd($encours, GUICtrlRead($Input101) & "||" & GUICtrlRead($Input102) & "|" & GUICtrlRead($Input103) & "||" & GUICtrlRead($Input104) & "||")
								GUIDelete($Form10)
								ExitLoop
							EndIf
					EndSwitch
				WEnd
				_ArraySort($encours, 1, 0, 0, 0)
				_FileWriteFromArray(@ScriptDir & "\encours.txt", $encours)
				_ArraySort($encours, 0, 0, 0, 0)
				_GUICtrlListView_DeleteAllItems($ListView1)
				for $y = 0 To UBound($encours) - 2
					GUICtrlCreateListViewItem(StringLeft($encours[$y][0], 7) & " " & StringRight($encours[$y][0], 4) & "|" & $encours[$y][1] & "|" & $encours[$y][2] & "|" & $encours[$y][3] & "|" & $encours[$y][5] & "|" & $encours[$y][4] & "|" & $encours[$y][6] & "|" & $encours[$y][7], $ListView1)
					_GUICtrlListView_SetColumnWidth($ListView1, 0, $LVSCW_AUTOSIZE)
					GUICtrlSetBkColor(-1, 0xD3D3D3)
					GUICtrlSetColor(-1, "0x000000")
					$aret = _ArrayFindAll($encours, $encours[$y][2], 0, 0, 0, 0, 2)
					if UBound($aret, 1) > 1 Then
						GUICtrlSetColor(-1, "0x0000FF")
					EndIf
					if UBound($aret, 1) = 3 Then ; NE PAS COLORIER LES 3 HEMOCS QUI SE SUIVENT
						If $encours[$aret[0]][0] + 1 = $encours[$aret[1]][0] AND $encours[$aret[2]][0] - 1 = $encours[$aret[1]][0] Then GUICtrlSetColor(-1, "0x000000")
					EndIf
					if UBound($aret, 1) = 2 Then ; NE PAS COLORIER LES 2 HEMOCS QUI SE SUIVENT
						If $encours[$aret[0]][0] + 1 = $encours[$aret[1]][0] Then GUICtrlSetColor(-1, "0x000000")
					EndIf

					$dat2 = StringMid($encours[$y][1], 7, 4) & "/" & StringMid($encours[$y][1], 4, 2) & "/" & StringMid($encours[$y][1], 1, 2)
					$delai1 = _DateDiff("d", $dat2, stringleft(_nowcalc(), 10))
					$delai2 = 0
					for $t = 0 to UBound($adelai) - 1
						if $encours[$y][5] = $adelai[$t][0] Then
							$delai2 = $adelai[$t][1]
						EndIf
					Next
					if $delai1 > $delai2 Then
						_GUICtrlListView_SetItemImage($ListView1, $y, 0)
					Else
						_GUICtrlListView_SetItemImage($ListView1, $y, 2)
					EndIf
				Next
				_GUICtrlListView_HideColumn($ListView1, 1)
				_GUICtrlListView_HideColumn($ListView1, 6)
				_GUICtrlListView_HideColumn($ListView1, 7)

		EndSwitch
	WEnd
WEnd

Func _IsChecked($idControlID)
	Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked



