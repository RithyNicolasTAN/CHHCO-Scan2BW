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
#include <GDIPlus.au3>
#include <WinAPIGdi.au3>
#include <WinAPISys.au3>
#include "ExtMsgBox.au3"
#include <IE.au3>
#include <WinAPIFiles.au3>

;~ for $i=0 to 10
;~    ProcessClose ( "iexplore.exe" )
;~ Next

;~ Down_PDF_CERBA("1804180094")

;~ Exit




global $arint, $encours, $aJPG, $file
Local $hBitmap, $hImage, $sCLSID, $tData, $tParams, $ar, $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""

If _Singleton("cerba", 1) = 0 Then Exit
SplashTextOn("Chargement de l'interface en cours...", "Chargement de l'interface en cours...", 200, 70, -1, -1, 0, "", 12)

$dos = "\\cha-vmsdata1\Echange\Laboratoire\INTERFACES\cerba\IMG\"
$dos2 = "\\cha-vmsdata1\Echange\Laboratoire\INTERFACES\cerba\PDF\"

;RAJOUT DES 2 DEDVANT SI NECESSAIRE (FONCTIONNE PAS
$aFiles=_FileListToArray ($dos2,"*.pdf",1,0)
;~ _ArrayDisplay($afiles)
if UBound($afiles) <> 0 Then
For $i=1 to $afiles[0]
   if StringLeft($afiles[$i],1)<>2 Then
	  $afiles[$i]="2"&$afiles[$i]
	  FileMove($dos2&$afiles[$i],$dos2&"2"&$afiles[$i],1)
	  EndIf
   Next
   EndIf



$aFiles=_FileListToArray ($dos2,"*.pdf",1,1)
;~ _ArrayDisplay($afiles)
if UBound($afiles) <> 0 Then
For $i=1 to $afiles[0]

;~     For $i=1 to 1
;~    GUICtrlSetData($Progress11, $i/$afiles[0]*100)
;~    GUICtrlSetData($Label11, "Recupération des PDF en cours... "&$i&"/"&$afiles[0] & @CRLF & $afiles[$i])

	  FileRecycle(StringReplace($afiles[$i],".pdf",".txt"))
;~ 	  ConsoleWrite($afiles[$i]&" : ")
	  RunWait("pdftotext.exe "&$afiles[$i],@ScriptDir,@SW_HIDE)
	  $yy=_FileCountLines (StringReplace($afiles[$i],".pdf",".txt"))
	  Local $hFileOpen = FileOpen(StringReplace($afiles[$i],".pdf",".txt"), $FO_READ)
	  $ref=""
	  for $y=1 to $yy

;ESSAI n°1
$a=FileReadLine($hfileopen)
	  IF Stringmid($a,1,1)="1" and stringmid($a,7,1)="0" and StringLen($a)=10 and stringmid($a,1,2)<19 And stringmid($a,3,2)<13 and stringmid($a,5,2)<32 Then
   $ref=$a
   ExitLoop
EndIf

;ESSAI n°2
$a=StringReplace($a,"Vos références : ","")
IF Stringmid($a,1,1)="1" and stringmid($a,7,1)="0" and StringLen($a)=10 and stringmid($a,1,2)<19 And stringmid($a,3,2)<13 and stringmid($a,5,2)<32 Then
   $ref=$a
   ExitLoop
EndIf

;ESSAI n°3
;~ consolewrite($a&@crlf)
  $a=stringsplit($a," ",2)[0]
	    IF Stringmid($a,1,1)="1" and stringmid($a,7,1)="0" and StringLen($a)=10 and stringmid($a,1,2)<19 And stringmid($a,3,2)<13 and stringmid($a,5,2)<32 Then
   $ref=$a
   ExitLoop
EndIf


Next
if $ref<>"" Then FileMove($afiles[$i],$dos2&"2"&$a&".pdf",1)
FileClose($hFileOpen)
FileRecycle(StringReplace($afiles[$i],".pdf",".txt"))

next
EndIf

Local $hBitmap, $hImage, $sCLSID, $tData, $tParams
Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
$id = -1

$aFiles = _FileListToArray($dos2, "*.pdf", 1, 1)
if UBound($afiles) <> 0 Then

	#Region ### START Koda GUI section ### Form=
	$Form11 = GUICreate("SCANBAC APICRYPT", 624, 135, 192, 124)
	$Label11 = GUICtrlCreateLabel("Recupération des PDF en cours...", 8, 8, 192, 45)
	$Progress11 = GUICtrlCreateProgress(8, 64, 601, 25)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###
	GUICtrlSetData($Progress11, 0)


;~ SUPPRIMER LES FICHIERS TEMPORAIRES
	DirRemove($dos & "tmp\", 1)
	DirCreate($dos & "tmp\")
	DirRemove($dos2 & "tmp\", 1)
	DirCreate($dos2 & "tmp\")

;~ DEPLACER LES FICHIERS RES


	For $i = 1 to $afiles[0]
		GUICtrlSetData($Progress11, $i / $afiles[0] * 100)
		GUICtrlSetData($Label11, "Recupération des PDF en cours..." & @CRLF & $afiles[$i])
		Local $aPathSplit = _PathSplit($afiles[$i], $sDrive, $sDir, $sFileName, $sExtension)
		DirCreate($dos2 & '\tmp\' & $aPathSplit[3] & '\')
		RunWait("pdftopng.exe " & $dos2 & "\" & $aPathSplit[3] & ".pdf " & $dos2 & "\tmp\" & $aPathSplit[3] & "\" & $aPathSplit[3] & " -gray", @ScriptDir, @SW_HIDE)
		$aFiles2 = _FileListToArray($dos2 & "\tmp\" & $aPathSplit[3], "*.png", 1, 1)
		_GDIPlus_Startup()
		if UBound($afiles2) <> 0 Then
			For $y = 1 to $afiles2[0]
				$hImage = _GDIPlus_ImageLoadFromFile($afiles2[$y])
				$sCLSID = _GDIPlus_EncodersGetCLSID("JPG")
				_GDIPlus_ImageSaveToFileEx($hImage, $dos & "\" & $aPathSplit[3] & "-" & $y & ".jpg", $sCLSID, DllStructGetPtr($tParams))
				DirRemove($dos2 & "\" & $aPathSplit[3], 1)
			Next
			_GDIPlus_ShutDown()
			FileDelete($afiles[$i])
endif

	Next
	GUIDelete($Form11)
EndIf

;~ exit





_FileReadToArray(@ScriptDir & "\integration.txt", $ar, 0, "|")
;~ _ArrayDisplay($ar)
$id = -1

#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Attribution via Apicrypt", 880, 658, 400, 21)
$ListView1 = GUICtrlCreateListView("Dossier|Date-Heure|Nom|Prénom|Examen|UF|Etat|Maj", 8, 8, 425, 593, $LVS_SINGLESEL, $WS_EX_CLIENTEDGE)
_GUICtrlListView_SetExtendedListViewStyle($ListView1, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES))
GUICtrlSetBkColor($ListView1, $GUI_BKCOLOR_LV_ALTERNATE)
$Pic1 = GUICtrlCreatePic("", 440, 8, 430, 593)
$Button1 = GUICtrlCreateButton("Précédent", 440, 608, 81, 41)
$Button2 = GUICtrlCreateButton("Suivant", 528, 608, 73, 41)
$Button3 = GUICtrlCreateButton("Attribuer", 700, 608, 81, 41)
$Button4 = GUICtrlCreateButton("Supprimer", 800, 616, 73, 25)
$Button5 = GUICtrlCreateButton("Imprim Page", 8, 608, 105, 41)
$Button6 = GUICtrlCreateButton("Quitter", 120, 608, 97, 41)
;~ $CheckBox = GUICtrlCreateCheckbox("Mode Fax", 230, 615)
$Button7 = GUICtrlCreateButton("Supr. Dos.", 375, 608, 60, 41)
$Button8 = GUICtrlCreateButton("Créa. Dos.", 305, 608, 60, 41)
$LabelA = GUICtrlCreateLabel("2AAMMJJDDDD", 440, 292, 430, 13, $SS_CENTER)
;~ $Button9 = GUICtrlCreateButton("Téléch. PDF", 610, 608, 81, 41)
#EndRegion ### END Koda GUI section ###

GUICtrlSetState($LabelA, $GUI_HIDE)
GUISetBkColor($COLOR_RED, $Form1)
GUICtrlSetState($Button1, $GUI_DISABLE)
GUICtrlSetState($Button2, $GUI_DISABLE)
;~ GUICtrlSetState($CheckBox, $GUI_CHECKED)
;~ $hImage = _GUIImageList_Create()
;~ _GUIImageList_Add($hImage, _GUICtrlListView_CreateSolidBitMap($ListView1, 0xFF0000, 16, 16))
;~ _GUIImageList_Add($hImage, _GUICtrlListView_CreateSolidBitMap($ListView1, 0xFFA500, 16, 16))
;~ _GUIImageList_Add($hImage, _GUICtrlListView_CreateSolidBitMap($ListView1, 0x008000, 16, 16))
;~ _GUICtrlListView_SetImageList($ListView1, $hImage, 1)

_FileReadToArray(@ScriptDir & "\encoursCERBA.txt", $encours, 0, "|")
_Rafraichir()

;~ SplashOff()

;DEBUT BOUCLE GUI
$file = 1
GUISetState(@SW_SHOW)
while 1
	_Rafraichir_IMG()

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
				RunWait('SumatraPDF.exe ' & $aJPG[$file] & ' -print-to "Lexmark T652 (MS)"')
				SplashOff()
				GUISetState(@SW_ENABLE, $Form1)
				Local $hFileOpen2 = FileOpen("log_imprim.txt", $FO_APPEND)
				FileWrite($hFileOpen2, _now() & " : IMPRESSION FICHIER : " & $aJPG[$file] & " sur imprimante Lexmark T652 (MS)" & @CRLF)
				FileClose($hFileOpen2)

;~ 			Case $CheckBox
;~ 				If _IsChecked($CheckBox) Then
;~ 					GUISetBkColor($COLOR_RED, $Form1)
;~ 				Else
;~ 					GUISetBkColor($COLOR_AQUA, $Form1)
;~ 				EndIf

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
;~ 						$Button23 = GUICtrlCreateButton("Fichiers déjà scannés (fax). Ne pas ajouter au dossier (juste saisir la date de retour)", 105, 315, 435, 33)
;~ 						If _IsChecked($CheckBox) = True Then GUICtrlSetState($Button23, $GUI_DISABLE)
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

;~ 								Case $Button23 ;NE SAISIR QUE LA DATE DE RETOUR
;~ 									$action = "R"
;~ 									GUIDelete($Form2)
;~ 									exitloop

								Case $Button22 ;ANNULER
									GUIDelete($Form2)
									exitloop(2)

								Case $Button21 ;AJOUTER LES FICHIERS AU DOSSIER
;~ 									If _IsChecked($CheckBox) Then
										$action = "F"
;~ 									Else
;~ 										$action = "S"
;~ 									EndIf
									GUIDelete($Form2)
									GUICtrlSetState($Button3, $GUI_ENABLE)
									GUICtrlSetState($Button4, $GUI_ENABLE)
									exitloop
							EndSwitch
						WEnd

						$nfich = "Y:\SCANNER\" & stringright($item, 10) & "-" & UBound($aSCAN2) - 1 & ".jpg"

					Else
;~ 						If _IsChecked($CheckBox) Then
							$action = "F"
;~ 						Else
;~ 							$action = "S"
;~ 						EndIf
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

						_Rafraichir(1)
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
;~ 				$Label104 = GUICtrlCreateLabel("Analyse : ", 8, 80, 60, 17)
				GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
				$Input102 = GUICtrlCreateInput("", 88, 32, 201, 21)
				$Input103 = GUICtrlCreateInput("", 72, 56, 217, 21)
;~ 				$Input104 = GUICtrlCreateInput("", 72, 80, 57, 21)
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
							if GUICtrlRead($Input101) = "" or GUICtrlRead($Input102) = "" or GUICtrlRead($Input103) = "" or stringlen(GUICtrlRead($Input101)) <> 11 Then
								MsgBox(0, "Erreur", "Données non conformes pour l'ajout du dossier !" & @crlf & "Merci de vérifier.")
							Else
								_ArrayAdd($encours, GUICtrlRead($Input101) & "||" & GUICtrlRead($Input102) & "|" & GUICtrlRead($Input103) & "||*CERBA||")
								GUIDelete($Form10)
								ExitLoop
							EndIf
					EndSwitch
				WEnd
				_Rafraichir(1)

;~ 			Case $Button9 ;TELECHARGEMENT PDF
;~ 				GUISetState(@SW_DISABLE, $Form1)
;~ 				$selection = GUICtrlRead($ListView1)
;~ 				If $selection <> 0 Then
;~ 					$index = ControlListView("Attribution via Apicrypt", "", $ListView1, "GetSelected")
;~ 					$item = ControlListView("Attribution via Apicrypt", "", $ListView1, "GetText", $index)
;~ 					Down_PDF_CERBA(stringright(StringLeft($item, 7) & StringRight($item, 4), 10))
;~ 					$file = 1

;~ 				Endif
;~ 				GUISetState(@SW_ENABLE, $Form1)
;~ 				ExitLoop

				_Rafraichir_IMG()
		EndSwitch
	WEnd
WEnd
;~
Func _Rafraichir_IMG()
;~ RECUPERATION DES IMAGES
	GUICtrlSetState($Button1, $GUI_DISABLE)
	GUICtrlSetState($Button2, $GUI_DISABLE)
	GUICtrlSetState($Button3, $GUI_DISABLE)
	GUICtrlSetState($Button4, $GUI_DISABLE)
	GUICtrlSetState($Button5, $GUI_DISABLE)
	$aJPG = _FileListToArray($dos, "*.jpg", 1, 1)
	if @error = 4 or @error = 1 Then $file = 0
	   if $file <> 0 then $file = _Min(UBound($aJPG)-1, $file)
;~ 	ConsoleWrite($file&@crlf)
	if $file > 1 Then GUICtrlSetState($Button1, $GUI_ENABLE)
	if $file < UBound($aJPG)-1 then GUICtrlSetState($Button2, $GUI_ENABLE)


	if $file <> 0 Then
;~ 	   GUIDelete($Pic1)
;~ 	   $Pic1 = GUICtrlCreatePic("", 440, 8, 430, 593)
;~ GUICtrlSetImage($Pic1, $aJPG[$file])
		GUICtrlSetState($Pic1, $GUI_SHOW)
		GUICtrlSetImage($Pic1, $aJPG[$file])
		GUICtrlSetState($Button3, $GUI_ENABLE)
		GUICtrlSetState($Button4, $GUI_ENABLE)
		GUICtrlSetState($Button5, $GUI_ENABLE)
	Else
;~ 		GUICtrlSetImage($Pic1, "")
		GUICtrlSetState($Pic1, $GUI_HIDE)
	EndIf



;~ RECONNAISSANCE AUTOMATIQUE DU DOSSIER
GUICtrlSetState($LabelA, $GUI_HIDE)
	if $file <> 0 Then
	   			if $id <> -1 then _GUICtrlListView_SetItemSelected($ListView1, $id, 0, 0)
		$id = -1
		Local $aPathSplit = _PathSplit($aJPG[$file], $sDrive, $sDir, $sFileName, $sExtension)
		if stringlen($aPathSplit[3]) = 13 Then
			$id = _ArraySearch($encours, stringleft($aPathSplit[3], 11), 0, 0, 0, 0, 1, 0)
			if $id <> -1 Then
				GUICtrlSetData($LabelA, $aPathSplit[3])
				GUICtrlSetState($LabelA, $GUI_SHOW)
				_GUICtrlListView_SetItemSelected($ListView1, $id)
				_WinAPI_SetFocus(ControlGetHandle("Attribution via Apicrypt", "", $ListView1))
			EndIf
		EndIf
	EndIf
EndFunc   ;==>_Rafraichir_IMG



Func Down_PDF_CERBA($num)


	$Form11 = GUICreate("Téléchargement du PDF depuis le serveur CERBA", 624, 135, 192, 124)
	$Label11 = GUICtrlCreateLabel("Connection au serveur CERBA...", 8, 8, 192, 45)
	$Progress11 = GUICtrlCreateProgress(8, 64, 601, 25)
	GUISetState(@SW_SHOW)

	GUICtrlSetData($Progress11, 10)
	$wb = _IECreate("https://serveur-resultats.mycerba.com", 0, 1, 1, 1)
sleep(500)
	GUICtrlSetData($Label11, "Connection sécurisée en cours...")
	GUICtrlSetData($Progress11, 20)
	Local $oForm = _IEFormGetCollection($wb, 0)
   _IEFormElementSetValue(_IEFormElementGetCollection($oForm, 0), "")
_IEFormElementSetValue(_IEFormElementGetCollection($oForm, 1), "")
sleep(500)
_IEAction ( _IEFormElementGetCollection($oForm, 4),"click")
;~ _IEFormSubmit($oForm)
;~ _IELinkClickByIndex($oForm,0

;~ 		$identif = _IEFormGetObjByName($wb, "form")
;~ 			$login = _IEFormElementGetObjByName($identif, "username")
;~ 	$mdp = _IEFormElementGetObjByName($identif, "password")
;~ exit
	GUICtrlSetData($Label11, "Accès par dossier...")
	GUICtrlSetData($Progress11, 30)
	_IENavigate($wb, "https://serveur-resultats.mycerba.com/#/cerba/serveur-resultats")
	sleep(1000)

	GUICtrlSetData($Label11, "Recherche dossier " & $num & " en cours...")
	GUICtrlSetData($Progress11, 40)


	Local $oForm = _IEFormGetCollection($wb, 0)
   _IEFormElementSetValue(_IEFormElementGetCollection($oForm, 0), "E0")
_IEFormElementSetValue(_IEFormElementGetCollection($oForm, 1), "E1")
_IEFormElementSetValue(_IEFormElementGetCollection($oForm, 2), "E2")
_IEFormElementSetValue(_IEFormElementGetCollection($oForm, 3), "E3")

exit
	$identif2 = _IEFormGetObjByName($wb, "form1")
	$ref = _IEFormElementGetObjByName($identif2, "ref")
	_IEFormElementSetValue($ref, $num)
	_IEFormSubmit($identif2)

	GUICtrlSetData($Label11, "Recherche du PDF n°" & $num & " en cours...")
	GUICtrlSetData($Progress11, 50)
	$oLinks = _IELinkGetCollection($wb)
	For $oLink In $oLinks
		$sTxt = $oLink.href
		if Stringlen($sTxt) = 78 and StringRight($sTxt, 4) = ".pdf" and StringLeft($sTxt, 50) = "https://www.serveur.lab-cerba.com/pdf/temp/nomfich" then ExitLoop
	Next

	GUICtrlSetData($Label11, "Téléchargement du PDF n°" & $num & " en cours...")
	GUICtrlSetData($Progress11, 60)
;~ 	ConsoleWrite($sTxt)
	InetGet($sTxt, "\\cha-vmsdata1\Echange\Laboratoire\INTERFACES\cerba\PDF\" & $num & ".pdf", 1, 0)
;~ 	_IEQuit($wb)


	GUICtrlSetData($Label11, "Conversion du PDF n°" & $num & " en cours")
	GUICtrlSetData($Progress11, 70)
	DirRemove(@ScriptDir & "\PDF\tmp\", 1)
	DirCreate(@ScriptDir & "\PDF\tmp\")
	DirCreate(@ScriptDir & '\PDF\tmp\' & $num & '\')
	RunWait("pdftopng.exe " & @ScriptDir & "\PDF\" & $num & ".pdf " & @ScriptDir & "\PDF\tmp\" & $num & "\" & $num & " -gray", @ScriptDir, @SW_HIDE)
	$aFiles2 = _FileListToArray(@ScriptDir & "\PDF\tmp\" & $num, "*.png", 1, 1)
	_GDIPlus_Startup()
	if UBound($afiles2) <> 0 Then
		For $y = 1 to $afiles2[0]
			$hImage = _GDIPlus_ImageLoadFromFile($afiles2[$y])
			$sCLSID = _GDIPlus_EncodersGetCLSID("JPG")
			_GDIPlus_ImageSaveToFileEx($hImage, "\\cha-vmsdata1\Echange\Laboratoire\INTERFACES\cerba\IMG\2" & $num & "-" & $y & ".jpg", $sCLSID, DllStructGetPtr($tParams))
			DirRemove(@ScriptDir & "\PDF\" & $num)
		Next
		_GDIPlus_ShutDown()
		FileDelete("\\cha-vmsdata1\Echange\Laboratoire\INTERFACES\cerba\PDF\" & $num & ".pdf")
		$return = 1
		GUICtrlSetData($Progress11, 100)
	Else
		GUICtrlSetData($Progress11, 100)
		MsgBox(16, "Erreur", "Dossier CERBA n°" & $num & " non trouvé. PDF non téléchargé.")
		$return = 0
	EndIf
	GUIDelete($Form11)
	Return $return
EndFunc

Func _Rafraichir($opt = 0)
   SplashTextOn("Chargement de l'interface en cours...", "Chargement de l'interface en cours...", 200, 70, -1, -1, 0, "", 12)
_FileReadToArray(@ScriptDir & "\integration.txt", $arint, 0, "|")
;~ _ArrayDisplay($arint)
	_ArraySort($encours, 1, 0, 0, 0)

	if $opt = 0 then
		$aFFL = _FileListToArrayRec("Y:\SAUVPDF", "*.pdf|echec.pdf|echec", 1, 1)
;~ _ArrayDisplay($aFFL)
		for $y = 1 To UBound($encours) - 2
;~ 		 ConsoleWrite($y&"="&"Y:\scanner\"&Stringright($encours[$y][0],10)&"-0.jpg : "&_ArraySearch($ar,"Y:\scanner\"&Stringright($encours[$y][0],10)&"-0.jpg")&@crlf)
;~ 		 ConsoleWrite($y&@crlf)
			if $y >= UBound($encours) then ExitLoop
			if _ArraySearch($aFFL, Stringleft($encours[$y][0], 11), 0, 0, 0, 1) > 0 Then
;~ 		 ConsoleWrite("DEL $y="&$y&" : "&Stringleft($encours[$y][0],11)&@crlf)
				_ArrayDelete($encours, $y)
				$y = $y - 1
				if $y >= UBound($encours) - 4 then ExitLoop
			Elseif FileExists("Y:\scanner\" & Stringright($encours[$y][0], 10) & "-0.jpg") Then
;~ 			ConsoleWrite("DEL $y="&$y&" : "&Stringright($encours[$y][0],10)&" ???"&@crlf)
;~ ConsoleWrite("Y:\scanner\"&Stringright($encours[$y][0],10))
if _ArraySearch($arint,"Y:\scanner\"&Stringright($encours[$y][0],10),0,0,0,1,1)=-1 Then
				if _ArraySearch($ar, "Y:\scanner\" & Stringright($encours[$y][0], 10) & "-0.jpg") = -1 Then
;~ 		 		 ConsoleWrite("DEL $y="&$y&" : "&Stringright($encours[$y][0],10)&@crlf)
					_ArrayDelete($encours, $y)
					$y = $y - 1
					if $y >= UBound($encours) - 4 then ExitLoop
EndIf
				EndIf
			endif
		Next
	endif
;~  exit
	_FileWriteFromArray(@ScriptDir & "\encoursCERBA.txt", $encours)
	_ArraySort($encours, 0, 0, 0, 0)
	_GUICtrlListView_DeleteAllItems($ListView1)

	for $y = 0 To UBound($encours) - 2

		GUICtrlCreateListViewItem(StringLeft($encours[$y][0], 7) & " " & StringRight($encours[$y][0], 4) & "|" & $encours[$y][1] & "|" & $encours[$y][2] & "|" & $encours[$y][3] & "|" & $encours[$y][5] & "|" & $encours[$y][4] & "|" & $encours[$y][6] & "|" & $encours[$y][7], $ListView1)
		_GUICtrlListView_SetColumnWidth($ListView1, 0, $LVSCW_AUTOSIZE)
		GUICtrlSetBkColor(-1, 0xD3D3D3)
		GUICtrlSetColor(-1, "0x000000")
		_GUICtrlListView_SetItemImage($ListView1, $y, 2)
	Next
	_GUICtrlListView_HideColumn($ListView1, 1)
	_GUICtrlListView_HideColumn($ListView1, 6)
	_GUICtrlListView_HideColumn($ListView1, 7)
	SplashOff()
EndFunc   ;==>_Rafraichir

Func _IsChecked($idControlID)
	Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked



