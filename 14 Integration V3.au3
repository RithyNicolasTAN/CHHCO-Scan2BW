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

If _Singleton("integration", 1) = 0 Then Exit

$handle = WinGetHandle("biowin")

;~ ;LECTURE FICHIER SACNNERS (DOUBLONS)
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

;~ $aJPG = _FileListToArray("Y:\SCANNER\", "*.jpg", 1)
;~   _ArrayDisplay($aSCAN)
;~ for $t=1 to UBound($aJPG)-1
;~    ConsoleWrite($aJPG[$t]&@crlf)
;~    $iIndex = _ArrayFindAll($aSCAN, $aJPG[$t], 0, 0, 0, 3, 0, 0)
;~    if ubound($iIndex) > 0 Then
;~ 	  _ArrayDisplay($iindex)
;~ 	  _ArrayDisplay($aJPG)
;~ 	  _ArrayDelete ( $aJPG, $t)
;~ 	  _ArrayDisplay($aJPG)
;~ 	  EndIf
;~    Next
;~    _ArrayDisplay($aJPG)


;~ ;LECTURE FICHIER A FAIRE

Global $ar[1]



_FileReadToArray(@ScriptDir & "\integration.txt", $ar, 0, "|")
$ar = _2d_Arr_Uniq($ar)
_FileWriteFromArray(@ScriptDir & "\integration.txt", $ar, 0)

;~ Local $hFileOpen = FileOpen(@ScriptDir & "\integration.txt", $FO_READ)
;~ while 1=1
;~ $txt = FileReadLine($hFileOpen) ; Lecture de la ligne du fichier source
;~ if @error = -1 then ExitLoop ; Si on est à la fin, on quitte la bouche
;~    _ArrayAdd($ar,$txt)
;~  WEnd
;~ FileClose($hFileOpen)
;~ _ArrayDisplay($ar)

Global $encours, $adelai
_FileReadToArray(@ScriptDir & "\encoursCERBA.txt", $encours, 0, "|")

#Region ### START Koda GUI section ### Form=C:\Users\tani01\Downloads\Koda\Forms\Form1.kxf
$Form1 = GUICreate("Intégration", 274, 238, 0, 0, -1, BitOR($WS_EX_TOPMOST, $WS_EX_WINDOWEDGE))
$Group1 = GUICtrlCreateGroup("Patient en cours : ", 0, 0, 265, 78)
$Label1 = GUICtrlCreateLabel("Nom : ", 8, 14, 249, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$Label2 = GUICtrlCreateLabel("Prénom : ", 8, 36, 250, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
$Label3 = GUICtrlCreateLabel("Date de naissance : ", 8, 58, 250, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
GUICtrlCreateGroup("", -99, -99, 1, 1)
;~ $Progress1 = GUICtrlCreateProgress(0, 96, 273, 9)
$Button1 = GUICtrlCreateButton("Saisir patient", 0, 80, 137, 33)
$Button7 = GUICtrlCreateButton("Saisir dossier", 136, 80, 129, 33)
$Button2 = GUICtrlCreateButton("Retirer BLOCC", 0, 112, 265, 33)
;~ $Button8 = GUICtrlCreateButton("Date retour avec VF", 136, 112, 129, 33)
$Button3 = GUICtrlCreateButton("Intégration scan", 0, 144, 265, 33)
$Button4 = GUICtrlCreateButton("Dossier suivant", 0, 176, 265, 33)
$Button5 = GUICtrlCreateButton("Afficher image", 0, 216, 137, 17)
$Button6 = GUICtrlCreateButton("Supprimer le dossier", 136, 216, 129, 17)

#EndRegion ### END Koda GUI section ###

GUISetState(@SW_SHOW)

for $i = 0 to UBound($ar, 1) - 1
	if $i >= UBound($ar, 1) then ExitLoop
;~    GUICtrlSetState($Button2, $GUI_DISABLE)
	GUICtrlSetState($Button3, $GUI_DISABLE)
	GUICtrlSetState($Button4, $GUI_DISABLE)
;~ GUICtrlSetData($Button2, "Date retour sans VF")

	GUICtrlSetData($Label3, "Dossier : 2" & StringMid($ar[$i][0], 12, 12))
	$ind = _ArraySearch($encours, "2" & StringMid($ar[$i][0], 12, 10), 0, 0, 0, 1, 0, 0)
	if $ind > 0 then

		GUICtrlSetData($Label1, "Nom : " & $encours[$ind][2])
		GUICtrlSetData($Label2, "Prénom : " & $encours[$ind][3])

		if $ar[$i][1] = "S" or $ar[$i][1] = "F" or $ar[$i][1] = "B" then GUICtrlSetState($Button3, $GUI_ENABLE)
		if $ar[$i][1] = "R" or $ar[$i][1] = "A" or $ar[$i][1] = "S" then GUICtrlSetState($Button2, $GUI_ENABLE)
;~  if $ar[$i][1]="A" then GUICtrlSetData($Button2, "Date de retour par etape 4")

		If $ar[$i][1] = "F" then GUISetBkColor($COLOR_RED, $Form1)
		if $ar[$i][1] = "S" or $ar[$i][1] = "R" or $ar[$i][1] = "A" or $ar[$i][1] = "B" then GUISetBkColor($COLOR_AQUA, $Form1)

		if $i < UBound($ar, 1) - 1 then GUICtrlSetState($Button4, $GUI_ENABLE)
		if StringInStr(WinGetTitle("[active]"), "biowin") = 0 Then
			WinActivate("biowin")
		EndIf
		GUISetState(@SW_SHOW)
		While 1
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $GUI_EVENT_CLOSE
					exitloop(2)
				Case $Button1 ;SAISIE PATIENT
					if StringInStr(WinGetTitle("[active]"), "biowin") = 0 Then
						WinActivate("biowin")
					EndIf
					ControlSend($handle, "", "", "{F2}")
;~  Send ("{F2}")
					Send($encours[$ind][2])
					Send("{F2}")
					Send("{F2}")
					Send($encours[$ind][3])
					Send("{ENTER}")
					sleep(200)


				Case $Button7 ;SAISIE DOSSIER
					if StringInStr(WinGetTitle("[active]"), "biowin") = 0 Then
						WinActivate("biowin")
					EndIf
					ControlSend($handle, "", "", "2" & StringMid($ar[$i][0], 12, 6) & "-" & StringMid($ar[$i][0], 18, 4) & "{ENTER}")

;~ 					 Case $Button8 ;SAISIE DATE DE RETOUR AVEC VF
;~ 			if StringInStr (WinGetTitle("[active]"),"biowin")=0 Then
;~    WinActivate("biowin")
;~ EndIf
;~ ControlSend($handle, "", "", "{F4}")
;~ 	sleep(200)
;~ 	Send("{DOWN}")
;~ 			Send("{DOWN}")
;~ 			Send("{DOWN}")
;~ Send("="&stringmid(_nowcalc(),9,2)&"/"&stringmid(_nowcalc(),6,2)&"/"&stringmid(_nowcalc(),3,2))
;~ 	Send("{ENTER}")
;~ 		Send("{F1}")
;~ 		Send("{ESC}")
;~ 	sleep(200)

;~ if $ar[$i][1]="R" Then ; RETOUR UNIQUEMENT, PASSAGE DOSSIER SUIVANT
;~    Send("{F1}")
;~    if UBound($ar,1)>1 Then
;~ 	   _ArrayDelete($ar,$i)
;~ _FileWriteFromArray(@ScriptDir&"\integration.txt", $ar, 0)
;~ Else
;~    _ArrayDelete($ar,$i)
;~    $hFO = FileOpen(@ScriptDir&"\integration.txt", $FO_OVERWRITE)
;~     FileWriteLine($hFO, "")
;~ 	 FileClose($hFO)
;~ 	   EndIf
;~ 					$i=$i-2
;~ 					if $i=-2 Then $i=-1
;~ 					   ExitLoop
;~ Send("{F1}")
;~    EndIf


				Case $Button2 ;SAISIE DATE DE RETOUR SANS VF
					if StringInStr(WinGetTitle("[active]"), "biowin") = 0 Then WinActivate("biowin")
					Send("{F4}")
					sleep(100)
					Send("{F5}")
					Send("BLOCC")
					Send("{ENTER}")
					sleep(100)
					Send("{ENTER}")
					Send("{ESC}")

				Case $Button3 ;SCANNER
					FileDelete("Y:\SCANNER\SCAN99\" & stringMid($ar[$i][0], 12, 12) & ".dem")
					FileDelete("Y:\SCANNER\SCAN99\" & stringMid($ar[$i][0], 12, 12) & ".ok")
					if StringInStr(WinGetTitle("[active]"), "biowin") = 0 Then
						WinActivate("biowin")
					EndIf

					ControlSend($handle, "", "", "{F6}")
					sleep(100)
					Send("{ENTER}")
					sleep(100)
;~ 			Send("{ENTER}")
;~ 		sleep(100)
					Send("{F7}")
					sleep(100)
					Send("NNNNNNNO")
					Send("{ENTER}")
					sleep(100)
					Send("N{ENTER}O{ENTER}N{ENTER}")
					sleep(100)
					Send("{F1}")

					$Form3 = GUICreate("Attente de la réponse de BioWin", 391, 89, 302, 218, $WS_POPUP, $WS_EX_TOPMOST)
					$Label31 = GUICtrlCreateLabel("Attente de la réponse de BioWin", 64, 8, 264, 24)
					GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
					GUICtrlSetColor(-1, 0xFF0000)
					$Button31 = GUICtrlCreateButton("Annuler", 88, 40, 185, 33)
					GUISetState(@SW_SHOW)
					While 1


						$aDEM = _FileListToArray("Y:\SCANNER\SCAN99\", StringMid($ar[$i][0], 12, 12) & ".dem", 1)

						if UBound($aDEM) > 1 Then ;FICHIER DEM CREE PAR BIOWIN
							sleep(500)
							$ret99 = FileMove("Y:\SCANNER\SCAN99\" & stringMid($ar[$i][0], 12, 12) & ".dem", "Y:\SCANNER\SCAN99\" & StringMid($ar[$i][0], 12, 12) & ".ok")
							sleep(500)
							if $ret99 = 0 then
								msgbox(0, "", "Supresion du fichier *.dem impossible. Fichier non intégré")
								exitloop(2)
							EndIf
							GUIDelete($Form3)
							ExitLoop
						EndIf

						$nMsg = GUIGetMsg()
						Switch $nMsg

							Case $Button31
								msgbox(0, "", "Ficher *.dem non trouvé. Fichier non intégré")
								exitloop(3)
						EndSwitch
					WEnd
					sleep(1700)
					if MsgBox(4, "Scannerisation,", "La scannerisation est-elle reussie ?") = 6 Then ;SCAN OK
						if UBound($ar, 1) > 1 Then
							_ArrayDelete($ar, $i)
							_FileWriteFromArray(@ScriptDir & "\integration.txt", $ar, 0)
						Else
							_ArrayDelete($ar, $i)
							$hFO = FileOpen(@ScriptDir & "\integration.txt", $FO_OVERWRITE)
							FileWriteLine($hFO, "")
							FileClose($hFO)
						EndIf
;~ _ArrayDisplay($ar)

						if StringInStr(WinGetTitle("[active]"), "biowin") = 0 then WinActivate("biowin")
						Send("{ENTER}")
						Send("{ESC}")
						$i = $i - 1
;~ if $i<UBound($ar,1)-2 Then

;~    $ind2=_ArraySearch($encours, "2"&StringMid($ar[$i][0],12,10), 0, 0, 0, 1, 0, 0)


;~    If $GUICtrlRead($label1)=$encours[$ind2][2] and GUICtrlRead ($label2)=$encours[$ind2][3] Then


;~ 	  Send("{F1}")
;~  	  EndIf
;~    EndIf


					Else
						Exit
					EndIf

					exitloop





					Send("=" & stringmid(_nowcalc(), 9, 2) & "/" & stringmid(_nowcalc(), 6, 2) & "/" & stringmid(_nowcalc(), 3, 2))
					Send("{ENTER}")
					Send("{F1}")
					Send("{ESC}")
					sleep(200)

					if $ar[$i][1] = "R" Then ; RETOUR UNIQUEMENT, PASSAGE DOSSIER SUIVANT
						Send("{F1}")
					EndIf
					if UBound($ar, 1) > 1 Then
						_ArrayDelete($ar, $i)
						_FileWriteFromArray(@ScriptDir & "\integration.txt", $ar, 0)
					Else
						_ArrayDelete($ar, $i)
						$hFO = FileOpen(@ScriptDir & "\integration.txt", $FO_OVERWRITE)
						FileWriteLine($hFO, "")
						FileClose($hFO)
					EndIf

					ExitLoop





				Case $Button4 ;DOSSIER SUIVANT
					ExitLoop
				Case $Button5 ;AFFICHER IMAGE

					$hGUI = GUICreate("", @DesktopWidth, @DesktopHeight, -1, -1, $WS_POPUP, $WS_EX_TOPMOST)
					$Pic2 = GUICtrlCreatePic($ar[$i][0], 0, 0, @DesktopWidth, @DesktopHeight)
					GUISetState(@SW_SHOW, $hGUI)
					While 1
						if GUIGetMsg() = $Pic2 Then ExitLoop
					WEnd
					GUIDelete($hGUI)

				Case $Button6 ;SUPPRIMER DOSSIER

					$hGUI = GUICreate("", @DesktopWidth, @DesktopHeight, -1, -1, $WS_POPUP)
					$Pic2 = GUICtrlCreatePic($ar[$i][0], 0, 0, @DesktopWidth, @DesktopHeight)
					GUISetState(@SW_SHOW, $hGUI)
					if MsgBox(1, "Supprimer ?", "Etes-vous sur de vouloir suprimer cet élément ?") = 1 Then
						FileRecycle($ar[$i][0])
						if UBound($ar, 1) > 1 Then
							_ArrayDelete($ar, $i)
							_FileWriteFromArray(@ScriptDir & "\integration.txt", $ar, 0)
						Else
							_ArrayDelete($ar, $i)
							$hFO = FileOpen(@ScriptDir & "\integration.txt", $FO_OVERWRITE)
							FileWriteLine($hFO, "")
							FileClose($hFO)
						EndIf
						GUIDelete($hGUI)
						$i = $i - 1
						ExitLoop
					EndIf

			EndSwitch
		WEnd

	Else
		$hGUI = GUICreate("", @DesktopWidth, @DesktopHeight, -1, -1, $WS_POPUP)
		$Pic2 = GUICtrlCreatePic($ar[$i][0], 0, 0, @DesktopWidth, @DesktopHeight)
		GUISetState(@SW_SHOW, $hGUI)
		if MsgBox(4, "Dossier non trouvé", "Dossier n°" & StringMid($ar[$i][0], 12, 10) & " non trouvé !" & @crlf & "Voulez-vous le supprimer ?" & @crlf & "Si non, passage au dossier suivant") = 6 Then

			FileDelete($ar[$i][0])
			if UBound($ar, 1) > 1 Then
				FileDelete($ar[$i][0])
				_ArrayDelete($ar, $i)
				_FileWriteFromArray(@ScriptDir & "\integration.txt", $ar, 0)
			Else
				FileDelete($ar[$i][0])
				_ArrayDelete($ar, $i)
				$hFO = FileOpen(@ScriptDir & "\integration.txt", $FO_OVERWRITE)
				FileWriteLine($hFO, "")
				FileClose($hFO)
			EndIf
			GUIDelete($hGUI)
			$i = $i - 1
		EndIf

	endif
Next
_FileWriteFromArray(@ScriptDir & "\integration.txt", $ar, 0)



Func _2d_Arr_Uniq($aArray)
	If UBound($aArray, 0) <> 2 Then Return SetError(1, 0, 0)

	Local $iBoundY = UBound($aArray), $iBoundX = UBound($aArray, 2)

	Local $aTemp[$iBoundY]
	For $y = 0 To $iBoundY - 1
		For $x = 0 To $iBoundX - 1
			$aTemp[$y] &= $aArray[$y][$x] & "|"
		Next
		$aTemp[$y] = StringTrimRight($aTemp[$y], 1)
	Next

	$aTemp = _ArrayUnique($aTemp)

	ReDim $aArray[$aTemp[0]][$iBoundX]
	For $y = 1 To $aTemp[0]
		$aRow = StringSplit($aTemp[$y], "|", 2)
		For $x = 0 To $iBoundX - 1
			$aArray[$y - 1][$x] = $aRow[$x]
		Next
	Next

	Return $aArray
EndFunc   ;==>_2d_Arr_Uniq
