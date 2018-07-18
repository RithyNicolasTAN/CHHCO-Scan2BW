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
#include "ExtMsgBox.au3"
#include <Misc.au3>
If _Singleton("verif_apicrypt", 1) = 0 Then Exit

;~ CONFIGURATION
$apitunnel="Y:\Apitunnel"
$scan="Y:\SCAN\"
$scan2="Y:\SCAN2\"

$aFiles=_FileListToArray ($apitunnel,"*.pdf",1,1)
$aFiles2=_FileListToArray ($scan,"*.jpg",1,1)
$aFiles3=_FileListToArray ($scan2,"*.jpg",1,1)

$Form1 = GUICreate("FICHIERS EN ATTENTE", 366, 56, @DesktopWidth-366, 18,$WS_POPUP, $WS_EX_TOPMOST)
$Label1 = GUICtrlCreateLabel("FICHIERS APICRYPT EN ATTENTE : "&UBound($afiles)-1, 8, 8, 350, 20,$SS_CENTER)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
$Label2 = GUICtrlCreateLabel("FICHIERS A ATTRIBUER EN ATTENTE : "&UBound($afiles2)-1, 8, 28, 350, 20,$SS_CENTER)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
GUISetBkColor($COLOR_RED, $Form1)

$hTimer = TimerInit()
$hTimer2 = TimerInit()
While 1
   ; TIMER REFRESH NB DOSSIERS APICRYPT
   if TimerDiff($hTimer) >= 60000 Then
			   $hTimer = TimerInit()
      $aFiles=_FileListToArray ($apitunnel,"*.pdf",1,1)
	  $aFiles2=_FileListToArray ($scan,"*.jpg",1,1)
	  $aFiles3=_FileListToArray ($scan2,"*.jpg",1,1)
	  GUISetState(@SW_HIDE,$Form1)
   EndIf

; TIMER CLIGNOTEMENT
		 	if TimerDiff($hTimer2) >= 10000 Then
			   $hTimer2 = TimerInit()
			   GUICtrlSetData($Label1,"")
			   GUICtrlSetData($Label2,"")
			   $beep=False
			   if UBound($aFILES)>0 Then
;~ 			GUISetState(@SW_SHOW)
			GUICtrlSetData($Label1,"FICHIERS APICRYPT EN ATTENTE : "&UBound($afiles)-1)
;~ 			ConsoleWrite("FICHIERS APICRYPT EN ATTENTE : "&UBound($afiles)-1&@crlf)
			$beep=True
		 EndIf
		 if UBound($aFILES2)+UBound($afiles3)>0 Then
;~ 			GUISetState(@SW_SHOW)
			GUICtrlSetData($Label2,"FICHIERS A ATTRIBUER EN ATTENTE : "&UBound($afiles2)+UBound($afiles3)-1)
;~ 			Consolewrite("FICHIERS A ATTRIBUER EN ATTENTE : "&UBound($afiles2)+UBound($afiles3)-1&@crlf)
			$beep=True
		 EndIf
		 if $beep=True then GUISetState(@SW_SHOW,$Form1)
		 if $beep AND WinExists ("Intégration")=False  AND WinExists ("Attribution via Scanner")=False  AND WinExists ("Attribution via Apicrypt")=False Then

			beep(500,300)
			endif

EndIf
;~  	$nMsg = GUIGetMsg()
;~ 	Switch $nMsg
;~ 		Case $GUI_EVENT_CLOSE
;~ 			exit
;~ 		 EndSwitch

sleep(1000)
WEnd
